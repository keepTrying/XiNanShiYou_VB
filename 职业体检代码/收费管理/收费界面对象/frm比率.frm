VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm���� 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "��������"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   5925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmd���� 
      Caption         =   "����"
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox txt���� 
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   3120
      Width           =   1215
   End
   Begin MSComctlLib.TreeView ctrwItem 
      Height          =   2895
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   5106
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Label lbl�ٷֱ� 
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   3
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label lbl���� 
      Caption         =   "���ñ���Ϊ��"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   3240
      Width           =   1095
   End
End
Attribute VB_Name = "frm����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public pstr�շ���Ŀ��� As String

'���ܣ� �����õ��շѿ��ҷ��������Ϣ���浽���ݿ���
'�޸ģ����� 2006/01/23
Private Sub cmd����_Click()
 On Error GoTo errHandler
 If Left(ctrwItem.SelectedItem.Key, 1) = "k" Then
    Dim mstr�������� As Recordset
    Set mstr�������� = dafuncGetData("select * from ϵͳ����_�����ֵ�� where ����= '" & Left(Trim(ctrwItem.SelectedItem.Text), (InStr(Trim(ctrwItem.SelectedItem.Text), "(") - 1)) & "'")
    sub�޸ı��� pstr�շ���Ŀ���, mstr��������("���")

 End If
 ctrwItem.Nodes.Clear
 Form_Load
 Exit Sub
errHandler:
    sfsub������ "�շѽ���", "frm����", "cmd����_Click", Err.Number, Err.Description, True
End Sub

Private Sub ctrwItem_NodeClick(ByVal Node As MSComctlLib.Node)
  On Error GoTo errHandler
  If Left(ctrwItem.SelectedItem.Key, 1) = "k" Then
    txt����.SetFocus
    cmd����.Enabled = True
    Dim mstr���ұ�� As Recordset
'    MsgBox Left(Trim(Node.Text), (InStr(Trim(Node.Text), "(") - 1))
    Set mstr���ұ�� = dafuncGetData("select * from ϵͳ����_�����ֵ�� where ����= '" & Left(Trim(Node.Text), (InStr(Trim(Node.Text), "(") - 1)) & "'")
    Dim lobjRec As Recordset
    Set lobjRec = func��ѯ������Ŀ1(pstr�շ���Ŀ���, mstr���ұ��("���"))
    txt����.Text = IIf(IsNull(lobjRec("����")), "", lobjRec("����")) 'lobjRec("����")'left(Trim(Node.Text),(instr(Trim(Node.Text)-1),"("))
  Else
    txt����.Text = ""
    cmd����.Enabled = False
  End If
  Exit Sub
errHandler:
    sfsub������ "�շѽ���", "frm����", "ctrwItem_NodeClick", Err.Number, Err.Description, True
End Sub

Private Sub Form_Load()
  On Error GoTo errHandler
  Dim lobjRec As Recordset
  Dim mstr�շ���Ŀ���� As Recordset
  Dim mstr�������� As Recordset
  Dim lobjNode As Node
  cmd����.Enabled = False
  Set mstr�շ���Ŀ���� = dafuncGetData("select * from �շѹ���_�շ���Ŀ�ֵ�� where �շ���Ŀ���= '" & pstr�շ���Ŀ��� & "'")
  Set lobjNode = ctrwItem.Nodes.Add(, , "s" & pstr�շ���Ŀ���, mstr�շ���Ŀ����("�շ���Ŀ����"))
  Set lobjRec = func��ѯ������Ŀ("'" & pstr�շ���Ŀ��� & "'")
  Do While Not lobjRec.EOF
     Set mstr�������� = dafuncGetData("select * from ϵͳ����_�����ֵ�� where ���= '" & lobjRec("���ұ��") & "'")
     Set lobjNode = ctrwItem.Nodes.Add("s" & pstr�շ���Ŀ���, tvwChild, "k" & mstr��������("����") & "s" & pstr�շ���Ŀ���, mstr��������("����") & "(" & lobjRec("����") & "%" & ")")
     lobjNode.EnsureVisible

     lobjRec.MoveNext
  Loop
Exit Sub
errHandler:
    sfsub������ "�շѽ���", "frm����", "Form_Load", Err.Number, Err.Description, True
End Sub

Public Sub sub�޸ı���(para�շ���Ŀ���, para���ұ�� As String)
    On Error GoTo errHandler
    Dim dec As Variant
    Dim dec1 As Variant
    Dim i As Integer
    Dim lobjRec As Object '�����
    Set lobjRec = dafuncGetData("select ���� from �շѹ���_���ұ��ʱ� where �շ���Ŀ���='" + para�շ���Ŀ��� + "'")
    dec = 0
    If lobjRec.RecordCount > 0 Then
       Do While Not lobjRec.EOF

          If IsNull(lobjRec(0)) Then
            dec1 = 0
          Else
            dec1 = lobjRec(0)
          End If
          dec = dec + dec1
          lobjRec.MoveNext
       Loop
    End If

    Set lobjRec = dafuncGetData("select ���� from �շѹ���_���ұ��ʱ� where �շ���Ŀ���='" + para�շ���Ŀ��� + "'and ���ұ��='" + para���ұ�� + "'")
    If lobjRec.RecordCount > 0 Then
       If IsNull(lobjRec(0)) Then
          dec1 = 0
       Else
          dec1 = lobjRec(0)
       End If
          dec = dec - dec1
          dec = dec + Val(Trim(txt����.Text))

    Else
       dec = dec + dec(Trim(txt����.Text))
    End If

    If dec > 100 Then
       MsgBox "�ܱ����ѳ���100%�����������á�", vbOKOnly, "ϵͳ��ʾ"
       txt����.Text = ""
       txt����.SetFocus
       Exit Sub
    End If
    
    dafuncGetData ("update �շѹ���_���ұ��ʱ� set ����='" & Val(Trim(txt����.Text)) & "' where �շ���Ŀ���='" + para�շ���Ŀ��� + "' and ���ұ��='" + para���ұ�� + "'")
    'MsgBox "����ɹ�", vbOKOnly, "ϵͳ��ʾ"
    Exit Sub
errHandler:
    sfsub������ "�շѽ���", "frm����", "sub�޸ı���", Err.Number, Err.Description, True
End Sub
'func��ѯ������Ŀ
'����:���ݲ�ѯ������ѯ���������ļ�¼
Public Function func��ѯ������Ŀ1(ByVal para�շ���Ŀ���, para���ұ�� As String) As Object
    On Error GoTo errHandler
    If para��ѯ���� = "ALL" Then
        Set func��ѯ������Ŀ1 = dafuncGetData("select * from  �շѹ���_���ұ��ʱ�")
    Else
        Set func��ѯ������Ŀ1 = dafuncGetData("select * from �շѹ���_���ұ��ʱ� where �շ���Ŀ���='" + para�շ���Ŀ��� + "' and ���ұ��='" + para���ұ�� + "'")
    End If
    
    Exit Function
errHandler:
    sfsub������ "�շѽ���", "frm����", "func��ѯ������Ŀ", Err.Number, Err.Description, True
End Function
'func��ѯ������Ŀ
'����:���ݲ�ѯ������ѯ���������ļ�¼
Public Function func��ѯ������Ŀ(ByVal para��ѯ���� As String) As Object
    On Error GoTo errHandler
    If para��ѯ���� = "ALL" Then
        Set func��ѯ������Ŀ = dafuncGetData("select * from  �շѹ���_���ұ��ʱ�")
    Else
        Set func��ѯ������Ŀ = dafuncGetData("select * from �շѹ���_���ұ��ʱ� where �շ���Ŀ���=" + para��ѯ����)
    End If
    
    Exit Function
errHandler:
    sfsub������ "�շѽ���", "frm����", "func��ѯ������Ŀ1", Err.Number, Err.Description, True
End Function

Private Sub txt����_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    KeyAscii = gfuncKeyNum(KeyAscii)
    If KeyAscii = 13 Then
        cmd����_Click
    End If
End Sub
