VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSetDoctor 
   Caption         =   "ҽʦ�����Ŀ����"
   ClientHeight    =   8055
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   11295
   Icon            =   "frmSetDoctor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8055
   ScaleWidth      =   11295
   Begin VB.CommandButton ccmdExit 
      Caption         =   "����(&X)"
      Height          =   375
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6840
      Width           =   1275
   End
   Begin MSComctlLib.TreeView ctrwAllItem 
      Height          =   5925
      Left            =   6480
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   480
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   10451
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
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
   Begin VB.CommandButton ccmdDel 
      Caption         =   ">> ȥ��"
      Enabled         =   0   'False
      Height          =   435
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton ccmdAdd 
      Caption         =   "<< ���"
      Enabled         =   0   'False
      Height          =   435
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   960
      Width           =   855
   End
   Begin MSComctlLib.TreeView ctrwDoctor 
      Height          =   5895
      Left            =   45
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   450
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   10398
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      BorderStyle     =   1
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
   Begin MSComctlLib.TreeView ctrwSelectedItem 
      Height          =   5895
      Left            =   2280
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   480
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   10398
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
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
      Caption         =   "˵�������ø�ҽʦ��������Ŀ��ҽʦ����д���������ʱ��ֻ�ܿ�������д���������Ŀ��"
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   6840
      Width           =   7740
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���п�ѡ�����Ŀ(˫������������)��"
      Height          =   180
      Left            =   6480
      TabIndex        =   6
      Top             =   240
      Width           =   3240
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���������Ŀ�б�"
      Height          =   180
      Left            =   2280
      TabIndex        =   2
      Top             =   195
      Width           =   1620
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���ҽʦ�б�"
      Height          =   180
      Index           =   1
      Left            =   90
      TabIndex        =   0
      Top             =   195
      Width           =   1260
   End
End
Attribute VB_Name = "frmSetDoctor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'���ߣ��

Private mobj���ҽʦ  As Object 'ClsMedicalExamer,�������ӡ�ɾ�����ҽʦ�����������Ŀ��

Private Sub ctrwSelectedItem_DblClick()
    On Error Resume Next
    'Ҷ�ڵ����˫��ɾ����
    If ccmdDel.Enabled = True Then
        If Not ctrwSelectedItem.Parent Is Nothing Then
            If Not ctrwSelectedItem.Parent.Parent Is Nothing Then
                ccmdDel_Click
            End If
        End If
    End If
End Sub

Private Sub ctrwSelectedItem_NodeClick(ByVal Node As MSComctlLib.Node)
    ccmdDel.Enabled = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If (Shift & vbAltMask) = vbAltMask And KeyCode = vbKeyX Then
        Unload Me
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
    Dim lobj�����Ŀ�� As Object 'clsTestItemSet��
    Dim lobjRec As Object        '��ȡ�������û���¼�����������ֵ���Ŀ����
    Dim lobjItem As Object       'ĳ����������Ŀ��¼����
    
    On Error GoTo errHandler
    '���á��û�����.umfunc��ȡ�����û�����ȡ�����û��б�
    Set lobjRec = pobjDict.FetchEx("Ա���ֵ�")
    
    '��ʾ��ctrvDoctor�У����нڵ��key=�û���ţ���
    '��ʾ�û���ҽʦ���С�
    ctrwDoctor.Nodes.Add , , "R", "���ҽʦ"
    Do While Not lobjRec.EOF
        '�޸ģ�2001-11-7���������ʾ0000��gues��
        If lobjRec("���") <> "0000" And lobjRec("���") <> "gues" Then
            ctrwDoctor.Nodes.Add "R", tvwChild, "I" & lobjRec("���"), lobjRec("���") & " " & lobjRec("����")
        End If
        lobjRec.movenext
    Loop
    If ctrwDoctor.Nodes.Count > 0 Then
        ctrwDoctor.Nodes(1).Expanded = True
    End If
    
    '���������Ŀ������
    Set lobj�����Ŀ�� = CreateObject("ְҵ������.clsTestItemSet")
    
    '��ȡ���������ࡢ�����Ŀ��
    Set lobjRec = pobjDict.Fetch("ְҵ���������ֵ�")
    
    '��ʾ��������ctrvItem�У����нڵ��key=������id����
    ctrwAllItem.Nodes.Add , , "R", "������"
    Do While Not lobjRec.EOF
        
        'ͨ��"lobj�����Ŀ��"���λ�ȡ������������Ŀ��
        lobj�����Ŀ��.������ = lobjRec("InnerID")
        Set lobjItem = lobj�����Ŀ��.�����Ŀ
        If Not lobjItem.EOF Then
            ctrwAllItem.Nodes.Add "R", tvwChild, "I" & lobjRec("InnerID"), lobjRec("���") & " " & lobjRec("����")
        End If
        '��ʾ�����Ŀ��ctrvItem�У����нڵ��key=����,parent=�����ࣩ��
        Do While Not lobjItem.EOF
            ctrwAllItem.Nodes.Add "I" & lobjRec("InnerID"), tvwChild, "I" & lobjItem("����"), lobjItem("����") & " " & lobjItem("����")
            lobjItem.movenext
        Loop
        
        lobjRec.movenext
    Loop
    
    '��������"mobj���ҽʦ"��
    Set mobj���ҽʦ = CreateObject("ְҵ������.clsMedicalExaminer")
    
    On Error Resume Next
    If ctrwAllItem.Nodes.Count > 0 Then
        ctrwAllItem.Nodes(1).Expanded = True
    End If
    
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����ý���", "frmSetDoctor", "Form_load", 6666, lstrError, False
End Sub

Private Sub ctrwAllItem_DblClick()
    On Error Resume Next
    'Ҷ�ڵ����˫��ɾ����
    If ccmdAdd.Enabled = True Then
        If Not ctrwAllItem.Parent Is Nothing Then
            If Not ctrwAllItem.Parent.Parent Is Nothing Then
                ccmdAdd_Click
            End If
        End If
    End If
End Sub


Private Sub ccmdAdd_Click()
    Dim lstr���� As String   '�����Ŀ���롣
    Dim lobjNode As Node
    Dim i As Long
    
    On Error GoTo errHandler
    
    MousePointer = 11
    Me.Enabled = False
    
    If ctrwAllItem.SelectedItem.Children = 0 Then
        '������Ŀ��
        lstr���� = ctrwAllItem.SelectedItem.Key
        lstr���� = Right(lstr����, Len(lstr����) - 1)
        If lstr���� = "" Then Exit Sub
        
        '�жϸ���Ŀ�Ƿ����ڡ�
        If Not mobj���ҽʦ.func�Ƿ������Ŀ(lstr����) Then
            '�ӿ�����Ӹ���Ŀ���á�
            mobj���ҽʦ.Sub��������Ŀ lstr����
        
            '�Ѵ�����롣
            On Error Resume Next
            If ctrwSelectedItem.Nodes.Count = 0 Then
                ctrwSelectedItem.Nodes.Add , , "R", "������"
            End If
            ctrwSelectedItem.Nodes.Add "R", tvwChild, ctrwAllItem.SelectedItem.Parent.Key, ctrwAllItem.SelectedItem.Parent.Text
            '���������Ŀ�����нڵ��key=����,parent=�����ࣩ��
            ctrwSelectedItem.Nodes.Add ctrwAllItem.SelectedItem.Parent.Key, tvwChild, ctrwAllItem.SelectedItem.Key, ctrwAllItem.SelectedItem.Text
            
        End If
        
    Else '������Ŀ�������������
        
        If ctrwAllItem.SelectedItem.Parent Is Nothing Then
            '���������롣
            If ctrwAllItem.SelectedItem.Children > 0 Then
                '������ڵ㡣
                On Error Resume Next
                ctrwSelectedItem.Nodes.Add , , "R", "������"
                
                On Error GoTo errHandler
                '���μ�����ࡣ
                Set lobjNode = ctrwAllItem.SelectedItem.Child
                For i = 1 To ctrwAllItem.SelectedItem.Children
                    sub��Ӵ��� lobjNode
                    Set lobjNode = lobjNode.Next
                Next
            End If
        Else
            'ĳ��������롣
            sub��Ӵ��� ctrwAllItem.SelectedItem
        End If
        
    End If
        
    ccmdAdd.Enabled = False
    MousePointer = 0
    Me.Enabled = True
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func������(Err.Number, Err.Description)
    MousePointer = 0
    Me.Enabled = True
    sfsub������ "ְҵ�����ý���", "frmSetDoctor", "ccmdAdd_Click", 6666, lstrError, False
    Exit Sub
    Resume
End Sub
Private Sub sub��Ӵ���(ByVal paraAllItemNode As Node)
    Dim i As Long
    Dim lstr���� As String
    Dim lobjNode As Node
    
    If paraAllItemNode.Children = 0 Then Exit Sub
    
    '�Ѵ�����롣
    On Error Resume Next
    If ctrwSelectedItem.Nodes.Count = 0 Then
        ctrwSelectedItem.Nodes.Add , , "R", "������"
    End If
    
    '�Ѵ�����롣
    ctrwSelectedItem.Nodes.Add "R", tvwChild, paraAllItemNode.Key, paraAllItemNode.Text
    
    On Error GoTo errHandler
    
    '���ΰѸô������ĿҶ�ڵ���롣
    For i = 1 To paraAllItemNode.Children
        Set lobjNode = ctrwAllItem.Nodes(paraAllItemNode.Index + i)
        
        lstr���� = lobjNode.Key
        lstr���� = Right(lstr����, Len(lstr����) - 1)
        
        If lstr���� <> "" Then
            If Not mobj���ҽʦ.func�Ƿ������Ŀ(lstr����) Then
                '�ӿ�����Ӹ���Ŀ���á�
                mobj���ҽʦ.Sub��������Ŀ lstr����
            
                '���������Ŀ�����нڵ��key=����,parent=�����ࣩ��
                On Error Resume Next
                ctrwSelectedItem.Nodes.Add lobjNode.Parent.Key, tvwChild, lobjNode.Key, lobjNode.Text
                On Error GoTo errHandler
            End If
        End If
    Next
    Exit Sub
errHandler:
    Err.Raise Err.Number, , Err.Description
End Sub

'���ܣ�ɾ��������Ŀ��һ�����ࡢ������Ŀ��
Private Sub ccmdDel_Click()
    Dim llngIndex As Long
    Dim lobjNode As Node
    Dim i As Long
    
    On Error GoTo errHandler
    If ctrwSelectedItem.SelectedItem Is Nothing Then Exit Sub
    
    MousePointer = 11
    Me.Enabled = False
    If ctrwSelectedItem.SelectedItem.Parent Is Nothing Then
        'ɾ����������
        
        '�ӿ���ɾ����ǰҽʦ���������������Ŀ���á�
        mobj���ҽʦ.Subɾ�����������Ŀ
    
        ctrwSelectedItem.Nodes.Clear
        
    Else
        If Not ctrwSelectedItem.SelectedItem.Parent.Parent Is Nothing Then
            'ɾ��������Ŀ��
            Set lobjNode = ctrwSelectedItem.SelectedItem
            
            '�ӿ���ɾ����ǰҽʦ�����ĸ���Ŀ���á�
            mobj���ҽʦ.Subɾ�������Ŀ Right(lobjNode.Key, Len(lobjNode.Key) - 1)
            
            '�����ڵ���û���ӽڵ㣬ɾ�����ڵ㡣
            If lobjNode.Parent.Children = 0 Then
                ctrwSelectedItem.Nodes.Remove lobjNode.Parent.Key
            Else
                'ɾ����ǰѡ�нڵ㡣
                ctrwSelectedItem.Nodes.Remove lobjNode.Key
            End If
        Else
            'ɾ���������ࡣ
            For i = 1 To ctrwSelectedItem.SelectedItem.Children
                Set lobjNode = ctrwSelectedItem.Nodes(ctrwSelectedItem.SelectedItem.Index + i)
                '�ӿ���ɾ����ǰҽʦ�����ĸ���Ŀ���á�
                mobj���ҽʦ.Subɾ�������Ŀ Right(lobjNode.Key, Len(lobjNode.Key) - 1)
            Next
            
            '�����ڵ���û���ӽڵ㣬ɾ�����ڵ㡣
            If ctrwSelectedItem.SelectedItem.Parent.Children = 0 Then
                ctrwSelectedItem.Nodes.Remove ctrwSelectedItem.SelectedItem.Parent.Key
            Else
                'ɾ����ǰѡ�нڵ㡣
                ctrwSelectedItem.Nodes.Remove ctrwSelectedItem.SelectedItem.Key
            End If
        End If
        
        
        '�����������ˣ�ɾ�����ڵ㡣
        If ctrwSelectedItem.Nodes.Count = 1 Then
            ctrwSelectedItem.Nodes.Clear
        End If
        
    End If
    
    ccmdDel.Enabled = False
    MousePointer = 0
    Me.Enabled = True
    
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func������(Err.Number, Err.Description)
    MousePointer = 0
    Me.Enabled = True
    sfsub������ "ְҵ�����ý���", "frmSetDoctor", "ccmdDel_Click", 6666, lstrError, False
    Exit Sub
    Resume
End Sub

Private Sub ctrwDoctor_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim lcolItemSet As Collection '��ǰҽʦ������Ŀ�ļ��ϡ�
    Dim lcolItem As Variant       'lcolItemSet�е�ĳ��Ԫ��[���룬����]��
    
    On Error GoTo errHandler
    
    ctrwSelectedItem.Nodes.Clear
    ccmdDel.Enabled = False
    ccmdAdd.Enabled = False
    
    If Node.Parent Is Nothing Then
        Label2 = "���������Ŀ�б�"
        mobj���ҽʦ.��� = ""
        Exit Sub
    End If
    
    '����mobj���ҽʦ.���=��ǰ�ڵ��key��
    mobj���ҽʦ.��� = Right(Node.Key, Len(Node.Key) - 1)
    Label2 = Right(Node.Text, Len(Node.Text) - InStr(Node.Text, " ")) & "���������Ŀ�б�"
    
    '��ȡ��ǰҽʦ���������������Ŀ
    Set lcolItemSet = mobj���ҽʦ.���������Ŀ
    
    If lcolItemSet.Count > 0 Then
        ctrwSelectedItem.Nodes.Add , , "R", "������"
    End If
    
    '��ʾ��Ŀ��clstItem��(����+''+����)��
    On Error Resume Next
    Dim lobjNode As Node
    For Each lcolItem In lcolItemSet
    
        '��ȡ����Ŀ��ctrwAllItem���еĽڵ㡣
        Set lobjNode = Nothing
        Set lobjNode = ctrwAllItem.Nodes("I" & lcolItem("����"))
        
        If Not lobjNode Is Nothing Then
            '�������ڵ㡣
            ctrwSelectedItem.Nodes.Add "R", tvwChild, lobjNode.Parent.Key, lobjNode.Parent.Text
            '���������Ŀ�����нڵ��key=����,parent=�����ࣩ��
            ctrwSelectedItem.Nodes.Add lobjNode.Parent.Key, tvwChild, lobjNode.Key, lobjNode.Text
        End If
        
    Next
    
    On Error Resume Next
    If ctrwSelectedItem.Nodes.Count > 0 Then
        ctrwSelectedItem.Nodes(1).Expanded = True
    End If
    
    ccmdDel.Enabled = False
    If ctrwAllItem.SelectedItem Is Nothing Then
        ccmdAdd.Enabled = False
    Else
        ccmdAdd.Enabled = True
    End If
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����ý���", "frmSetDoctor", "ctrwDoctor_NodeClick", 6666, lstrError, False
End Sub

Private Sub ccmdExit_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub clstItem_Click()
    On Error Resume Next
    ccmdDel.Enabled = True
End Sub

Private Sub ctrwAllItem_NodeClick(ByVal Node As MSComctlLib.Node)
    On Error GoTo errHandler
    
    If mobj���ҽʦ.��� = "" Then
        ccmdAdd.Enabled = False
    Else
        ccmdAdd.Enabled = True
    End If
    
    Exit Sub
errHandler:
    'sfsub������ "ְҵ�����ý���", "frmSetDoctor", "ctrwAllItem_NodeClick", Err.Number, Err.Description, False
End Sub



Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set mobj���ҽʦ = Nothing
End Sub

