VERSION 5.00
Object = "{8099FCC2-0A81-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frm�����շѱ�׼ 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "�շѱ�׼����"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   10335
   ClipControls    =   0   'False
   Icon            =   "frm�����շѱ�׼.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   10335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame Frame1 
      Caption         =   "����ѡ����ϵͳ��ʹ�÷�Χ"
      ForeColor       =   &H00C00000&
      Height          =   6375
      Left            =   3240
      TabIndex        =   7
      Top             =   960
      Width           =   7095
      Begin VB.TextBox ctxt�ɱ�׼���� 
         Height          =   375
         Left            =   4440
         TabIndex        =   9
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton ccmdDel 
         Caption         =   "-->"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2520
         Width           =   495
      End
      Begin VB.CommandButton ccmdAdd 
         Caption         =   "<--"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1920
         Width           =   495
      End
      Begin VSFlex6Ctl.vsFlexGrid cgrdMain 
         Height          =   5175
         Left            =   120
         TabIndex        =   1
         Top             =   1080
         Width           =   3735
         _cx             =   23861692
         _cy             =   23864232
         _ConvInfo       =   1
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
         BackColorBkg    =   16777215
         BackColorAlternate=   15791081
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
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   "�շ���Ŀ             |����     |���� "
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
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
      Begin VB.TextBox ctxt��׼���� 
         Height          =   375
         Left            =   1200
         TabIndex        =   0
         Top             =   360
         Width           =   3015
      End
      Begin MSComctlLib.TreeView ctvwItem 
         Height          =   5175
         Left            =   4560
         TabIndex        =   2
         Top             =   1080
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   9128
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   353
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   6
         Appearance      =   1
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��׼���ƣ�"
         Height          =   180
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   900
      End
   End
   Begin MSComctlLib.ImageList cimg��ťͼ�� 
      Left            =   5760
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar ctlb���� 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "cimg��ťͼ��"
      _Version        =   393216
   End
   Begin MSComctlLib.TreeView ctvwStandard 
      Height          =   6105
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   3090
      _ExtentX        =   5450
      _ExtentY        =   10769
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      Appearance      =   1
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "��ϵͳ��ʹ�÷�Χ����׼"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   2175
   End
End
Attribute VB_Name = "frm�����շѱ�׼"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public pblnInUse As Boolean

Private WithEvents mobjGUI As cls����ͨ�ö���
Attribute mobjGUI.VB_VarHelpID = -1

Private pint��Ŀ���� As Integer

Private Sub ccmdAdd_Click()
    Dim i As Long
    On Error GoTo errhandler
    If ctvwItem.SelectedItem Is Nothing Then
        Err.Raise 6666, , "��ѡ��Ҫ��ӵ��շ���Ŀ��"
    ElseIf ctvwItem.SelectedItem.Key = "s" Then
        Err.Raise 6666, , "��ѡ��Ҫ��ӵ��շ���Ŀ��"
    End If
    sub��ӵ�ǰ�� ctvwItem.SelectedItem
    Exit Sub
errhandler:
    sfsub������ "�շѽ��沿��", "frm�����շѱ�׼", "ccmdAdd_Click", Err.Number, Err.Description, False
    Exit Sub
    Resume
    
End Sub

Private Sub sub��ӵ�ǰ��(ByVal paraNode As Node)
    Dim lobjChild As Node
    Dim i As Long
    On Error GoTo errhandler
    If Len(paraNode.Key) = pint��Ŀ���� * 3 + 1 Then
        'ѡ�����ĩ����Ŀ
        sub���ָ����Ŀ Right(paraNode.Key, Len(paraNode.Key) - 1)
    Else
        '��������¼���Ŀ��
        If paraNode.Children > 0 Then
            Set lobjChild = paraNode.Child
            For i = 1 To paraNode.Children
                sub��ӵ�ǰ�� lobjChild
                Set lobjChild = lobjChild.Next
            Next
        End If
    End If
    Exit Sub
errhandler:
    sfsub������ "�շѽ��沿��", "frm�����շѱ�׼", "sub��ӵ�ǰ��", Err.Number, Err.Description, True
    Exit Sub
    Resume
End Sub
Private Sub sub���ָ����Ŀ(ByVal para��� As String)
    On Error GoTo errhandler
    '������Ŀ�Ƿ�����ӡ�
    Dim i As Long
    For i = 1 To cgrdMain.Rows - 1
        If cgrdMain.TextMatrix(i, 3) = para��� Then
            Exit Sub
        End If
    Next
    
    '��Ҫ��ӡ�
    Dim lobjItem As Object
    Set lobjItem = CreateObject("�շѶ��󲿼�.cls�շ���Ŀ")
    lobjItem.�շ���Ŀ��� = para���
    cgrdMain.Rows = cgrdMain.Rows + 1
    cgrdMain.TextMatrix(cgrdMain.Rows - 1, 0) = lobjItem.�շ���Ŀ����
    cgrdMain.TextMatrix(cgrdMain.Rows - 1, 1) = lobjItem.����
    cgrdMain.TextMatrix(cgrdMain.Rows - 1, 2) = 1
    cgrdMain.TextMatrix(cgrdMain.Rows - 1, 3) = para���
    
    Exit Sub
errhandler:
    sfsub������ "�շѽ��沿��", "frm�����շѱ�׼", "sub���ָ����Ŀ", Err.Number, Err.Description, True
    Exit Sub
    Resume
    
End Sub

Private Sub ccmdDel_Click()
    On Error GoTo errhandler
    If cgrdMain.Rows = 1 Then
        MsgBox "û���շ���Ŀ��ȥ����", vbOKOnly + vbExclamation, "ϵͳ��ʾ"
        Exit Sub
    ElseIf cgrdMain.Row < 1 Then
        MsgBox "����������ѡ��Ҫȥ�����շ���Ŀ��", vbOKOnly + vbExclamation, "ϵͳ��ʾ"
        Exit Sub
    End If
    cgrdMain.RemoveItem cgrdMain.Row
    
    Exit Sub
errhandler:
    sfsub������ "�շѽ��沿��", "frm�����շѱ�׼", "ccmdDel_Click", Err.Number, Err.Description, False
    Exit Sub
    Resume
End Sub

Private Sub cgrdMain_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    'ֻ���޸ĵ��ۡ�������
    If Col = 0 Then Cancel = True
End Sub


Private Sub ctvwItem_DblClick()
    ccmdAdd_Click
    
End Sub

Private Sub ctvwStandard_NodeClick(ByVal Node As MSComctlLib.Node)
    On Error GoTo errhandler
    Dim i As Long
    
    If Left(Node.Key, 1) = "S" Then
        ctxt��׼���� = Node.Text
        ctxt�ɱ�׼���� = Node.Text
        Frame1.Caption = "��׼��" & Node.Text
        Frame1.Enabled = True
        
        '��ʾ��׼��Ϣ��
        Dim lobj��׼ As Object
        Dim lcol��Ŀ As Collection
        Set lobj��׼ = CreateObject("�շѶ��󲿼�.cls�շѱ�׼")
        lobj��׼.�շѱ�׼���� = Node.Text
        Set lcol��Ŀ = lobj��׼.�շ���Ŀ
        cgrdMain.Rows = lcol��Ŀ.Count + 1
        For i = 1 To lcol��Ŀ.Count
            cgrdMain.TextMatrix(i, 0) = lcol��Ŀ(i)("�շ���Ŀ����")
            cgrdMain.TextMatrix(i, 1) = lcol��Ŀ(i)("����")
            cgrdMain.TextMatrix(i, 2) = lcol��Ŀ(i)("����")
            cgrdMain.TextMatrix(i, 3) = lcol��Ŀ(i)("�շ���Ŀ���")
        Next
    Else
        ctxt�ɱ�׼���� = ""
        ctxt��׼���� = ""
        If Node.Parent Is Nothing Then
            Frame1.Caption = "��ѡ�������ʹ�÷�Χ�����"
            Frame1.Enabled = False
        Else
            Frame1.Caption = "���" & Node.Parent.Text & "-" & Node.Text & "�ı�׼"
            Frame1.Enabled = True
        End If
        cgrdMain.Rows = 1
    End If
    Exit Sub
errhandler:
    sfsub������ "�շѽ��沿��", "frm�����շѱ�׼", "ctvwStandard_NodeClick", Err.Number, Err.Description, False
    Exit Sub
    Resume

End Sub

Private Sub Form_Load()
    Dim lobjRec As Object
    
    On Error GoTo errhandler
    
    If pblnInUse = True Then Exit Sub
    
    pblnInUse = True

    '��ʼ��������
    Dim lcol��������ť As Collection
    Set lcol��������ť = New Collection
    Set mobjGUI = New cls����ͨ�ö���
    Set mobjGUI.Form = Me
    Set mobjGUI.c������ = ctlb����

    lcol��������ť.Add "���"
    lcol��������ť.Add "ɾ��"
    lcol��������ť.Add "|"
    lcol��������ť.Add "����"
    lcol��������ť.Add "|"
    lcol��������ť.Add "�˳�"
    
    mobjGUI.subInitialize lcol��������ť, ""
    
    '��ȡ���е��շѱ�׼��
    Dim lobjBase As Object
    Dim lstrSysName As String
    Set lobjBase = dafuncGetData("select * from �շѹ���_��ϵͳ���������� order by ϵͳ��,������")
    Do While Not lobjBase.EOF
        ctvwStandard.Nodes.Add , , lobjBase!ϵͳ��, lobjBase!ϵͳ��
        lstrSysName = lobjBase!ϵͳ��
        Do While lstrSysName = lobjBase!ϵͳ��
            ctvwStandard.Nodes.Add lstrSysName, tvwChild, "F" & lobjBase!���, lobjBase!������
            
            '��ȡ����ϵͳ�����������б�׼��
            Set lobjRec = dafuncGetData("select distinct �շѱ�׼���� from �շѹ���_�շѱ�׼��Ϣ�� where ��������=" & lobjBase!���)
            Do While Not lobjRec.EOF
                ctvwStandard.Nodes.Add "F" & lobjBase!���, tvwChild, "S" & lobjRec!�շѱ�׼����, lobjRec!�շѱ�׼����
                lobjRec.MoveNext
            Loop
            
            ctvwStandard.Nodes("F" & lobjBase!���).Expanded = True
            
            lobjBase.MoveNext
            If lobjBase.EOF Then Exit Do
            
        Loop
    Loop
    If ctvwStandard.Nodes.Count > 1 Then
        ctvwStandard.Nodes(2).Selected = True
        Frame1.Enabled = True
    Else
        Frame1.Enabled = False
    End If
    
    '��ʼ���շ���Ŀ����
    Dim lint���� As Long
    Dim lstrKey As String
    
    pint��Ŀ���� = Val(pobj�շѹ���.ҵ������("��Ŀ����"))
    If pint��Ŀ���� = 0 Then pint��Ŀ���� = 2
    
    ctvwItem.Nodes.Add , , "s", "�շ���Ŀ"
    For lint���� = 1 To pint��Ŀ����
        Set lobjRec = dafuncGetData("select * from �շѹ���_�շ���Ŀ�ֵ�� where Len(�շ���Ŀ���) =" & lint���� * 3 & " order by �շ���Ŀ���")
        Do While (Not lobjRec.EOF)
            lstrKey = "s" & lobjRec("�շ���Ŀ���").Value
            ctvwItem.Nodes.Add "s" & Mid(lstrKey, 2, ((lint���� - 1) * 3)), tvwChild, lstrKey, lobjRec("�շ���Ŀ����").Value
            lobjRec.MoveNext
        Loop
    Next
    ctvwItem.Nodes(1).Expanded = True
    
    cgrdMain.Cols = 4
    cgrdMain.ColHidden(3) = True '�����շ���Ŀ��š�
    cgrdMain.Editable = True
    Exit Sub
errhandler:
    sfsub������ "�շѽ��沿��", "frm�����շѱ�׼", "Form_Load", Err.Number, Err.Description, False
    Exit Sub
    Resume
End Sub

Private Sub Form_Unload(Cancel As Integer)
    pblnInUse = False
End Sub

Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    Dim i As Long
    Dim lobj��׼ As Object
    On Error GoTo errhandler
    Select Case Operate
    Case "���"
        Cancel = True
        If ctvwStandard.SelectedItem Is Nothing Then
            Err.Raise 6666, , "��ѡ��ʹ�÷�Χ��"
        ElseIf Left(ctvwStandard.SelectedItem.Key, 1) <> "S" And Left(ctvwStandard.SelectedItem.Key, 1) <> "F" Then
            Err.Raise 6666, , "��ѡ��ʹ�÷�Χ��"
        End If
        
        ctxt�ɱ�׼���� = ""
        ctxt��׼����.Text = ""
        cgrdMain.Rows = 1
        ctxt��׼����.SetFocus
        Frame1.Caption = "��ӱ�׼��"
        
    Case "����"
        Cancel = True
        If ctxt��׼����.Text = "" Then
            ctxt��׼����.SetFocus
            Err.Raise 6666, , "�������շѱ�׼���ƣ�"
        End If
        If cgrdMain.Rows = 1 Then
            Err.Raise 6666, , "������շ���Ŀ��"
        End If
        Set lobj��׼ = CreateObject("�շѶ��󲿼�.cls�շѱ�׼")
        If ctxt�ɱ�׼����.Text <> "" Then
            lobj��׼.�շѱ�׼���� = ctxt�ɱ�׼����
        End If
        If Left(ctvwStandard.SelectedItem.Key, 1) = "S" Then
            lobj��׼.�������� = Right(ctvwStandard.SelectedItem.Parent.Key, Len(ctvwStandard.SelectedItem.Parent.Key) - 1)
        Else
            lobj��׼.�������� = Right(ctvwStandard.SelectedItem.Key, Len(ctvwStandard.SelectedItem.Key) - 1)
        End If
        
        For i = 1 To cgrdMain.Rows - 1
            lobj��׼.sub�����Ŀ cgrdMain.TextMatrix(i, 3), cgrdMain.TextMatrix(i, 0), cgrdMain.TextMatrix(i, 1), cgrdMain.TextMatrix(i, 2)
        Next
        lobj��׼.sub���� (ctxt��׼����.Text)
        If ctxt�ɱ�׼����.Text <> "" Then
            ctvwStandard.SelectedItem.Text = ctxt��׼����.Text
            ctvwStandard.SelectedItem.Key = "S" & ctxt��׼����.Text
        Else
            '��ӽڵ㡣
            Dim lstrParent As String
            If Left(ctvwStandard.SelectedItem.Key, 1) = "S" Then
                lstrParent = ctvwStandard.SelectedItem.Parent.Key
            Else
                lstrParent = ctvwStandard.SelectedItem.Key
            End If
            ctvwStandard.Nodes.Add lstrParent, tvwChild, "S" & ctxt��׼����.Text, ctxt��׼����.Text
            ctvwStandard.Nodes("S" & ctxt��׼����.Text).Selected = True
            ctxt�ɱ�׼����.Text = ctxt��׼����.Text
        End If
        
    Case "ɾ��"
        Cancel = True
        
        Set lobj��׼ = CreateObject("�շѶ��󲿼�.cls�շѱ�׼")
        If ctxt�ɱ�׼����.Text = "" Then
            Err.Raise 6666, , "��ѡ��Ҫɾ���ı�׼��������ѡ��������÷�Χ��"
        End If
        lobj��׼.�շѱ�׼���� = ctxt�ɱ�׼����
        lobj��׼.subɾ����׼
        
        ctvwStandard.Nodes.Remove ctvwStandard.SelectedItem.Key
        
    End Select

    Exit Sub
errhandler:
    sfsub������ "�շѽ��沿��", "frm�����շѱ�׼", "mobjGUI_BeforeOperate", Err.Number, Err.Description, False
    Exit Sub
    Resume

End Sub
