VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSetDoctorPermission 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ҽʦȨ�޹���"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   11970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton ccmdExit 
      Caption         =   "�˳�(&X)"
      Height          =   375
      Left            =   7920
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin MSComctlLib.TreeView ctvPerm 
      Height          =   6015
      Left            =   5400
      TabIndex        =   3
      Top             =   600
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   10610
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Checkboxes      =   -1  'True
      Appearance      =   1
   End
   Begin MSComctlLib.TreeView ctvDept 
      Height          =   6015
      Left            =   2400
      TabIndex        =   2
      Top             =   600
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   10610
      _Version        =   393217
      HideSelection   =   0   'False
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Checkboxes      =   -1  'True
      Appearance      =   1
   End
   Begin VSFlex8Ctl.VSFlexGrid cgrdDoctor 
      Height          =   6015
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1935
      _cx             =   3413
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   "ҽʦ���|ҽʦ����"
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
   Begin VB.Label Label3 
      Caption         =   "����Ȩ�ޣ�"
      Height          =   255
      Left            =   5520
      TabIndex        =   5
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "�ɲ������ң�"
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "ҽʦ�б�"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmSetDoctorPermission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'2012-03-15 �ڵ�� ����������������Ȩ�޿��ƹ���
'����ҽʦ�б�ѡ�е�ĳһ��ҽʦ����������(�ɶ�ѡ)������ѡ���������Ŀ��ò�����
'�ô����Ȩ�ޣ���ϵͳ���á��û���Ȩ������������
'�ⲿ�ֵ�ҽʦ��ֻ����ϵͳ���á�Ա�������б���Ϊ06ְҵ�����Ƶ�ҽʦ���ܷŽ�����
'ҽʦ���һ���ò����򹳻�ȡ��ʱ�����ݿ�������������ӻ�ɾ������������ʱ��ʾ��
'��Ӧ��ȥ�ֵ����������ҵ������������½���"clsPermissionConfigure"��

'-------���غ��޸Ŀ��ҺͿ��ò������뷱��������Ϊ��˵��-------
'1��form_loadʱ����������ҽʦ�����п��ң������ؿ��ò���
'2��ÿ��ѡ��ҽʦʱ��ҽʦ���п��Ҵ򹳣����ظ�ҽʦ�������ҵ����в�������ѡ�����򹳣����в����ڵ�չ��(����鿴)
'3��2�����̣�LoadDocAllDept �� LoadDocAllOperate �� LoadOneDeptNowOperate(LoadOneDeptAllOperate) �� ExpandAllNodes
'4����������ʱ�����������п��ò���ѡ�У���չ��������ȡ��ʱ�����������п��ò���ɾ��
'5�����ò������ʱ��������ǰ�ڵ����ڣ������ӽڵ�Ͷ�Ӧ���и��ڵ�������
'6�����ò���ɾ��ʱ��������ǰ�ڵ����ڣ������ӽڵ㱻ɾ��
'7���ӽڵ����Ϊ�ݹ飬���ڵ����Ϊ���Բ���
'----------------------------˵�����--------------------------

Option Explicit
Private mblnInUse As Boolean    '������ǰ�����Ƿ��Ѽ��ء�
Private mobjCls As Object       '����form��ȫ�ֱ�����ר�ŵ���clsPermissionConfigure
Private DoctorNo As String      '����form��ȫ�ֱ�������¼ҽʦ��ţ������û���ţ�
Private DoctorDept As Object    '����form��ȫ�ֱ�������¼ҽʦ��������
Private DoctorPerm As Object    '����form��ȫ�ֱ�������¼ҽʦ����Ȩ��

Public Property Get pblnInUse() As Boolean
    pblnInUse = mblnInUse
End Property

Private Sub ccmdExit_Click()
    Unload Me
    Set mobjCls = Nothing
    Set DoctorDept = Nothing
    Set DoctorPerm = Nothing
    Set frmSetDoctorPermission = Nothing
End Sub

Private Sub cgrdDoctor_Click()
    If cgrdDoctor.MouseRow < 0 Or cgrdDoctor.MouseCol < 0 Then
        Exit Sub
    Else
        ctvDept.Enabled = True
        ctvPerm.Enabled = True
        DoctorNo = cgrdDoctor.TextMatrix(cgrdDoctor.SelectedRow(0), 0)
        LoadDocAllDept DoctorNo
    End If
End Sub

Private Sub ctvDept_NodeCheck(ByVal Node As MSComctlLib.Node)
    If DoctorNo = "" Then Exit Sub
    
    On Error GoTo errHandler
    Dim rootNode As Object
    
    'Node.Checked=false��ʾȡ��Ȩ��;true��ʾ����Ȩ��
    Call mobjCls.func�޸�ְҵ�����ҽʦ��������(DoctorNo, Node.Checked, Right(Node.Key, Len(Node.Key) - 1))
    If Node.Checked = True Then
        LoadDocAllDept (DoctorNo)
        Set rootNode = func��ÿ��Ҳ����ĸ��ڵ�(Node)
        rootNode.Checked = Node.Checked
        Call mobjCls.func�޸�ҽʦ��������ϵͳȨ��(DoctorNo, Node.Checked, Right(rootNode.Key, Len(rootNode.Key) - 1))
        Call mobjCls.func�޸�ְҵ�����ҽʦ�������ò���(DoctorNo, Node.Checked, Right(rootNode.Key, Len(rootNode.Key) - 1))
        If rootNode.Children > 0 Then Call DownwardModify(rootNode.Child, Node.Checked)
    Else
        Set rootNode = func��ÿ��Ҳ����ĸ��ڵ�(Node)
        rootNode.Checked = Node.Checked
        Call mobjCls.func�޸�ҽʦ��������ϵͳȨ��(DoctorNo, Node.Checked, Right(rootNode.Key, Len(rootNode.Key) - 1))
        Call mobjCls.func�޸�ְҵ�����ҽʦ�������ò���(DoctorNo, Node.Checked, Right(rootNode.Key, Len(rootNode.Key) - 1))
        If rootNode.Children > 0 Then Call DownwardModify(rootNode.Child, Node.Checked)
        LoadDocAllDept (DoctorNo)
    End If
    
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����ý���", "frmSetDoctorPermission", "ctvDept_NodeCheck", 6666, lstrError, False
End Sub

Private Sub ctvPerm_NodeCheck(ByVal Node As MSComctlLib.Node)
    If DoctorNo = "" Then Exit Sub
    
    On Error GoTo errHandler
    
    'Node.Checked=false��ʾȡ��Ȩ��;true��ʾ����Ȩ��
    Call mobjCls.func�޸�ְҵ�����ҽʦ�������ò���(DoctorNo, Node.Checked, Right(Node.Key, Len(Node.Key) - 1))
    If Node.Children > 0 Then Call DownwardModify(Node.Child, Node.Checked)
    If Node.Checked = True Then Call UpwardModify(Node, Node.Checked)
    
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����ý���", "frmSetDoctorPermission", "ctvPerm_NodeCheck", 6666, lstrError, False
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    
    cgrdDoctor.SelectionMode = flexSelectionListBox
    cgrdDoctor.AllowSelection = False
    
    Set mobjCls = CreateObject("ְҵ������.clsPermissionConfigure")
    DoctorNo = ""
    ctvDept.Enabled = False
    ctvPerm.Enabled = False
    LoadDoctor
    LoadAllDept
    
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����ý���", "frmSetDoctorPermission", "Form_Load", 6666, lstrError, False
End Sub

Sub LoadDoctor()
    Dim lobjRec As Object
    Dim i As Integer
    On Error GoTo errHandler
    
    Set lobjRec = pobjDict.FetchEx("Ա���ֵ�")
    
    If lobjRec.recordcount = 0 Then Exit Sub
    lobjRec.movefirst
'    lobjRec.Filter = "����=06"
    For i = 1 To lobjRec.recordcount
        If lobjRec("���") <> "0000" And lobjRec("���") <> "gues" Then
            cgrdDoctor.AddItem lobjRec("���") & vbTab & lobjRec("����"), cgrdDoctor.Rows
        End If
        lobjRec.movenext
    Next
    
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����ý���", "frmSetDoctorPermission", "LoadDoctor", 6666, lstrError, False
End Sub

Sub LoadAllDept()
    Dim lobjDeptSet As Object
    Set lobjDeptSet = pobjDict.Fetch("ְҵ���������ֵ�")
    If lobjDeptSet.recordcount > 0 Then lobjDeptSet.movefirst
    Do While Not lobjDeptSet.EOF
        ctvDept.Nodes.Add , , "R" & lobjDeptSet("���"), lobjDeptSet("���") & "  " & lobjDeptSet("����")
        lobjDeptSet.movenext
    Loop
End Sub

'���ص�ǰҽʦ�����п���
Sub LoadDocAllDept(ByVal paraDoctorNo As String)
    On Error GoTo errHandler
    
    Dim lobjRec As Object
    Set lobjRec = pobjDict.Fetch("ְҵ���������ֵ�")

    '���֮ǰ���ع����ң�����ʾʱ���뽫ǰ��Ĺ���ȥ��(������Щ���)
    ctvDept.Checkboxes = False
    ctvDept.Checkboxes = True
    
    '���»��ҽ��������Ϣ��ѡ�п��Ҵ�
    '������Ϣ���������Ӧ�Ŀ����£�ҽʦ����Ȩ�޴�
    ctvPerm.Nodes.Clear
    Set DoctorDept = mobjCls.func��ȡְҵ����쵥��ҽʦ����(paraDoctorNo)
    If DoctorDept.recordcount > 0 Then DoctorDept.movefirst
    Do While Not DoctorDept.EOF
        ctvDept.Nodes.Item("R" & DoctorDept("���ұ��").Value).Checked = True
        DoctorDept.movenext
    Loop

    '���ؿ���֮�󣬼������е����в���(������)
    LoadDocAllOperate paraDoctorNo

    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����ý���", "frmSetDoctorPermission", "LoadDocAllDept", 6666, lstrError, False
End Sub

'ˢ�µ�ǰҽʦ�����в���(������)
Sub LoadDocAllOperate(ByVal paraDoctorNo As String)
    On Error GoTo errHandler
    
    Dim i As Integer
    Set DoctorPerm = mobjCls.func��ȡְҵ����쵥��ҽʦ����Ȩ��(paraDoctorNo) '[���в���Ȩ��]
    For i = 1 To ctvDept.Nodes.Count
        If ctvDept.Nodes.Item(i).Checked = True Then
            Dim L As Integer
            Dim lstrDeptName As String
            L = Len(ctvDept.Nodes.Item(i).Text) - 4  'key����λ�Ǽ�����ĸ��R����
            lstrDeptName = Right(ctvDept.Nodes.Item(i).Text, L - 0)
            Call LoadOneDeptNowOperate(paraDoctorNo, lstrDeptName, DoctorPerm)
        End If
    Next

    ExpandAllNodes

    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����ý���", "frmSetDoctorPermission", "LoadDocAllOperate", 6666, lstrError, False
End Sub

'���ص�ǰҽʦ�ĵ������ҵ����в���(������)
Sub LoadOneDeptNowOperate(ByVal paraDoctorNo As String, ByVal paraDeptName As String, ByVal paraDoctorPerm As Object)
    Call LoadOneDeptAllOperate(paraDeptName)
    
    If paraDoctorPerm.recordcount = 0 Then Exit Sub
    Dim i As Integer
    For i = 1 To ctvPerm.Nodes.Count
        paraDoctorPerm.movefirst
        Do While Not paraDoctorPerm.EOF
            If ctvPerm.Nodes.Item(i).Text = paraDoctorPerm("Ȩ����") Then
                ctvPerm.Nodes.Item(i).Checked = True
            End If
            paraDoctorPerm.movenext
        Loop
    Next
End Sub

'���ص������ҵ����в������������򹳣�
Sub LoadOneDeptAllOperate(ByVal paraDeptName As String)
    Dim lobjPerm As Object
    Set lobjPerm = mobjCls.func��ȡְҵ����쵥���������в���Ȩ��(paraDeptName)
    
    If lobjPerm.recordcount = 0 Then Exit Sub
    lobjPerm.movefirst
    Do While Not lobjPerm.EOF
        If IsNull(lobjPerm("�ϼ�������")) = True Then
            ctvPerm.Nodes.Add , , "R" & lobjPerm("������"), lobjPerm("������")
        Else
            ctvPerm.Nodes.Add "R" & lobjPerm("�ϼ�������"), tvwChild, "R" & lobjPerm("������"), lobjPerm("������")
        End If
        lobjPerm.movenext
    Loop
End Sub

'��treeview��ĳ���ض��ڵ㿪ʼ����������ӽڵ������ӻ�ɾ������(����)
Sub UpwardModify(ByVal paraNode As Object, ByVal paraCheck As Boolean)
    Dim rootNode As Object
    Set rootNode = func��ÿ��Ҳ����ĸ��ڵ�(paraNode)
    Do While paraNode <> rootNode
        paraNode.Checked = paraCheck
        Call mobjCls.func�޸�ְҵ�����ҽʦ�������ò���(DoctorNo, paraNode.Checked, Right(paraNode.Key, Len(paraNode.Key) - 1))
        Set paraNode = paraNode.Parent
    Loop
    paraNode.Checked = paraCheck
    Call mobjCls.func�޸�ְҵ�����ҽʦ�������ò���(DoctorNo, paraNode.Checked, Right(paraNode.Key, Len(paraNode.Key) - 1))
End Sub

'��treeview��ĳ���ض��ڵ㿪ʼ����������ӽڵ������ӻ�ɾ������(�ݹ�)
Sub DownwardModify(ByVal paraNode As Object, ByVal paraCheck As Boolean)
    paraNode.Checked = paraCheck
    Call mobjCls.func�޸�ְҵ�����ҽʦ�������ò���(DoctorNo, paraNode.Checked, Right(paraNode.Key, Len(paraNode.Key) - 1))
    
    If paraNode.Children > 0 Then Call DownwardModify(paraNode.Child, paraCheck)
    If paraNode <> paraNode.LastSibling Then Call DownwardModify(paraNode.Next, paraCheck)
End Sub

'չ�����ò��������нڵ㡣
Sub ExpandAllNodes()
    Dim i As Integer
    For i = 1 To ctvPerm.Nodes.Count
        ctvPerm.Nodes(i).Expanded = True
    Next
End Sub

'��ÿ��Ҷ�Ӧ���ò����ĸ��ڵ㡣�ദ�ط��õ��ģ��Ƚ���Ҫ�ĺ�����
'1�������ݿ��п��ò�������������ȷ��Ҫ��
'(1)ÿһ���ҵĿ��ò����ܽڵ㣬��������Ϊ��ְҵ�����_XXX����ʽ���硰ְҵ�����_��ٿƽ��¼�롱
'(2)�ܽ��������ӽڵ㣬����ʱ�����ԡ�ְҵ�����_XXX_��Ϊǰ׺�����������Ӧ�������ơ��硰ְҵ�����_��ٿƽ��¼��_���桱
Private Function func��ÿ��Ҳ����ĸ��ڵ�(ByVal paraNodeDept As Object) As Object
    On Error GoTo errHandler
    
    Dim i, idx As Integer
    Dim returnNode As Object
    Dim strArray
    strArray = Split(paraNodeDept.Text, "_", -1, vbBinaryCompare)
    If UBound(strArray) = 0 Then strArray = Split(paraNodeDept.Text, "  ", -1, vbBinaryCompare)
    For i = 1 To ctvPerm.Nodes.Count
        idx = InStr(ctvPerm.Nodes.Item(i).Text, strArray(1))
        If idx <> 0 Then Set returnNode = ctvPerm.Nodes.Item(i): Exit For
    Next
    Set func��ÿ��Ҳ����ĸ��ڵ� = returnNode
    
    Exit Function
errHandler:
    Dim lstrError  As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����ý���", "frmSetDoctorPermission", "func��ÿ��Ҳ����ĸ��ڵ�", 6666, lstrError, False
End Function

Private Sub Form_Resize()
    On Error Resume Next
    cgrdDoctor.Height = Me.ScaleHeight - cgrdDoctor.Top - 20
    ctvDept.Height = cgrdDoctor.Height
    ctvPerm.Height = cgrdDoctor.Height
End Sub
