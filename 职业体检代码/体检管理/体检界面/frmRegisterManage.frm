VERSION 5.00
Object = "{8099FCC2-0A81-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmRegisterManage 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "������"
   ClientHeight    =   7635
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   11070
   Icon            =   "frmRegisterManage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   11070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.OptionButton coptType 
      Caption         =   "���½���"
      Height          =   255
      Index           =   2
      Left            =   3000
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.OptionButton coptType 
      Caption         =   "������"
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.OptionButton coptType 
      Caption         =   "δ�½���"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VSFlex6Ctl.vsFlexGrid cgrdMain 
      Height          =   6375
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   10815
      _cx             =   88492676
      _cy             =   88484845
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
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
   Begin MSComctlLib.ImageList cimg��ťͼ�� 
      Left            =   1440
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar ctlb������ 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11070
      _ExtentX        =   19526
      _ExtentY        =   1005
      ButtonWidth     =   1455
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      ImageList       =   "cimg��ťͼ��"
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog ccmdFile 
      Left            =   0
      Top             =   500
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   6148
   End
   Begin VB.Label clblInfo 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   180
      Left            =   6240
      TabIndex        =   6
      Top             =   720
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ܼ�¼����"
      Height          =   180
      Left            =   5280
      TabIndex        =   5
      Top             =   720
      Width           =   900
   End
   Begin VB.Menu cmnuView 
      Caption         =   "�鿴   "
      Begin VB.Menu cmnuItemView 
         Caption         =   "��ѯ(&Q)"
         Index           =   1
      End
      Begin VB.Menu cmnuItemView 
         Caption         =   "ˢ��"
         Index           =   2
      End
      Begin VB.Menu cmnuItemView 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu cmnuItemView 
         Caption         =   "�˳�(&X)"
         Index           =   4
      End
   End
   Begin VB.Menu cmnuRegister 
      Caption         =   "���Ǽ�   "
      Begin VB.Menu cmnuItemRegister 
         Caption         =   "����Ǽ�(&N)"
         Index           =   1
      End
      Begin VB.Menu cmnuItemRegister 
         Caption         =   "���Ǽ�(&Y)"
         Index           =   2
      End
      Begin VB.Menu cmnuItemRegister 
         Caption         =   "����Ǽ�(&R)"
         Index           =   3
      End
      Begin VB.Menu cmnuItemRegister 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu cmnuItemRegister 
         Caption         =   "�޸�(&U)"
         Index           =   5
      End
      Begin VB.Menu cmnuItemRegister 
         Caption         =   "ɾ��(&D)"
         Index           =   6
      End
   End
   Begin VB.Menu cmnuPrint 
      Caption         =   "��ӡ"
      Begin VB.Menu cmnuItemPrint 
         Caption         =   "����"
         Index           =   1
      End
      Begin VB.Menu cmnuItemPrint 
         Caption         =   "�������"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmRegisterManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnInUse As Boolean

Private WithEvents mobjGUI As cls����ͨ�ö���
Attribute mobjGUI.VB_VarHelpID = -1

'��ѯ����
Private mstr��ʼ���� As String
Private mstr��ֹ���� As String
Private mstr�������� As String
Private mstr��λ���� As String
Private mstr���� As String
Private mstr��쵥�� As String
Private mstr�Թܱ�� As String
Private mstrϵͳ��� As String
'��ѯ���
Private mobjQueryResult As Object

Private mcolIndex As New Collection

'���ܣ����ص�ǰ�����Ƿ��Ѿ����ر�־������ϵͳƽ̨��Ҫ��ġ�
Public Property Get pblnInUse() As Boolean
    On Error Resume Next
    pblnInUse = mblnInUse
End Property

Private Sub cmnuItemPrint_Click(Index As Integer)
    Dim lcol��� As Collection
    On Error GoTo errHandler
    Set lcol��� = New Collection
    Select Case Index
    Case 1
        '��ӡ����
        lcol���.Add cgrdMain.TextMatrix(cgrdMain.Row, 0)
        pobjҵ�����.Sub��ӡ���� "����", lcol���, True
        
    Case 2
        '��ӡ�������
        lcol���.Add cgrdMain.TextMatrix(cgrdMain.Row, 0)
        pobjҵ�����.Sub��ӡ���� "�������", lcol���, True
    End Select
    
    Exit Sub
errHandler:
    sfsub������ "������", "frmRegisterManage", "cmnuItemPrint_Click", Err.Number, Err.Description, False
    
End Sub

Private Sub cmnuItemRegister_Click(Index As Integer)
    On Error GoTo errHandler
    Select Case Index
    Case 1 '����Ǽ�
        FrmRegister.pstrϵͳ��� = ""
        FrmRegister.Show 1, Me
        
        '���²�ѯ��
        sub��ѯ����ʾ
        
    Case 2 '���Ǽ�
        If cgrdMain.Row >= 1 Then
            FrmRegister.pstrϵͳ��� = cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("ϵͳ���"))
        Else
            FrmRegister.pstrϵͳ��� = ""
        End If
        FrmRegister.Show 1, Me
        
        '���²�ѯ��
        sub��ѯ����ʾ
    
    Case 3 '����Ǽ�
        If cgrdMain.Row < 1 Then
            MsgBox "û����Ҫ������ˣ�", vbOKOnly + vbExclamation, "ϵͳ��ʾ"
            Exit Sub
        End If
        FrmRegisterAgain.pstr��ϵͳ��� = cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("ϵͳ���"))
        FrmRegisterAgain.Show 1, Me
        
    Case 5 '�޸�
        If cgrdMain.Row < 1 Then
            MsgBox "û����Ҫ�޸ĵļ�¼��", vbOKOnly + vbExclamation, "ϵͳ��ʾ"
            Exit Sub
        End If
        
        FrmEditRegister.ϵͳ��� = cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("ϵͳ���"))
        FrmEditRegister.Show 1, Me
        
        '���²�ѯ��
        sub��ѯ����ʾ
    
    Case 6 'ɾ��
        If cgrdMain.Row < 1 Then
            MsgBox "û�п���ɾ���ļ�¼��", vbOKOnly + vbExclamation, "ϵͳ��ʾ"
            Exit Sub
        End If
        
        If coptType(0) Then
            If MsgBox("��ȷ��Ҫɾ��������¼��һ��ɾ���󽫲��ָܻ���", vbYesNo + vbQuestion + vbDefaultButton2, "ϵͳ��ʾ") = vbYes Then
                pobjҵ�����.subɾ�����Ǽ� cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("ϵͳ���"))
                oesubSave "ɾ�������Ա��Ϣ����ϵͳ���Ϊ��" & cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("ϵͳ���")) & "������Ϊ��" & cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("����")), "ɾ�������Ա"
                cgrdMain.RemoveItem cgrdMain.Row
            End If
        Else
            MsgBox "���������۵ļ�¼������ɾ����", vbOKOnly + vbExclamation, "ϵͳ��ʾ"
            Exit Sub
        End If
    
    End Select
    Exit Sub
errHandler:
    sfsub������ "������", "frmRegisterManage", "cmnuItemRegister_Click", Err.Number, Err.Description, False
End Sub

Private Sub cmnuItemView_Click(Index As Integer)
    On Error GoTo errHandler
    Select Case Index
    Case 1 '��ѯ
        With frmQuery
            '��ʾ�ɵĲ�ѯ������
            .pstr��ʼ���� = mstr��ʼ����
            .pstr��ֹ���� = mstr��ֹ����
            .pstr�������� = mstr��������
            .pstr���� = mstr����
            .pstr��λ���� = mstr��λ����
            .pstr��쵥�� = mstr��쵥��
            .pstr�Թܱ�� = mstr�Թܱ��
            .pstrϵͳ��� = mstrϵͳ���
            '��ȡ�µĲ�ѯ������
            .Show 1, Me
            If .pblnOk Then
                mstr��ʼ���� = .pstr��ʼ����
                mstr��ֹ���� = .pstr��ֹ����
                mstr�������� = .pstr��������
                mstr��λ���� = .pstr��λ����
                mstr���� = .pstr����
                mstr��쵥�� = .pstr��쵥��
                mstr�Թܱ�� = .pstr�Թܱ��
                mstrϵͳ��� = .pstrϵͳ���
                
                '���²�ѯ��
                sub��ѯ����ʾ
            End If
        End With
    
    Case 2 'ˢ��
        sub��ʾ��ѯ���
    Case 4
        Unload Me
    End Select
    Exit Sub
errHandler:
    sfsub������ "������", "frmRegisterManage", "cmnuItemView_Click", Err.Number, Err.Description, False
End Sub

Private Sub coptType_Click(Index As Integer)
    On Error GoTo errHandler
    sub��ʾ��ѯ���
    
    ctlb������.Buttons(4).Enabled = coptType(1).Value
    cmnuItemRegister(3).Enabled = coptType(1).Value
    
    ctlb������.Buttons(6).Enabled = coptType(0).Value
    ctlb������.Buttons(7).Enabled = coptType(0).Value
    cmnuItemRegister(5).Enabled = coptType(0).Value
    cmnuItemRegister(6).Enabled = coptType(0).Value
    
    cmnuItemPrint(1).Enabled = coptType(0).Value
    cmnuItemPrint(2).Enabled = coptType(2).Value
    Exit Sub
errHandler:
    sfsub������ "������", "frmRegisterManage", "coptType_Click", Err.Number, Err.Description, False
End Sub

Private Sub Form_Load()
    Dim lcol��������ť As New Collection '�������ϵİ�ť��ʼ�����ϡ�
    
    On Error GoTo errHandler
    
    '��������Ѿ���ʼ�������ٽ��г�ʼ����
    If mblnInUse Then Exit Sub
    
    '���ô�������ʹ�õı�־��
    mblnInUse = True
    
    '��ʼ������ͨ�ö���ÿ�������Ӧһ������ͨ�ö���ʵ�������ɻ��ã��м��мǣ���
    Set mobjGUI = New cls����ͨ�ö���
    
    '���ù�����������Ҫ�ĸ��ְ�ť��
    '�޸ģ�2002-7-1�������ȡ�����۵Ĳ�������Ϊ������ѡ��
    With lcol��������ť
        .Add "��ѯ(&Q)108"
        .Add "|"
        .Add "����Ǽ�(&R)101"
        .Add "����Ǽ�(&R)103"
        .Add "|"
        .Add "�޸�"
        .Add "ɾ��"
        .Add "|"
        .Add "����(&O)111"
        .Add "����(&O)112"
        .Add "|"
        .Add "�˳�"
    End With
    
    'Ϊ��Ҫͨ������ͨ�ö�����Ƶĸ��ֿؼ�����ʼֵ��
    With mobjGUI
        Set .Form = Me
        Set .c������ = ctlb������
    End With
    
    '���ý���ͨ�ö����ṩ�ķ������Խ���ؼ����г�ʼ����
    mobjGUI.subInitialize lcol��������ť, ""

    'ȱʡ��ʾ���һ�ܵ������Ա��
    mstr��ʼ���� = Format(DateAdd("d", -7, Date), "yyyy-mm-dd")
    mstr��ֹ���� = Format(Date, "yyyy-mm-dd")
    mstr�������� = ""
    mstr��λ���� = ""
    mstr���� = ""
    mstr��쵥�� = ""
    mstr�Թܱ�� = ""
    
    sub��ѯ����ʾ
    
    ctlb������.Buttons(4).Enabled = coptType(1).Value
    cmnuItemRegister(3).Enabled = coptType(1).Value
    
    ctlb������.Buttons(6).Enabled = coptType(0).Value
    ctlb������.Buttons(7).Enabled = coptType(0).Value
    cmnuItemRegister(5).Enabled = coptType(0).Value
    cmnuItemRegister(6).Enabled = coptType(0).Value

    cmnuItemPrint(1).Enabled = coptType(0).Value
    cmnuItemPrint(2).Enabled = coptType(2).Value

    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "�����沿��", "frmRegistermanage", "Form_Load", 6666, lstrError, False
    Exit Sub
    Resume

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    mblnInUse = False
    
    Set mobjGUI = Nothing
    Set mobjQueryResult = Nothing
End Sub

Public Sub sub��ѯ����ʾ()
    On Error GoTo errHandler
    Set mobjQueryResult = pobjҵ�����.func����������ѯ(mstr��ʼ����, mstr��ֹ����, mstr��������, mstr��λ����, mstr����, mstr��쵥��, mstr�Թܱ��, mstrϵͳ���)
    
    sub��ʾ��ѯ���

    Dim i As Long
    Set mcolIndex = New Collection
    For i = 0 To cgrdMain.Cols - 1
        mcolIndex.Add i, cgrdMain.TextMatrix(0, i)
    Next
    Exit Sub
errHandler:
    sfsub������ "������", "frmRegisterManage", "sub��ѯ����ʾ", Err.Number, Err.Description, True
End Sub

Private Sub sub��ʾ��ѯ���()
    On Error GoTo errHandler
    If coptType(0).Value Then
        mobjQueryResult.Filter = "���״̬='δ�½���'"
    ElseIf coptType(1).Value Then
        mobjQueryResult.Filter = "���״̬='���½���' and ����������<>'' and ����ϵͳ���=''"
    Else
        mobjQueryResult.Filter = "(���״̬='���½���' and  ����������='') or (���״̬='���½���' and ����ϵͳ���<>'')"
    End If
    Set cgrdMain.DataSource = mobjQueryResult
    
    clblInfo = cgrdMain.Rows - 1

    Exit Sub
errHandler:
    sfsub������ "������", "frmRegisterManage", "sub��ʾ��ѯ���", Err.Number, Err.Description, True

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    cgrdMain.Width = Me.ScaleWidth - cgrdMain.Left - 60
    cgrdMain.Height = Me.ScaleHeight - cgrdMain.Top - 60
    
End Sub

Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    On Error GoTo errHandler
    
    Select Case Operate
    Case "��ѯ"
        cmnuItemView_Click 1
        
    Case "����Ǽ�"
        cmnuItemRegister_Click 1
        Cancel = True
        
    Case "����Ǽ�"
        cmnuItemRegister_Click 3
    
    Case "�޸�"
        Cancel = True
        cmnuItemRegister_Click 5
    
    Case "ɾ��"
        Cancel = True
        cmnuItemRegister_Click 6
    Case "����"
        Dim lstrFile As String
        ccmdFile.Filter = "Excel�ļ� (*.xls)|*.xls|�ı��ļ� (*.txt)|*.txt"
        ccmdFile.ShowSave
        lstrFile = ccmdFile.FileName
        If lstrFile <> "" Then
            cgrdMain.SaveGrid lstrFile, flexFileTabText, True
        End If
    Case "����"
        frmTransfer.Show 1
    End Select
    Exit Sub
errHandler:
    sfsub������ "������", "frmRegisterManage", "mobjGUI_BeforeOperate", Err.Number, Err.Description, False
    Cancel = True
    Exit Sub
    Resume
End Sub
