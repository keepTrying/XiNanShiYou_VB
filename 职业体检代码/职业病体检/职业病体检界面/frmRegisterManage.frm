VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRegisterManage 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ְҵ�������Ǽ�"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   13755
   Icon            =   "frmRegisterManage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   13755
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox ctxtLabelNum 
      Height          =   270
      Left            =   8760
      TabIndex        =   11
      Text            =   "2"
      Top             =   900
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.OptionButton coptType 
      Caption         =   "�����"
      Height          =   255
      Index           =   2
      Left            =   1200
      TabIndex        =   9
      Top             =   960
      Width           =   975
   End
   Begin VB.OptionButton coptType 
      Caption         =   "������"
      Height          =   255
      Index           =   5
      Left            =   5040
      TabIndex        =   8
      Top             =   960
      Width           =   1095
   End
   Begin VB.OptionButton coptType 
      Caption         =   "���½���"
      Height          =   255
      Index           =   4
      Left            =   3720
      TabIndex        =   7
      Top             =   960
      Width           =   1095
   End
   Begin VB.OptionButton coptType 
      Caption         =   "δ�½���"
      Height          =   255
      Index           =   3
      Left            =   2280
      TabIndex        =   6
      Top             =   960
      Width           =   1095
   End
   Begin VSFlex8Ctl.VSFlexGrid cgrdMain 
      Height          =   6015
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   12375
      _cx             =   2088785220
      _cy             =   2088774002
      Appearance      =   1
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
      ScrollTrack     =   -1  'True
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
   Begin VB.OptionButton coptType 
      Caption         =   "δ���嵥"
      Height          =   300
      Index           =   1
      Left            =   9480
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton coptType 
      Caption         =   "δ����"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Value           =   -1  'True
      Width           =   975
   End
   Begin MSComctlLib.Toolbar ctlb������ 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13755
      _ExtentX        =   24262
      _ExtentY        =   1005
      ButtonWidth     =   1455
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      ImageList       =   "cimg��ťͼ��"
      _Version        =   393216
      Begin MSComctlLib.ImageList cimg��ťͼ�� 
         Left            =   0
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   393216
      End
   End
   Begin MSComDlg.CommonDialog ccmdFile 
      Left            =   0
      Top             =   500
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   6148
   End
   Begin VB.Label clblLabelNum 
      Caption         =   "��ǩ��"
      Height          =   255
      Left            =   8160
      TabIndex        =   10
      Top             =   960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label clblInfo 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   180
      Left            =   7440
      TabIndex        =   4
      Top             =   960
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ܼ�¼����"
      Height          =   180
      Left            =   6360
      TabIndex        =   3
      Top             =   960
      Width           =   900
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
Public selectneednum As String '�������֤��ѯʱ�õ���ϵͳ���
Private mstr��ʼ���� As String
Public mstr��ֹ���� As String
Private mstr�������� As String
Private mstr��λ���� As String
Private mstr���� As String
Private mstr��쵥�� As String
Private mstr�Թܱ�� As String
Private mstrϵͳ��� As String
Private mstr���֤�� As String
'��ѯ���
Private mobjQueryResult As Object
Private mintState As Integer
Private mcolIndex As New Collection
Private hasQueryWindows As Boolean
Private printLabelNumLegal As Boolean

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
    sfsub������ "ְҵ������", "frmRegisterManage", "cmnuItemPrint_Click", Err.Number, Err.Description, False
    
End Sub

Private Sub cmnuItemRegister_Click(Index As Integer)
    On Error GoTo errHandler
    Select Case Index
    Case 1 '����Ǽ�
        ���ʱ�־ = 0 '���ڿ��ƽ���
        FrmRegister.pstrϵͳ��� = ""
        
        
        'frmˢ����.Show 1, Me
'        FrmRegister.Move 1000, 500
        FrmRegister.Show 0, Me
'        FrmRegister.Move 700, 350
        '���²�ѯ��
'        sub��ѯ����ʾ
        '2012-07-13 �ڵ�� ��
        'û������Щ���ܣ���ȥ��
'''    Case 2 '���Ǽ�
'''        If cgrdMain.Row >= 1 Then
'''            FrmRegister.pstrϵͳ��� = cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("ϵͳ���"))
'''        Else
'''            FrmRegister.pstrϵͳ��� = ""
'''        End If
'''        FrmRegister.Show 1, Me
'''
'''        '���²�ѯ��
'''        sub��ѯ����ʾ
'''
'''    Case 3 '����Ǽ�
'''        If cgrdMain.Row < 1 Then
'''            MsgBox "û����Ҫ������ˣ�", vbOKOnly + vbExclamation, "ϵͳ��ʾ"
'''            Exit Sub
'''        End If
'''        FrmRegisterAgain.pstr��ϵͳ��� = cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("ϵͳ���"))
'''        FrmRegisterAgain.Show 1, Me
        '2012-07-13 �ڵ�� ��
    Case 5 '�޸�
        If cgrdMain.Row < 1 Then
            MsgBox "û����Ҫ�޸ĵļ�¼��", vbOKOnly + vbExclamation, "ϵͳ��ʾ"
            Exit Sub
        End If
        
        FrmRegister.clblsysno = cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("ϵͳ���"))
        '2012-07-11 �ڵ�� ��
        '�޸�ʱ��Ĭ��ˢ���֤�޸�
        FrmRegister.Check���֤.Value = 1
        '2012-07-11 �ڵ�� ��
        
        ���ʱ�־ = 1
        FrmRegister.Show 0, Me
        
        '���²�ѯ��
        'sub��ʾ��ѯ���
    
    Case 6 'ɾ��
        If cgrdMain.Row < 1 Then
            MsgBox "û�п���ɾ���ļ�¼��", vbOKOnly + vbExclamation, "ϵͳ��ʾ"
            Exit Sub
        End If
        
        If coptType(0) Or coptType(1) Or coptType(2) Or coptType(3) Then
            If MsgBox("��ȷ��Ҫɾ��������¼��һ��ɾ���󽫲��ָܻ���", vbYesNo + vbQuestion + vbDefaultButton2, "ϵͳ��ʾ") = vbYes Then
                
                pobjҵ�����.subɾ�����Ǽ� cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("ϵͳ���"))
                
                cgrdMain.RemoveItem cgrdMain.Row
                clblInfo = cgrdMain.rows - 1
                
                '2012-06-13 �ڵ�� ��
                'ɾ������¼ʱ��ɾ���������Ա������Ƭ�����֤��Ƭ
                Dim lobjDelPhoto As Object
                Set lobjDelPhoto = CreateObject("ְҵ������.clsPersonExamed")
                lobjDelPhoto.funcɾ�����֤��Ƭ cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("ϵͳ���")) & "IDcard", "ְҵ�����"
                lobjDelPhoto.funcɾ��������Ƭ cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("ϵͳ���")), "ְҵ�����"
                sub��ѯ����ʾ
                '2012-06-13 �ڵ�� ��
            End If
        ElseIf coptType(4).Value Then
            MsgBox "���������۵ļ�¼������ɾ����", vbOKOnly + vbExclamation, "ϵͳ��ʾ"
            Exit Sub
        ElseIf coptType(5).Value Then
            If MsgBox("��ȷ��Ҫɾ��������¼��һ��ɾ���󽫲��ָܻ���", vbYesNo + vbQuestion + vbDefaultButton2, "ϵͳ��ʾ") = vbYes Then
                dafuncGetData "update ְҵ�����_��������Ϣ�� set ����״̬='' where ϵͳ���='" & cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("ϵͳ���")) & "'"
            End If
            sub��ѯ����ʾ
            Exit Sub
        End If
    '2012-08-17 �ڵ�� ��
    '���Ӹ���Ǽǹ���
    Case 7  '����
        If cgrdMain.rows < 1 Then
            MsgBox "û�п��Ը���ļ�¼��", vbOKOnly + vbExclamation, "ϵͳ��ʾ"
            Exit Sub
        End If
        If cgrdMain.Row < 1 Or cgrdMain.Row > cgrdMain.rows - 1 Then
            MsgBox "��ѡ��Ҫ����ļ�¼��", vbOKOnly + vbExclamation, "ϵͳ��ʾ"
            Exit Sub
        End If
        ���ʱ�־ = 2
        FrmRegisterAgain.clblsysno = cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("ϵͳ���"))
        FrmRegisterAgain.pstr����ϵͳ��� = cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("����ϵͳ���"))
        
        '����ʱ��Ĭ��ˢ���֤�޸�
'        FrmRegisterAgain.Check���֤.Value = 1
        FrmRegisterAgain.Show 0, Me
        
        '���²�ѯ��
        sub��ѯ����ʾ
    '2012-08-17 �ڵ�� ��
    Case 8
        If cgrdMain.Row < 1 Then
            MsgBox "û����Ҫ����ļ�¼��", vbOKOnly + vbExclamation, "ϵͳ��ʾ"
            Exit Sub
        End If
        
        pstrPhoto = cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("ϵͳ���"))
        FrmPhoto.clblsysno.Text = pstrPhoto
        '2012-07-11 �ڵ�� ��
        '�޸�ʱ��Ĭ��ˢ���֤�޸�
'        frmPhoto.Check���֤.Value = 1
        '2012-07-11 �ڵ�� ��
        
'        ���ʱ�־ = 1
        FrmPhoto.Show 0, Me
    End Select
    Exit Sub
errHandler:
    sfsub������ "ְҵ������", "frmRegisterManage", "cmnuItemRegister_Click", Err.Number, Err.Description, False
End Sub

Private Sub cmnuItemView_Click(Index As Integer)
    On Error GoTo errHandler
    Select Case Index
    Case 1 '��ѯ
        hasQueryWindows = False
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
           .pstr���֤�� = mstr���֤��
            '��ȡ�µĲ�ѯ������
            .Show 1, Me
            If .pblnOk Then
                hasQueryWindows = True
                mstr��ʼ���� = .pstr��ʼ����
                mstr��ֹ���� = .pstr��ֹ����
                mstr�������� = .pstr��������
                mstr��λ���� = .pstr��λ����
                mstr���� = .pstr����
                mstr��쵥�� = .pstr��쵥��
                mstr�Թܱ�� = .pstr�Թܱ��
                mstrϵͳ��� = .pstrϵͳ���
                mstr���֤�� = .pstr���֤��
                '���²�ѯ��
               sub��ѯ����ʾ
            'add 20150504 �Ӳ�ѯ��������ѯ��ˢ��cgrdmian��Ȼ������м�¼��ֱ��ѡ�е�һ�У�Ȼ��ֵ��frmregister
            If cgrdMain.rows > 1 And mstr���֤�� <> "" Then
            'ֻҪ��ѯ���м�¼������ʾ  2015-11-17 by Ĳ��
            If cgrdMain.rows > 1 Then
                cgrdMain.Row = 1
                cmnuItemRegister_Click 5
                selectneednum = mstrϵͳ���
            Else
'             sub��ѯ����ʾ
              MsgBox "û�в�ѯ����Ϣ����ȷ�������Ϣ��", vbOKOnly + vbExclamation, "ϵͳ��ʾ"
              subȫ����ʾ
              Exit Sub
            End If
            End If
            End If
            
        End With
        
    Case 2 'ˢ��
        sub��ʾ��ѯ���
    Case 4
        Unload Me
    End Select
    Exit Sub
errHandler:
    sfsub������ "ְҵ������", "frmRegisterManage", "cmnuItemView_Click", Err.Number, Err.Description, False
End Sub



Private Sub coptType_Click(Index As Integer)
    On Error GoTo errHandler
    sub��ʾ��ѯ���
    If cgrdMain.rows > 1 Then
        cgrdMain.ColHidden(mcolIndex("����������")) = False
        cgrdMain.ColHidden(mcolIndex("����ϵͳ���")) = False
        cgrdMain.ColHidden(mcolIndex("���״̬")) = False
        If coptType(0).Value = True Then
            cgrdMain.ColHidden(mcolIndex("����������")) = True
            cgrdMain.ColHidden(mcolIndex("����ϵͳ���")) = True
            cgrdMain.ColHidden(mcolIndex("���״̬")) = True
        End If
    End If
    '2012-06-15 �ڵ�� ��
    'coptType�ؼ�����ȫ�����ģ���Ӧ����Ҳ���и���
    ctlb������.Buttons(4).Enabled = coptType(5).Value
'    cmnuItemRegister(3).Enabled = coptType(1).Value

    
    ctlb������.Buttons(6).Enabled = coptType(0).Value Or coptType(1).Value 'Or coptType(5).Value
    ctlb������.Buttons(7).Enabled = coptType(0).Value Or coptType(1).Value Or coptType(5).Value
    '2012-12-18 ������
    'Bug No:0000033
    '������������˵��༭����һ��
'    cmnuItemRegister(5).Enabled = coptType(0).Value Or coptType(1).Value
'    cmnuItemRegister(6).Enabled = coptType(0).Value Or coptType(1).Value Or coptType(5).Value
    '2012-12-18 ������
'    ctlb������.Buttons(3).Enabled = Not coptType(5).Value
   ctlb������.Buttons(3).Enabled = coptType(0).Value
'    cmnuItemRegister(1).Enabled = Not coptType(5).Value
    ctlb������.Buttons(10).Enabled = coptType(0).Value
    ctlb������.Buttons(12).Enabled = coptType(3).Value Or coptType(2).Value
'    ctlb������.Buttons(13).Enabled = coptType(1).Value Or coptType(2).Value
    ctlb������.Buttons(15).Enabled = coptType(0).Value
    '2012-06-15 �ڵ�� ��

    '2012-07-06 �ڵ�� ��
    '���Ӵ�ӡ��ǩ������ʹ���ж�
    'clblLabelNum.Visible = ctlb������.Buttons(13).Enabled
    'ctxtLabelNum.Visible = ctlb������.Buttons(13).Enabled
    '2012-07-06 �ڵ�� ��
    
    '2012-04-14 �ڵ�� ��
    '�˵���ɾ������ӡ���������(1)��������(2)�����������
    'cmnuItemPrint(1).Enabled = coptType(0).Value
    'cmnuItemPrint(2).Enabled = coptType(2).Value
    '2012-04-14 �ڵ�� ��
    
    Exit Sub
errHandler:
    sfsub������ "ְҵ������", "frmRegisterManage", "coptType_Click", Err.Number, Err.Description, False
End Sub

Private Sub ctxtLabelNum_LostFocus()
    Dim num As Integer
    printLabelNumLegal = True
    If IsNumeric(ctxtLabelNum.Text) = True Then
        num = CInt(ctxtLabelNum.Text)
        If num <> ctxtLabelNum.Text Then
            ctxtLabelNum.Text = "2"
            printLabelNumLegal = False
        Else
            If num <= 0 Then ctxtLabelNum.Text = "2": printLabelNumLegal = False
        End If
    Else
        ctxtLabelNum.Text = "2"
        printLabelNumLegal = False
    End If
End Sub

Private Sub Form_Load()
    Dim lcol��������ť As New Collection '�������ϵİ�ť��ʼ�����ϡ�
    
    On Error GoTo errHandler
    
    '��������Ѿ���ʼ�������ٽ��г�ʼ����
    If mblnInUse Then Exit Sub
    '��ʾ���ȡ�
    frmProcess.proPercent.Max = 8
    frmProcess.Label1.Caption = "���ڳ�ʼ�����棬��ȴ�..."
    frmProcess.proPercent.Value = 1
    frmProcess.Show
    DoEvents
    '���ô�������ʹ�õı�־��
    mblnInUse = True
    
    '��ʼ������ͨ�ö���ÿ�������Ӧһ������ͨ�ö���ʵ�������ɻ��ã��м��мǣ���
    Set mobjGUI = New cls����ͨ�ö���
    With lcol��������ť
        .Add "��ѯ(&Q)108"
        .Add "����(&A)103"
'        .Add "|"
        .Add "����(&R)101"  '3
        .Add "����Ǽ�(&F)103"
        .Add "|"
        .Add "�޸�" '6
        .Add "ɾ��" '7
        .Add "|"
        .Add "����(&O)113"  '9
        .Add "��λ����(&I)102"   '10
        .Add "|"
        .Add "��ӡ�嵥(&P)107"    '12
        .Add "��ӡ��ǩ(&U)107"    '13
        .Add "|"
        .Add "У��ͨ��(&J)106"    '15
        .Add "|"
        .Add "�˳�"
    End With
    '2012-12-18 ������
    'Ϊ��Ҫͨ������ͨ�ö�����Ƶĸ��ֿؼ�����ʼֵ��
    With mobjGUI
        Set .Form = Me
        Set .c������ = ctlb������
    End With
    frmProcess.proPercent.Value = 2
    DoEvents
    '���ý���ͨ�ö����ṩ�ķ������Խ���ؼ����г�ʼ����
    mobjGUI.subInitialize lcol��������ť, ""
    ctlb������.Buttons(13).Visible = True
    ctlb������.Buttons(15).Visible = False
    ctlb������.Buttons(16).Visible = False
'    ctlb������.Buttons(13).Enabled = False    '����"��ӡ��ǩ" 2015-11-19 by Ĳ��
    '2012-05-22 ���� ������ 2012-06-15 �ڵ�� ΢��Ȩ������
    '����Ȩ������
    Dim lobjTmp As Object
    Set lobjTmp = CreateObject("ְҵ������.clsPermissionConfigure")
    If lobjTmp.func���Ҳ���Ȩ��(um�û����, "ְҵ�����_���Ǽ�_����Ǽ�") = False Then
        ctlb������.Buttons(2).Visible = False
        ctlb������.Buttons(3).Visible = False
'        cmnuItemRegister(1).Visible = False
    End If
    If lobjTmp.func���Ҳ���Ȩ��(um�û����, "ְҵ�����_���Ǽ�_����Ǽ�") = False Then
        ctlb������.Buttons(4).Visible = False
        ctlb������.Buttons(5).Visible = False
        'cmnuItemRegister(3).Checked = False
    End If
    frmProcess.proPercent.Value = 3
    DoEvents
    If lobjTmp.func���Ҳ���Ȩ��(um�û����, "ְҵ�����_���Ǽ�_�޸�") = False Then
        ctlb������.Buttons(6).Visible = False
'        cmnuItemRegister(5).Visible = False
'        cmnuItemRegister(5).Checked = False
    End If
    If lobjTmp.func���Ҳ���Ȩ��(um�û����, "ְҵ�����_���Ǽ�_ɾ��") = False Then
        ctlb������.Buttons(7).Visible = False
        ctlb������.Buttons(8).Visible = False
'        cmnuItemRegister(6).Visible = False
'        cmnuItemRegister(6).Checked = False
    End If
    If lobjTmp.func���Ҳ���Ȩ��(um�û����, "ְҵ�����_���Ǽ�_����") = False Then
        ctlb������.Buttons(9).Visible = False
    End If
    frmProcess.proPercent.Value = 4
    DoEvents
    If lobjTmp.func���Ҳ���Ȩ��(um�û����, "ְҵ�����_���Ǽ�_��λ����") = False Then
        ctlb������.Buttons(10).Visible = False
    End If
    If lobjTmp.func���Ҳ���Ȩ��(um�û����, "ְҵ�����_���Ǽ�_��ӡ�嵥") = False Then
        ctlb������.Buttons(12).Visible = False
    End If
    If lobjTmp.func���Ҳ���Ȩ��(um�û����, "ְҵ�����_���Ǽ�_��ӡ�Թܱ�ǩ") = False Then
'        ctlb������.Buttons(13).Visible = False
        ctlb������.Buttons(14).Visible = False
    End If
    frmProcess.proPercent.Value = 5
    DoEvents
'''    If lobjTmp.func���Ҳ���Ȩ��(um�û����, "ְҵ�����_���Ǽ�_���Ǽ�") = False Then
'''        cmnuItemRegister(2).Checked = False
'''    End If
    Set lobjTmp = Nothing
    '2012-05-22 ������

    'ȱʡ��ʾ���һ�ܵ������Ա��
    '2012-06-15 �ڵ�� ��
    'Ĭ�ϲ�ѯ����ʱ����������Ա
    mstr��ʼ���� = Format(DateAdd("d", -30, Now), "yyyy-mm-dd")
'    mstr��ʼ���� = mstr��ʼ���� & " 00:00:00"
'    mstr��ʼ���� = Format(CDate("2000-01-01 00:00:00"), "yyyy-mm-dd hh:mm:ss")
    '2012-06-15 �ڵ�� ��
    mstr��ֹ���� = Format(Now, "yyyy-mm-dd")
'    mstr��ֹ���� = Format(Now, "yyyy-mm-dd hh:mm:ss")
    mstr�������� = ""
    mstr��λ���� = ""
    mstr���� = ""
    mstr��쵥�� = ""
    mstr�Թܱ�� = ""
    mstr���֤�� = ""
    frmProcess.proPercent.Value = 6
    DoEvents
    sub��ѯ����ʾ
    frmProcess.proPercent.Value = 7
    DoEvents
    ctlb������.Buttons(4).Enabled = coptType(1).Value
    'cmnuItemRegister(3).Enabled = coptType(1).Value
    
    ctlb������.Buttons(6).Enabled = coptType(0).Value
    ctlb������.Buttons(7).Enabled = coptType(0).Value
'    cmnuItemRegister(5).Enabled = coptType(0).Value
'    cmnuItemRegister(6).Enabled = coptType(0).Value

    '2012-04-14 �ڵ�� ��
    '�˵���ɾ������ӡ���������(1)��������(2)�����������
    'cmnuItemPrint(1).Enabled = coptType(0).Value
    'cmnuItemPrint(2).Enabled = coptType(2).Value
    '2012-04-14 �ڵ�� ��
 
    '2012-02-23 �ڵ�� ��
    cgrdMain.HighLight = flexHighlightWithFocus
    cgrdMain.SelectionMode = flexSelectionListBox
    cgrdMain.AllowBigSelection = False
    '2012-02-23 �ڵ�� ��
    
    '2012-06-20 �ڵ�� ��
    '��ʼ��һϵ�е�ѡ��
    coptType(0).Value = 1
    coptType_Click (0)
    '2012-06-20 �ڵ�� ��
    frmProcess.proPercent.Value = 8
    Unload frmProcess
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����沿��", "frmRegistermanage", "Form_Load", 6666, lstrError, False
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
    If hasQueryWindows = False Then mstr��ֹ���� = Format(Now, "yyyy-mm-dd")       'mstr��ֹ���� = Now
    Set mobjQueryResult = pobjҵ�����.func����������ѯ(mstr��ʼ����, mstr��ֹ����, mstr��������, mstr��λ����, mstr����, mstr��쵥��, mstr�Թܱ��, mstrϵͳ���, mstr���֤��)
    
    sub��ʾ��ѯ���
    
'    Dim i As Long
'    If cgrdMain.rows > 1 Then
'        Set mcolIndex = New Collection
'        For i = 0 To cgrdMain.cols - 1
'            mcolIndex.Add i, cgrdMain.TextMatrix(0, i)
'        Next
'        cgrdMain.ColHidden(mcolIndex("�Թܱ��")) = True
'        cgrdMain.ColHidden(mcolIndex("������")) = True
'    End If
    Exit Sub
errHandler:
    sfsub������ "ְҵ������", "frmRegisterManage", "sub��ѯ����ʾ", Err.Number, Err.Description, True
End Sub
Public Sub subȫ����ʾ()
    On Error GoTo errHandler
    If hasQueryWindows = False Then mstr��ֹ���� = Now
    Set mobjQueryResult = pobjҵ�����.func����������ѯ(mstr��ʼ����, mstr��ֹ����, "", "", "", "", "", "", "")
    
    sub��ʾ��ѯ���
    
'    Dim i As Long
'    If cgrdMain.rows > 1 Then
'        Set mcolIndex = New Collection
'        For i = 0 To cgrdMain.cols - 1
'            mcolIndex.Add i, cgrdMain.TextMatrix(0, i)
'        Next
'        cgrdMain.ColHidden(mcolIndex("�Թܱ��")) = True
'        cgrdMain.ColHidden(mcolIndex("������")) = True
'    End If
    Exit Sub
errHandler:
    sfsub������ "ְҵ������", "frmRegisterManage", "sub��ѯ����ʾ", Err.Number, Err.Description, True
End Sub

Private Sub sub��ʾ��ѯ���()
    Dim i As Integer
    On Error GoTo errHandler
    
    '2012-06-15 �ڵ�� ��
    '�������״̬
'''    If coptType(0).Value Then
'''        mobjQueryResult.Filter = "���״̬='δУ��'"
'''    ElseIf coptType(1).Value Then
'''        mobjQueryResult.Filter = "���״̬='���½���' and ����������<>'' and ����ϵͳ���=''"
'''    Else
'''        mobjQueryResult.Filter = "(���״̬='���½���' and  ����������='') or (���״̬='���½���' and ����ϵͳ���<>'')"
'''    End If
    If coptType(0).Value Then
        mobjQueryResult.Filter = "���״̬='δУ��'or ���״̬='δ���嵥'"
'    ElseIf coptType(1).Value Then
'        mobjQueryResult.Filter = "���״̬='δ���嵥'"
    ElseIf coptType(2).Value Then
        mobjQueryResult.Filter = "���״̬='�����' or ���״̬='δ¼���ܼ��߸�����Ϣ'"
    ElseIf coptType(3).Value Then
        mobjQueryResult.Filter = "���״̬='δ�½���'"
    ElseIf coptType(4).Value Then
       ' mobjQueryResult.Filter = "���״̬='δ����' or ���״̬='�Ѹ���'"
       '�޸ģ������ 2012-12-7 ��
       '�������״̬�Ѿ��½��۵���Ϣ
       'bug�ţ�000072
       mobjQueryResult.Filter = "���״̬='������' or ���״̬='�Ѹ���' or ���״̬='�ѷ�����'"
         '�޸ģ������ 2012-12-7 ��
    ElseIf coptType(5).Value Then
'        mobjQueryResult.Filter = "���״̬='������'"
        '����״̬�����ǣ�2012-10-30
        Dim lobjRec As Object
        dasubSetQueryTimeout 600
        Dim lstrSql As String
        lstrSql = "select a.ϵͳ���,����,�Ա�,����,��λ����,Σ������,�ֹ���, ������ as ������,b.������ݺ��� ," _
                    & "convert(varchar(10),�������,120) as �������, " _
                    & "isnull(����������,'') as ����������, " _
                    & "isnull(����ϵͳ���,'') as ����ϵͳ���,a.����ԭ��,a.������Ŀ," _
                    & "���״̬= '������' from ְҵ�����_��������Ϣ�� a inner join  ְҵ�����_�����Ա������Ϣ�� b on a.ϵͳ���=b.ϵͳ��� and ����״̬ = '0'"
        Set lobjRec = dafuncGetData(lstrSql)
        cgrdMain.rows = 1
        
        If Not lobjRec.EOF Then
            With cgrdMain
                Set .DataSource = lobjRec
                If cgrdMain.rows > 1 Then
                    Set mcolIndex = New Collection
                    For i = 0 To cgrdMain.cols - 1
                        mcolIndex.Add i, cgrdMain.TextMatrix(0, i)
                    Next
                End If
                clblInfo = .rows - 1
                .Col = 0
'                .Sort = flexSortGenericDescending
                .AutoSize 0, .cols - 1, 0, 0
                .ExplorerBar = flexExSort
'                .DataMode = flexDMFree
                clblInfo = .rows - 1
            End With
            
            Exit Sub
        Else
            cgrdMain.rows = 1
            clblInfo = cgrdMain.rows - 1
            Exit Sub
        End If
    End If
    '2012-06-15 �ڵ�� ��
    
    With cgrdMain
        Set .DataSource = mobjQueryResult
        If cgrdMain.rows > 1 Then
            Set mcolIndex = New Collection
            For i = 0 To cgrdMain.cols - 1
                mcolIndex.Add i, cgrdMain.TextMatrix(0, i)
            Next
            cgrdMain.ColHidden(mcolIndex("�Թܱ��")) = True
'        '���û�鵽�ˣ���ʾ��ѯ��������  2016-1-6 by Ĳ��
'        Else
'            MsgBox "��ѯ��������", vbOKOnly
        End If
        clblInfo = .rows - 1
        .Col = 0
'        .Sort = flexSortGenericDescending
        '2012-06-15 �ڵ�� ��
        'vsflexgrid�п�Ȱ������Զ������������ͷ����ͷ����������
        .AutoSize 0, .cols - 1, 0, 0
        .ExplorerBar = flexExSort
'        .DataMode = flexDMFree
        '2012-05-15 �ڵ�� ��
        
    End With

        If cgrdMain.rows > 1 Then
'        Dim i As Long
            Set mcolIndex = New Collection
            For i = 0 To cgrdMain.cols - 1
                mcolIndex.Add i, cgrdMain.TextMatrix(0, i)
            Next
            cgrdMain.ColHidden(mcolIndex("�Թܱ��")) = True
            cgrdMain.ColHidden(mcolIndex("������")) = True
        End If
    
    
    Exit Sub
errHandler:
    sfsub������ "ְҵ������", "frmRegisterManage", "sub��ʾ��ѯ���", Err.Number, Err.Description, True

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    cgrdMain.Width = Me.ScaleWidth - cgrdMain.Left - 60
    cgrdMain.Height = Me.ScaleHeight - cgrdMain.Top - 60
    
End Sub

Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    On Error GoTo errHandler
    
    Dim lobjFile As Object
    Dim i As Integer
    
    Select Case Operate
    Case "��ѯ"
        cmnuItemView_Click 1
    
    Case "����"
        frmAddRegister.Show 1
   
    Case "����"
'    cmnuItemRegister_Click 1   '�����Ǽ�
        cmnuItemRegister_Click 8
        Cancel = True
        
    Case "�޸�"
        Cancel = True
        cmnuItemRegister_Click 5
    
    Case "ɾ��"
        Cancel = True
        cmnuItemRegister_Click 6
    Case "����"
        If cgrdMain.rows <= 1 Then
            MsgBox "û����Ҫ�����ļ�¼��", vbOKOnly + vbExclamation, "ϵͳ��ʾ"
            Exit Sub
        End If
        
        Dim lstrFile As String
        ccmdFile.Filter = "Excel�ļ� (*.xls)|*.xls|�ı��ļ� (*.txt)|*.txt"
        ccmdFile.ShowSave
        lstrFile = ccmdFile.FileName
        If lstrFile <> "" Then
            '2012-04-14 �ڵ�� ��
            '��Ϊ��0�У�Ϊϵͳ��š��������б���ʱΪstring
            cgrdMain.ColDataType(0) = flexDTString
            cgrdMain.SaveGrid lstrFile, flexFileExcel, True   '����excelϵͳ���Ϊ����
            'cgrdMain.SaveGrid lstrFile, flexFileTabText, True
            '2012-04-14 �ڵ���
        End If
    
    '2012-06-15 �ڵ�� �� ��ӡ��������ȫ��ȡ��
'''    '2012-02-13 �ڵ�� ��
'''    '�޸����ݣ�������ӡ���봰��
'''    Case "��ӡ������"
'''        FrmPrintBarCode.Show
'''    Case "���´�����"
'''        FrmPrintBarCodeAgain.Show
'''    '2012-02-13 �ڵ�� ��
    '2012-06-15 �ڵ�� ��
    
    '2012-02-16 ��
    Case "��λ����"
        frmImportExcel.Show
    '2012-02-16 �ڵ�� ��
'        sub��ѯ����ʾ
    '2012-06-15 �ڵ�� �� ���ģ�2012-10-17 �����
    Case "��ӡ�嵥"
        '���ô�ӡ�������������2012-10-25
        Dim j As Integer
        Dim paraϵͳ��� As Collection
        Set paraϵͳ��� = New Collection

        Set lobjFile = CreateObject("ְҵ������.cls����")
        For j = 0 To cgrdMain.SelectedRows - 1
             paraϵͳ���.Add (cgrdMain.TextMatrix(cgrdMain.SelectedRow(j), mcolIndex("ϵͳ���")))
        Next j


'        Dim j As Integer
'        Dim paraϵͳ��� As Collection
'        Set paraϵͳ��� = New Collection
'
'        Set lobjFile = CreateObject("ְҵ������.cls����")
'             For j = 1 To cgrdMain.SelectedRows
'             paraϵͳ���.Add (cgrdMain.TextMatrix(j, mcolIndex("ϵͳ���")))
'            Next j
        lobjFile.func��ӡ����嵥 paraϵͳ���

'             lobjFile.func��ӡ����嵥(cgrdMain.TextMatrix(cgrdMain.SelectedRows, mcolIndex("ϵͳ���")))
        Set lobjFile = Nothing

        '���ĵ�ǰ���״̬����ӡ�嵥֮�󣬾ͽ������״̬��
        For j = 0 To cgrdMain.SelectedRows - 1
             pobjҵ�����.funcд�뵥�˵�ǰ���״̬ cgrdMain.TextMatrix(cgrdMain.SelectedRow(j), mcolIndex("ϵͳ���")), 2
        Next j
        sub��ѯ����ʾ
        
    '2012-06-15 �ڵ�� �� ���ģ�2012-10-17  �����
    Case "��ӡ��ǩ"
'         Dim j As Integer
        '���ƴ�ӡ��ǩ����
        frmAssayDeptSelect.Show 1
        Dim paraSelectedDeptName As Collection
        If cgrdMain.SelectedRows < 1 Then Exit Sub
        If frmAssayDeptSelect.pblnOk Then
            If frmAssayDeptSelect.selectedDeptName.Count > 1 Then
                Set paraSelectedDeptName = frmAssayDeptSelect.selectedDeptName
                For j = 0 To cgrdMain.SelectedRows - 1
                    For i = 2 To paraSelectedDeptName.Count
                        Set lobjFile = CreateObject("ְҵ������.cls����")
'                       lobjFile.func��ӡ�Թܱ�ǩ cgrdMain.TextMatrix(cgrdMain.SelectedRow(j), mcolIndex("ϵͳ���")), cgrdMain.TextMatrix(cgrdMain.SelectedRow(j), mcolIndex("����")), paraSelectedDeptName.Item(i)
                        lobjFile.func��ӡ�Թܱ�ǩ cgrdMain.TextMatrix(cgrdMain.SelectedRow(j), mcolIndex("ϵͳ���")), cgrdMain.TextMatrix(cgrdMain.SelectedRow(j), mcolIndex("����")), cgrdMain.TextMatrix(cgrdMain.SelectedRow(j), mcolIndex("�Ա�")), cgrdMain.TextMatrix(cgrdMain.SelectedRow(j), mcolIndex("����")), paraSelectedDeptName.Item(i)
                    Next i
                Next j
            End If
        Else
            Exit Sub
        End If
        
'''        '���ƴ�ӡ��ǩ����
'''        ctxtLabelNum_LostFocus  '������ת�ƣ�ִ��ctxtLabelNum_lostfocus�ж�ֵ�Ƿ�Ϸ�
'''        If printLabelNumLegal = False Then Exit Sub
'''        Dim printTimes As Integer
'''        printTimes = CInt(ctxtLabelNum.Text)
'''        Set lobjFile = CreateObject("ְҵ������.cls����")
'''        While printTimes > 0
'''            lobjFile.func��ӡ�Թܱ�ǩ cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("ϵͳ���")), cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("����"))
'''            printTimes = printTimes - 1
'''        Wend
        
    Case "У��ͨ��"
        If cgrdMain.Row < 1 Then
            MsgBox "û����ҪУ�˵ļ�¼��", vbOKOnly + vbExclamation, "ϵͳ��ʾ"
            Exit Sub
        End If
        
        mintState = 1
        For i = 0 To cgrdMain.SelectedRows - 1
            pobjҵ�����.funcд�뵥�˵�ǰ���״̬ cgrdMain.TextMatrix(cgrdMain.SelectedRow(i), mcolIndex("ϵͳ���")), mintState
            pobjҵ�����.funcд��У������Ϣ cgrdMain.TextMatrix(cgrdMain.SelectedRow(i), mcolIndex("ϵͳ���")), um�û����
        Next
        sub��ѯ����ʾ
    '2012-08-17 �ڵ�� ��
    '���Ӹ���Ǽǹ���
    Case "����Ǽ�"
        cmnuItemRegister_Click 7
    '2012-08-17 �ڵ�� ��
    End Select
    Exit Sub
errHandler:
    sfsub������ "ְҵ������", "frmRegisterManage", "mobjGUI_BeforeOperate", Err.Number, Err.Description, False
    Cancel = True
    Exit Sub
    Resume
End Sub

