VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F48D4DEC-1198-11D5-91BE-0050BA06B70C}#5.9#0"; "¼��ؼ�.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmCompany 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "��λ����"
   ClientHeight    =   7350
   ClientLeft      =   240
   ClientTop       =   375
   ClientWidth     =   10050
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7724.65
   ScaleMode       =   0  'User
   ScaleWidth      =   10050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton Command1 
      Caption         =   "ѡ��"
      Height          =   375
      Left            =   7920
      TabIndex        =   37
      Top             =   600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox ctxt��� 
      Height          =   270
      Left            =   8160
      TabIndex        =   36
      Top             =   600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CheckBox copt�������� 
      Caption         =   "��������"
      Height          =   255
      Left            =   6360
      TabIndex        =   35
      Top             =   600
      Width           =   1335
   End
   Begin VSFlex8Ctl.VSFlexGrid cgrdMain 
      Height          =   4455
      Left            =   120
      TabIndex        =   32
      Top             =   2760
      Width           =   9735
      _cx             =   2088780563
      _cy             =   2088771250
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
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   1
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
      AutoSearchDelay =   2
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
   Begin VB.Frame Frame3 
      Caption         =   "  �ֵ�λ��Ϣ¼��   "
      ForeColor       =   &H000080FF&
      Height          =   1815
      Left            =   120
      TabIndex        =   20
      Top             =   960
      Width           =   9735
      Begin VB.TextBox Text1 
         BorderStyle     =   0  'None
         ForeColor       =   &H000000C0&
         Height          =   1335
         Left            =   5880
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   34
         Top             =   240
         Width           =   3615
      End
      Begin VB.TextBox ctxt��λ���� 
         Height          =   300
         Left            =   1440
         TabIndex        =   33
         Top             =   240
         Width           =   4335
      End
      Begin VB.TextBox ctxt������ 
         Height          =   300
         Left            =   1440
         TabIndex        =   25
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox ctxt��ϵ�绰 
         Height          =   300
         Left            =   4200
         TabIndex        =   24
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox ctxt��λ��ַ 
         Height          =   300
         Left            =   1440
         TabIndex        =   23
         Top             =   1320
         Width           =   4335
      End
      Begin VB.ComboBox ccmb�������� 
         Height          =   300
         Left            =   4200
         TabIndex        =   22
         Text            =   "ccmb��������"
         Top             =   960
         Width           =   1575
      End
      Begin VB.ComboBox Ccmb��ҵ��� 
         Height          =   300
         Left            =   1440
         TabIndex        =   21
         Text            =   "Ccmb��ҵ���"
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ���ƣ�"
         Height          =   180
         Index           =   5
         Left            =   480
         TabIndex        =   31
         Top             =   360
         Width           =   900
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "�����ˣ�"
         Height          =   180
         Left            =   600
         TabIndex        =   30
         Top             =   720
         Width           =   720
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "��ϵ�绰��"
         Height          =   180
         Left            =   3240
         TabIndex        =   29
         Top             =   720
         Width           =   900
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "��λ���ַ��"
         Height          =   180
         Left            =   240
         TabIndex        =   28
         Top             =   1440
         Width           =   1080
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "�������ʣ�"
         Height          =   180
         Left            =   3240
         TabIndex        =   27
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "��ҵ���"
         Height          =   180
         Left            =   480
         TabIndex        =   26
         Top             =   1080
         Width           =   900
      End
   End
   Begin VB.Frame cfram������Ϣ 
      Caption         =   "�Ǽǻ�����Ϣ���ǿ���¼��ʱ��ɫΪ��¼�����¼��ʱֻ������):"
      ClipControls    =   0   'False
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   2520
      TabIndex        =   5
      Top             =   9360
      Width           =   6300
      Begin VB.TextBox ctxt���� 
         Height          =   300
         Left            =   4800
         TabIndex        =   18
         Top             =   2040
         Width           =   1695
      End
      Begin VB.ComboBox ccmb���ʱ�� 
         Height          =   300
         Left            =   8160
         TabIndex        =   16
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox ctxt���� 
         Height          =   270
         Left            =   4440
         TabIndex        =   14
         Text            =   "1"
         Top             =   2880
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox ctxt��쵥�� 
         Height          =   315
         Left            =   4200
         MaxLength       =   20
         TabIndex        =   1
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox ctxtTubeNo 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6000
         TabIndex        =   0
         Top             =   1320
         Width           =   1575
      End
      Begin VB.VScrollBar cvscLetter 
         Height          =   345
         Left            =   6360
         TabIndex        =   3
         Top             =   1320
         Width           =   345
      End
      Begin ¼��ؼ�.ctlInputDictGrid c�ֵ�� 
         Height          =   3255
         Left            =   360
         TabIndex        =   9
         Top             =   240
         Visible         =   0   'False
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   5741
         Cols            =   10
         Count           =   0
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
      Begin ¼��ؼ�.ctlInputFrame ciptBase 
         Height          =   975
         Left            =   6120
         TabIndex        =   2
         Top             =   2280
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   1720
         BackColor       =   15791081
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackStyle       =   1
         Caption         =   ""
         Rows            =   1
         Cols            =   27
         DistanceofRow   =   0
         AutoSize        =   0   'False
         FormatString    =   "���֤��,1,0,12"
         Count           =   1
         titleInputBox0001=   "���֤��"
         statusinfoInputBox0001=   ""
         lengthInputBox0001=   12
         orderInputBox0001=   1
         valueInputBox0001=   ""
         datatypeInputBox0001=   3
         colInputBox0001 =   0
         rowInputBox0001 =   1
         PassWordCharInputBox0001=   0   'False
         ����InputBox0001=   0   'False
         ����������ֵInputBox0001=   0   'False
         ���������СֵInputBox0001=   0   'False
         �ֵ�����InputBox0001=   ""
         ��ʾ�ֵ��ֶ�InputBox0001=   ""
         �����ֵ��ֶ�InputBox0001=   ""
         ����InputBox0001=   "����� 1"
         ȱʡֵInputBox0001=   ""
         ����ȱʡֵInputBox0001=   ""
         ����InputBox0001=   0
         MaxInputBox0001 =   ""
         MinInputBox0001 =   ""
         VisibleInputBox0001=   -1  'True
         PermitNullInputBox0001=   -1  'True
         TriggerstrInputBox0001=   ""
         �����ѡInputBox0001=   0   'False
         ErrColor        =   12648447
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "�뽫�������֤���ڶ������ϣ�"
         Height          =   180
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   2520
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "���壺"
         Height          =   180
         Left            =   4800
         TabIndex        =   17
         Top             =   1800
         Width           =   540
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "���ʱ�ڣ�"
         Height          =   180
         Left            =   8160
         TabIndex        =   15
         Top             =   480
         Width           =   900
      End
      Begin VB.Label clbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   3600
         TabIndex        =   13
         Top             =   2880
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��쵥�ţ�"
         Height          =   180
         Index           =   7
         Left            =   4200
         TabIndex        =   12
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label clbl��������� 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   8760
         TabIndex        =   11
         Top             =   2760
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ϴ�������ڣ�"
         Height          =   180
         Index           =   4
         Left            =   8640
         TabIndex        =   10
         Top             =   2520
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Label clblTubeNo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "������뿴״̬��"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   6600
         TabIndex        =   8
         Top             =   1320
         Width           =   915
      End
      Begin VB.Label clblLetter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   6000
         TabIndex        =   7
         Top             =   1320
         Width           =   345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Թܱ�ţ�"
         Height          =   180
         Index           =   1
         Left            =   6000
         TabIndex        =   6
         Top             =   1080
         Width           =   900
      End
   End
   Begin MSComctlLib.Toolbar ctbMain 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   10050
      _ExtentX        =   17727
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
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
      Left            =   1320
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'���ܣ�������λ��������ʱ�䣬����С������û�зֿ���ȫ��д��һ���ˡ���
'���ߣ�����
'���ڣ�2013.02.28

Option Explicit
Private mblnInUse As Boolean
Private mcolIndex As Collection
Private WithEvents mobjGUI As cls����ͨ�ö���
Attribute mobjGUI.VB_VarHelpID = -1

'������񣬼����޸İ�ť��
Private Sub cgrdMain_Click()
    ctbMain.Buttons(5).Enabled = True
    
End Sub

'˫�������䵼�����ġ���λ���ơ�
Private Sub cgrdMain_DblClick()
    Dim lstrSysNo As String
    Dim lobjRec As Object
    lstrSysNo = cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("������"))
    Set lobjRec = dafuncGetData("select ��λ���� from ��λ����_��λ������Ϣ�� where ������='" & lstrSysNo & "'")
    frmImportExcel.ctxt��λ���� = lobjRec(0)
    Unload Me
End Sub

Private Sub Command1_Click()
    frmAddRegister.ccmbUnit.Text = cgrdMain.TextMatrix(cgrdMain.Row, 1)
    Unload Me
End Sub

Private Sub Form_Activate()
    On Error Resume Next
End Sub

Private Sub Form_Deactivate()
    On Error Resume Next
End Sub
'��ʼ������
Private Sub Form_Load()
    Dim lobjRec As Object
    Dim lobjDetl As Object
    Dim i As Integer
    On Error GoTo errHandler
   
    '��������Ѿ���ʼ�������ٽ��г�ʼ����
    If mblnInUse Then Exit Sub
    
    '���ô�������ʹ�õı�־��
    mblnInUse = True
    
    MousePointer = 0
    
    '��ʼ������ͨ�ö���ÿ�������Ӧһ������ͨ�ö���ʵ�������ɻ��ã��м��мǣ���
    Set mobjGUI = New cls����ͨ�ö���
    mobjGUI.pbln�Զ������ֵ�߶� = False
    
    '���ù�����������Ҫ�ĸ��ְ�ť��
    Dim lcol��������ť As New Collection           '�������ϵİ�ť��ʼ�����ϡ�
    With lcol��������ť
        .Add "��ѯ(&C)108"
        .Add "|"
        .Add "����(&T)101"
        .Add "|"
        .Add "�޸�"
        .Add "|"
        .Add "ɾ��"
        .Add "|"
        .Add "�˳�"
    End With
    
    'Ϊ��Ҫͨ������ͨ�ö�����Ƶĸ��ֿؼ�����ʼֵ��
    With mobjGUI
        Set .Form = Me
        Set .c������ = ctbMain
        Set .c¼��� = ciptBase
        Set .c�ֵ�� = c�ֵ��
        Set .c״̬�� = ctbMain
        
        '���ý���ͨ�ö����ṩ�ķ������Խ���ؼ����г�ʼ����
        .subInitialize lcol��������ť, ""
    End With
    Text1.Text = "����˵����1����ӣ��ڡ����е�λ��Ϣ¼��" _
            & "����¼����Ϣ�Ժ󣬵�����漴�ɡ�2���޸�" _
            & "����������һ����Ҫ�޸ĵ����ݣ�����޸�" _
            & "��Ȼ���޸������ݣ���󱣴漴�ɡ�3��ɾ��" _
            & ": ѡ����Ҫɾ�������ݵ��ɾ������?4?��" _
            & "ѯ�����ֵ�λ��Ϣ¼����¼�����ݣ������ѯ" _
            & "���ɣ�Ϊ�ձ�ʾ����������ѯ����"
    
    SubSelect
    subClear
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ " ��λ��������", "FrmCompany", "Form_Load", 6666, lstrError, False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    mblnInUse = False
    
    Set mobjGUI = Nothing
End Sub

'���ܣ����������ϰ�ť��
Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    On Error GoTo errHandler
    Dim lstrError As String
    Dim lstrSysNo As String
    
    SubSelect
'    lstrSysNo = cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("������"))
'    lstrSysNo = ""
    Select Case Operate
    Case "��ѯ"
        subQuery
    Case "����"
        If ctxt��λ���� = "" Then
            MsgBox "��λ���Ʋ���Ϊ�գ�"
            Exit Sub
        End If
        
        '���������š�
        subSave (ctxt���)
        SubSelect

        lstrSysNo = ""
        subClear
        If ctbMain.Buttons(5).Enabled = False Then
            ctbMain.Buttons(5).Enabled = True
        End If
    Case "�޸�"
        ctxt��� = cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("������"))
        lstrSysNo = ctxt���
        subUpdata lstrSysNo
        SubSelect
'        ctxt��� = ""
        ctbMain.Buttons(5).Enabled = False
    Case "ɾ��"
        ctxt��� = cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("������"))
        lstrSysNo = ctxt���
        If cgrdMain.Row = 0 Or cgrdMain.Row > cgrdMain.rows - 1 Then
            MsgBox "��ѡ����Ҫɾ�������ݣ�"
        Else
            If MsgBox("ȷ��Ҫɾ���õ�λ��Ϣ��", vbYesNo, "ϵͳ��ʾ") = vbYes Then
                subDelete lstrSysNo
            End If
        End If
        SubSelect
        ctxt��� = ""
        lstrSysNo = ""
    Case "�˳�"
        Unload Me
    End Select
    Exit Sub
errHandler:
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "��λ��������", "FrmCompany", "mobjGUI_BeforeOperate", 6666, lstrError, False
End Sub

 
'����
Private Sub subSave(ByVal paraNo As String)
    Dim lobjRec As Object
    Dim lstrSql As String
    Dim mstrSQNo As String
    Set lobjRec = dafuncGetData("select * from ��λ����_��λ������Ϣ�� where ������='" & paraNo & "'")
    If (lobjRec.BOF Or lobjRec.EOF) Then
    
        mstrSQNo = funcȡ�µ�������(pstr����վ����)
        lstrSql = " insert into ��λ����_��λ������Ϣ��(������,��λ����,������,�绰,��������,��ҵ���,��ַ) values('" & mstrSQNo & "','" & Trim(ctxt��λ����) & "" _
                            & "','" & Trim(ctxt������) & "','" & Trim(ctxt��ϵ�绰) & "','" & Trim(ccmb��������) & "','" & Trim(Ccmb��ҵ���) & "','" & Trim(ctxt��λ��ַ) & "')"
        dafuncGetData lstrSql
    Else
        lstrSql = "update ��λ����_��λ������Ϣ�� set ��λ����='" & Trim(ctxt��λ����) & "',������='" & Trim(ctxt������) & "',�绰='" & Trim(ctxt��ϵ�绰) & "" _
                                    & "',��������='" & Trim(ccmb��������) & "',��ҵ���='" & Trim(Ccmb��ҵ���) & "',��ַ='" & Trim(ctxt��λ��ַ) & "" _
                                    & "' where ������='" & paraNo & "'"
        dafuncGetData lstrSql
    End If
    If copt��������.Value = True Then
        subClear
    End If
End Sub

'��ѯ����ʾ
Private Sub SubSelect()
    Dim lobjRec As Object
    Dim i As Integer
    Set lobjRec = dafuncGetData("select ������,��λ����,������,�绰,��������,��ҵ���,��ַ from ��λ����_��λ������Ϣ��")
    Set cgrdMain.DataSource = lobjRec
    Set mcolIndex = New Collection
    For i = 0 To cgrdMain.cols - 1
        mcolIndex.Add i, cgrdMain.TextMatrix(0, i)
    Next
    cgrdMain.AutoSize 0, cgrdMain.cols - 1, 0, 0
    cgrdMain.ColHidden(mcolIndex("������")) = True
End Sub

'��ȡ������
Public Function funcȡ�µ�������(ByVal para�û���� As String) As String
    On Error GoTo errHandler
    Dim mtempSql As String
    Dim lrstTemp As Object
    'ע�⣺�����ݷ��ʶ���δ�ṩ�Դ洢����ʹ�ò������ؽ���ķ������ڴ˱����޸�Դ�洢���̣����ü�¼�����������ź͵������
    mtempSql = "exec ��λ����_���������� "
    Set lrstTemp = dafuncGetData(mtempSql)
    'lrstTemp.Open
    If lrstTemp.RecordCount <> 0 Then
        funcȡ�µ������� = lrstTemp.Fields("������")
    Else
        funcȡ�µ������� = ""
    End If
    Exit Function
errHandler:
    sfsub������ "��λ����ҵ�����", "ClsManageUnitFile", "funcȡ�µ�������", Err.Number, Err.Description, True
End Function

'�޸�
Private Sub subUpdata(ByVal paraSysNo As String)
    Dim lobjRec As Object
    Dim lstrSql As String
    Set lobjRec = dafuncGetData("select * from ��λ����_��λ������Ϣ�� where ������='" & paraSysNo & "'")
    If lobjRec.RecordCount = 0 Then Exit Sub
    ctxt��λ����.Text = IIf(IsNull(lobjRec("��λ����")), "", lobjRec("��λ����"))
    ctxt������.Text = IIf(IsNull(lobjRec("������")), "", lobjRec("������"))
    ctxt��ϵ�绰.Text = IIf(IsNull(lobjRec("�绰")), "", lobjRec("�绰"))
    ccmb��������.Text = IIf(IsNull(lobjRec("��������")), "", lobjRec("��������"))
    Ccmb��ҵ���.Text = IIf(IsNull(lobjRec("��ҵ���")), "", lobjRec("��ҵ���"))
    ctxt��λ��ַ.Text = IIf(IsNull(lobjRec("��ַ")), "", lobjRec("��ַ"))
End Sub

'ɾ��
Private Sub subDelete(ByVal paraSysNo As String)
    dafuncGetData ("delete ��λ����_��λ������Ϣ�� where ������='" & paraSysNo & "'")
End Sub

'�������ս���
Private Sub subClear()
    Dim lobjRec As Object
    Dim lobjDetl As Object
    Dim i As Integer
    ctxt��λ���� = ""
    ctxt������ = ""
    ctxt��ϵ�绰 = ""
    ctxt��λ��ַ = ""
    ctxt��� = ""
    '��ȡ��������
    Set lobjRec = pobjDict.FetchEx("���������ֵ�")
    ccmb��������.Clear
    ccmb��������.AddItem ""
    For i = 1 To lobjRec.RecordCount
        ccmb��������.AddItem lobjRec("����")
        ccmb��������.ItemData(ccmb��������.NewIndex) = lobjRec("���")
        lobjRec.MoveNext
    Next
    ccmb��������.ListIndex = 0
    Set lobjRec = CreateObject("ְҵ������.clsmedicalexamtemplateset")
    Set lobjDetl = lobjRec.��ҵ���
    Ccmb��ҵ���.Clear
    Ccmb��ҵ���.AddItem ""
    For i = 1 To lobjDetl.RecordCount
        Ccmb��ҵ���.AddItem lobjDetl("����")
        Ccmb��ҵ���.ItemData(Ccmb��ҵ���.NewIndex) = lobjDetl("���")
        lobjDetl.MoveNext
    Next
    Ccmb��ҵ���.ListIndex = 0
End Sub

'��������ѯ
Private Sub subQuery()
    Dim lstrSql As String
    Dim lobjRec As Object
    Dim i As Integer
    lstrSql = "select ������,��λ����,������,�绰,��������,��ҵ���,��ַ from ��λ����_��λ������Ϣ�� where 1=1"
    If ctxt��λ���� <> "" Then
        lstrSql = lstrSql & " and ��λ����='" & Trim(ctxt��λ����) & "'"
    End If
    If ctxt������ <> "" Then
        lstrSql = lstrSql & " and ������='" & Trim(ctxt������) & "'"
    End If
    If ctxt��ϵ�绰 <> "" Then
        lstrSql = lstrSql & " and �绰='" & Trim(ctxt��ϵ�绰) & "'"
    End If
    If ccmb�������� <> "" Then
        lstrSql = lstrSql & " and ��������='" & Trim(ccmb��������) & "'"
    End If
    If Ccmb��ҵ��� <> "" Then
        lstrSql = lstrSql & " and ��ҵ���='" & Trim(Ccmb��ҵ���) & "'"
    End If
    If ctxt��λ��ַ <> "" Then
        lstrSql = lstrSql & " and ��ַ='" & Trim(ctxt��λ��ַ) & "'"
    End If
    
    Set lobjRec = dafuncGetData(lstrSql)
    
    Set cgrdMain.DataSource = lobjRec
    Set mcolIndex = New Collection
    For i = 0 To cgrdMain.cols - 1
        mcolIndex.Add i, cgrdMain.TextMatrix(0, i)
    Next
    cgrdMain.AutoSize 0, cgrdMain.cols - 1, 0, 0
    cgrdMain.ColHidden(mcolIndex("������")) = True
End Sub
