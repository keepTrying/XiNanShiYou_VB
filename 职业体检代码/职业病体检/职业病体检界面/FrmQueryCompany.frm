VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmQueryCompany 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ְҵ�������-��λͳ��"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   12660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   12660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin MSComDlg.CommonDialog ccmdFile 
      Left            =   240
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   6148
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   11895
      Begin VSFlex8Ctl.VSFlexGrid cgrdList 
         Height          =   4965
         Left            =   0
         TabIndex        =   12
         Top             =   840
         Width           =   11895
         _cx             =   20981
         _cy             =   8758
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
      Begin VB.Timer Timer1 
         Left            =   6960
         Top             =   0
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   0
         TabIndex        =   1
         Top             =   -120
         Width           =   11895
         Begin VB.Frame Frame4 
            Caption         =   "��ѯ����"
            Height          =   735
            Left            =   0
            TabIndex        =   2
            Top             =   120
            Width           =   11775
            Begin VB.CommandButton ccmdQuery 
               Caption         =   "��ѯ"
               Height          =   375
               Left            =   10320
               TabIndex        =   14
               Top             =   240
               Width           =   1095
            End
            Begin VB.CheckBox flag���� 
               Caption         =   "Check1"
               Enabled         =   0   'False
               Height          =   255
               Left            =   120
               TabIndex        =   6
               Top             =   360
               Value           =   1  'Checked
               Width           =   255
            End
            Begin VB.CheckBox flag���� 
               Caption         =   "Check1"
               Enabled         =   0   'False
               Height          =   375
               Left            =   4440
               TabIndex        =   5
               Top             =   240
               Value           =   1  'Checked
               Width           =   255
            End
            Begin VB.TextBox ctxtCompanyName 
               Height          =   375
               Left            =   1080
               TabIndex        =   4
               Top             =   240
               Width           =   2175
            End
            Begin VB.CommandButton CmdFCompany 
               Caption         =   "��λ��λ"
               Height          =   375
               Left            =   3360
               TabIndex        =   3
               Top             =   240
               Width           =   975
            End
            Begin MSComCtl2.DTPicker DTP��ֹʱ�� 
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "yyyy-MM-dd"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   3
               EndProperty
               Height          =   375
               Left            =   8040
               TabIndex        =   7
               Top             =   240
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   661
               _Version        =   393216
               Format          =   60489728
               CurrentDate     =   41027
            End
            Begin MSComCtl2.DTPicker DTP��ʼʱ�� 
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "yyyy-MM-dd"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   3
               EndProperty
               Height          =   375
               Left            =   5520
               TabIndex        =   8
               Top             =   240
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   661
               _Version        =   393216
               Format          =   60489728
               CurrentDate     =   41027
            End
            Begin VB.Label Label3 
               Caption         =   "��"
               Height          =   255
               Left            =   7800
               TabIndex        =   11
               Top             =   360
               Width           =   255
            End
            Begin VB.Label Label2 
               Caption         =   "���ʱ��"
               Height          =   255
               Left            =   4680
               TabIndex        =   10
               Top             =   345
               Width           =   855
            End
            Begin VB.Label Label1 
               Caption         =   "��λ����"
               Height          =   255
               Left            =   360
               TabIndex        =   9
               Top             =   360
               Width           =   855
            End
         End
      End
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   5565
      Top             =   4485
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmQueryCompany.frx":0000
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmQueryCompany.frx":005E
            Key             =   "Back"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar ctlb������ 
      Height          =   600
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   1058
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "cimg��ťͼ��"
      _Version        =   393216
      Begin MSComctlLib.ImageList cimg��ťͼ�� 
         Left            =   3120
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   393216
      End
   End
   Begin VB.Label Lab��¼���� 
      Height          =   255
      Left            =   9600
      TabIndex        =   16
      Top             =   6960
      Width           =   2295
   End
   Begin VB.Label Lab��¼�� 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   9120
      TabIndex        =   15
      Top             =   1440
      Width           =   2175
   End
End
Attribute VB_Name = "FrmQueryCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'���壺ְҵ����쵥λ��Ա��Ϣ��ѯ��ӡ����
'���ܣ���ְҵ����쵥λ��Ա��Ϣ�Ĳ�ѯ�ʹ�ӡ
'���ߣ���¶
'ʱ�䣺2012-04-28
'��ע������

Option Explicit
Private WithEvents mobjGUI As cls����ͨ�ö���
Attribute mobjGUI.VB_VarHelpID = -1
Public mblnInUse As Boolean
Dim lojb��ѯͳ�ƺ��� As Object    '��ѯͳ�ƺ���
Private indX, indY As Integer       '��¼�����vsflexgrid�����ꡣ
Private pstrPerson As String
Private mobjRec As Object
'�ý��湲�ö���
Private pobj��� As Object
Private pobjItem As Object
Private pobj����ģ�� As Object
Private pobj�����ҵ�� As Object
Private pobj���� As Object

Public Property Get pblnInUse() As Boolean
    On Error Resume Next
    pblnInUse = mblnInUse
End Property

'ͨ����λ��λ���ƽ���������ѯ
Private Sub ccmdQuery_Click()
'    If cgrdList.rows > 1 Then Exit Sub
    sub��ʼ��ѯ
End Sub

Private Sub cgrdList_DblClick()
    indX = cgrdList.MouseRow
    indY = cgrdList.MouseCol
    If indX < 0 Or indY < 0 Then
        Exit Sub
    ElseIf indX >= 0 And indX < cgrdList.rows And indY >= 0 And indY < cgrdList.cols Then
        pstrPerson = cgrdList.TextMatrix(indX, 0)
'        sub�г���λ��Ϣ
    End If
End Sub



Private Sub CmdFCompany_Click()
    Dim lobjRec As Object                       '��λ��λ���صĽ����¼��
    
    On Error GoTo errHandler
'    Set lobjRec = pobjҵ�����.func��λ��λ      '������λ��λ���档  ע����2015-11-2 ��ΪҪ�����µĽ��棨��λ��λ��ѯ���棩

    frmQueryCompanyLocation.Show 1, Me            '���õ�λ��λ��ѯ���棬д��2015-11-2  ���˹��̽����������һ�䣬ͬʱҲ����ע�͵������Ǿ䣬����δ����
    
    
    
    '��������δ�Ķ���������ԭ��������λ��λ����ʱ������   2015-11-2
    '��ȡ��λ�ĵ�λ����ʾ�ڡ���λ���ơ�¼����С�(��ʱֻ��ʾ����λ���ơ�)
    If Not lobjRec Is Nothing Then
        If lobjRec.RecordCount > 0 Then
            ctxtCompanyName.Text = IIf(IsNull(lobjRec("��λ����")), "", lobjRec("��λ����"))
        End If
    End If
    Exit Sub
errHandler:
    'If Err.Number = 0 Then Exit Sub
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ������", "Form1", "CmdFCompany_Click", 6666, lstrError, False
End Sub

Private Sub CmdFCompany_LostFocus()
    flag����.Value = 1
End Sub

Private Sub DTP����ʱ��_GotFocus()
    flag����.Value = 1
End Sub

Private Sub ctxtCompanyName_Change()
    
'    If Trim(ctxtCompanyName.Text) <> "" Then
'        ccmdQuery.Enabled = True
'    Else
'        ccmdQuery.Enabled = False
'    End If

End Sub

Private Sub ctxtCompanyName_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 13 And ctxtCompanyName.Text <> "" Then
        sub��ʼ��ѯ
    End If
    
End Sub

Private Sub DTP��ʼʱ��_GotFocus()
    flag����.Value = 1
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
    With lcol��������ť
        .Add "Ԥ������(&N)108"
        .Add "|"
        .Add "��ӡ����(&M)107"
        .Add "|"
        .Add "����(@X)113"
        .Add "|"
        .Add "�˳�"
    End With
    
    'Ϊ��Ҫͨ������ͨ�ö�����Ƶĸ��ֿؼ�����ʼֵ��
    With mobjGUI
        Set .Form = Me
        Set .c������ = ctlb������
    End With
    
    '��סctrl��ѡ���������
    cgrdList.HighLight = flexHighlightWithFocus
    cgrdList.SelectionMode = flexSelectionListBox
    
    '���ý���ͨ�ö����ṩ�ķ������Խ���ؼ����г�ʼ����
    mobjGUI.subInitialize lcol��������ť, ""
    '���ҽʦ��ʾ����ʾ��ǰ�û���
'    ctxtDoctor.Text = um�û���
'    ctxtDoctor.Enabled = False
    '��ʾ��ǰ����
    DTP��ʼʱ��.Value = DateAdd("m", -1, Date)
    DTP��ֹʱ��.Value = Date
    
    '��ѯ������ʼ��
    '������ѯͳ�ƺ�������
    Set lojb��ѯͳ�ƺ��� = CreateObject("ְҵ������.clsQueryStatis")

    '������ʼ��
    Set pobj��� = CreateObject("ְҵ������.clsMedicalExam")
    Set pobj����ģ�� = CreateObject("ְҵ������.clsMedicalExamTemplate")
    Set pobj�����ҵ�� = CreateObject("ְҵ�������¼��.clsCommon")
    Set pobj���� = pobjDict.Fetch("ְҵ���������ֵ�")
    
    '����ʼ�������ǣ�2012-11-01
    '��סctrl��ѡ���������
    cgrdList.HighLight = flexHighlightWithFocus
    cgrdList.SelectionMode = flexSelectionListBox
    cgrdList.cols = 0
    With cgrdList
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "ϵͳ���"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "����"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "�Ա�"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "����"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "Σ������"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "ְҵ����"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "��λ����"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "�������"
        .AutoSize 0, .cols - 1, 0, 0
        .SelectionMode = flexSelectionListBox
    End With
    
    
    '2012-05-21 ��¶
    '����Ȩ������
    Dim lobjTmp As Object
    Set lobjTmp = CreateObject("ְҵ������.clsPermissionConfigure")
    If lobjTmp.func���Ҳ���Ȩ��(um�û����, "ְҵ�����_��λͳ��_��ӡ") = False Then
        ctlb������.Buttons(3).Visible = False
        ctlb������.Buttons(4).Visible = False
    End If
    Set lobjTmp = Nothing
    '2012-05-21
       
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ������", "frmRegistermanage", "Form_Load", 6666, lstrError, False
    Exit Sub
    Resume
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    mblnInUse = False
    Set mobjGUI = Nothing
End Sub

'��ѯ����
Private Sub sub��ʼ��ѯ()
    Dim lobjRec As Object
    Dim sql As String
    Dim sql�����ʱ�� As String
    Dim sql��ѯ��� As String
    
    On Error GoTo errHandler
    
    '������������ֹʱ�䣬��λ���Ƶȣ����ҵ���Ϣ������ҪֻҪ�����ģ�������û�м���꣩����ֻҪ�����ɵ���Ա������������ֻҪ�����ģ�����������ֻҪ�����ɵģ� 2015-10-29
    sql = "select ϵͳ���,����,�Ա�,����,Σ������,�ֹ���,��λ����,������� from ְҵ�����_���������ݿ� where ��λ���� = '" & Trim(ctxtCompanyName.Text) & "' and ���״̬ >=1 and convert(varchar(10),�������,120) between '" & Format(DTP��ʼʱ��.Value, "yyyy-mm-dd") & "' and '" & Format(DTP��ֹʱ��.Value, "yyyy-mm-dd") & "'"
'    sql = "select ϵͳ���,����,�Ա�,����,Σ������,ְҵ����,��λ����,�������,����״̬,����ԭ��,������Ŀ from ְҵ�����_���������ݿ� where ��λ���� = '" & Trim(ctxtCompanyName.Text) & "' and ���״̬ in (6,7) and convert(varchar(10),�������,120) between '" & Format(DTP��ʼʱ��.Value, "yyyy-mm-dd") & "' and '" & Format(DTP��ֹʱ��.Value, "yyyy-mm-dd") & "'"
    Set lobjRec = dafuncGetData(sql)
    cgrdList.Clear
    Set cgrdList.DataSource = lobjRec
    Set mobjRec = lobjRec
    
'    sql = "select ϵͳ���,����,�Ա�,����,Σ������,ְҵΣ������,������,��Ϻʹ������,������ʷ,������� from ְҵ�����_��챨����ͼ where"
'
'    sql = sql & " ��λ���� = '" & Trim(ctxtCompanyName.Text) & "' "
'
'    If flag����.Value = 1 Then
'        sql = sql & " and (������� >= '" & DTP��ʼʱ��.Value & "' and ������� <= '" & DTP��ֹʱ��.Value & "')"
'    End If
'
'    Set lobjRec = lojb��ѯͳ�ƺ���.func���ز�ѯ��Ϣ(sql)
'    Set mobjRec = lobjRec
'    sql��ѯ��� = sql
'    '��ʾ֮ǰ��������������е���Ϣ
'    cgrdList.Clear
'
'    Set cgrdList.DataSource = lobjRec
'    cgrdList.Editable = flexEDNone
'    cgrdList.AutoSize 1, cgrdList.Cols - 1, 0, 0
'    sql = ""
'
'    sql�����ʱ�� = "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ְҵ�����_��λ�����ӡ��]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)drop table [dbo].[ְҵ�����_��λ�����ӡ��]"
'    dafuncGetData (sql�����ʱ��)
'    dafuncGetData (sql��ѯ���)
    
    '2012-05-23 ��¶��
    'vsflexgrid�п�Ȱ������Զ������������ͷ����ͷ����������
    cgrdList.AutoSize 0, cgrdList.cols - 1, 0, 0
    cgrdList.ExplorerBar = flexExSort
    cgrdList.DataMode = flexDMFree
   '2012-05-23��
   
   
    Lab��¼����.Caption = "���Ϲ��м�¼��" & lobjRec.RecordCount & "��"   '�ڱ�ǩ����ʾ��¼��Ŀ  2015-11-3
    
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����ý���", "frmSetMedicalExamTemplate", "Timer1_Timer", 66666, lstrError, False
    MousePointer = 0
    '�ָ�������Բ�����
    Me.Enabled = True
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    ctlb������.Width = Me.ScaleWidth - ctlb������.Left * 2
    Frame1.Width = Me.ScaleWidth - ctlb������.Left * 2
'    Frame1.Height = Me.ScaleHeight - Frame1.Top - 20
'    Frame1.Height = Me.ScaleHeight - 50
    cgrdList.Width = Frame1.Width - cgrdList.Left * 2
    cgrdList.Height = Frame1.Height - cgrdList.Top - 20

End Sub

'���湤������ť�����趨
Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = True
    
    Dim lcolID As New Collection
    Dim lobj������� As Object
    Dim i As Integer
    Set lobj������� = CreateObject("ְҵ������.clsMedicalExam")
    lobj�������.ϵͳ������� = pstrPerson
    
    Select Case Operate
        Case "Ԥ������"
            
            If cgrdList.rows <= 1 Then
                MsgBox "���ܴ�ӡ���ܼ챨�棡����"
                Set lobj������� = Nothing
                Exit Sub
            End If
            
            Dim lobjRec As Object, ltempRec As Object
            Dim lcolInfo As Collection, lcolFactor As Collection, lcolInfo2 As Collection, lcolItem As Collection, lcol As Collection, lcol2 As Collection
            Dim lstr As String, ltemp As String
            Dim lint As Integer

            lstr = "select Σ������,count(*) as ���� from ְҵ�����_���������ݿ� where 1=1 and Σ������ <> '' "
            If flag����.Value = 1 Then
                lstr = lstr & " and ��λ���� = '" & Trim(ctxtCompanyName.Text) & "'"
            End If
            If flag����.Value = 1 Then
                lstr = lstr & " and (convert(varchar(10),�������,120) >= '" & Format(DTP��ʼʱ��.Value, "yyyy-mm-dd") & "' and convert(varchar(10),�������,120) <= '" & Format(DTP��ֹʱ��.Value, "yyyy-mm-dd") & "')"
            End If
            lstr = lstr & " group by Σ������"
            Set lobjRec = dafuncGetData(lstr)
'            If Not (lobjRec.BOF Or lobjRec.EOF) Then
'                MsgBox ""
'            End If
            Set lcol = New Collection
            Set lcol2 = New Collection
            Set lcolInfo = New Collection
            Set lcolInfo2 = New Collection
            Set lcolItem = New Collection
            Set lcolFactor = New Collection
            lint = 0
            
            Dim temp1 As String, temp2 As String
            '��һ����
            '����Ӧ���������Ӧ�������ǵǼ��˵�������  2015-10-29��
            Dim SlobjRec As Object
            Dim Slstr As String
            Dim sum���� As Integer
            Dim mstr��ʼ���� As String
            Dim mstr��ֹ���� As String
             Slstr = "select count(*) as ���� from ְҵ�����_���������ݿ� where 1=1 and Σ������ <> '' "
            If flag����.Value = 1 Then
                Slstr = Slstr & " and ��λ���� = '" & Trim(ctxtCompanyName.Text) & "'"
            End If
            If flag����.Value = 1 Then

                '�����Ǽǲ�ѯ�Ŀ�ʼʱ����ǰһ����  2015-11-3
                 mstr��ʼ���� = Format(DateAdd("d", -30, DTP��ʼʱ��), "yyyy-mm-dd")

                Slstr = Slstr & " and (convert(varchar(10),��������,120) >= '" & mstr��ʼ���� & "' and convert(varchar(10),��������,120) <= '" & Format(DTP��ֹʱ��.Value, "yyyy-mm-dd") & "')"
'                Slstr = Slstr & " and (convert(varchar(10),��������,120) >= '" & Format(DTP��ʼʱ��.Value, "yyyy-mm-dd") & "' and convert(varchar(10),��������,120) <= '" & Format(DTP��ֹʱ��.Value, "yyyy-mm-dd") & "')"
            End If
            Set SlobjRec = dafuncGetData(Slstr)
            sum���� = SlobjRec("����")
            '2015-10-29��
            
            For i = 1 To lobjRec.RecordCount
                temp1 = lobjRec("Σ������")
                temp2 = lobjRec("����")
                lcolFactor.Add temp1, "Σ������" & i
                lcolFactor.Add temp2, "����" & i
'                lcolInfo.Add lobjRec("Σ������")
'                lcolInfo2.Add lobjRec("����")
                lint = lint + Val(lobjRec("����"))
                lobjRec.MoveNext
                temp1 = ""
                temp2 = ""
            Next
            lcolFactor.Add Trim(ctxtCompanyName.Text), "��λ����"
            lcolFactor.Add Format(DTP��ʼʱ��.Value, "yyyy��mm��dd��"), "�������"
            lcolFactor.Add Str(sum����), "Ӧ������"   '����Ӧ������  2015-10-29
            lcolFactor.Add Str(lint), "ʵ������"
            lcolFactor.Add DTP��ʼʱ��.Value, "��ʼ����"
            lcolFactor.Add DTP��ֹʱ��.Value, "��ֹ����"
'            lcolFactor.Add Format(DTP��ʼʱ��.Value, yyyy - mm - dd), "��ʼ����"
'            lcolFactor.Add Format(DTP��ֹʱ��.Value, yyyy - mm - dd), "��ֹ����"

'            Set lobjRec = Nothing
'            lstr = "select distinct b.���� as �������� from ְҵ�����_���ҽ��۱� a,ϵͳ����_�ֵ�_�ֵ����ݱ� b where a.���� = b.���" _
'            & " and b.ID = (select ID from ϵͳ����_�ֵ�_�ֵ���б� where ���� = 'ְҵ���������ֵ�') and b.���� <> '���ս���¼��' " _
'            & " and ϵͳ��� in (select ϵͳ��� from ְҵ�����_���������ݿ� where ��λ���� = '" & Trim(ctxtCompanyName.Text) & "'" _
'            & " and (convert(varchar(10),�������,120) >= '" & Format(DTP��ʼʱ��.Value, "yyyy-mm-dd") & "' and convert(varchar(10),�������,120) <= '" & Format(DTP��ֹʱ��.Value, "yyy-mm-dd") & "') and ���״̬ = 7)"
            
            Set ltempRec = dafuncGetData("select distinct b.���� as �������� from ְҵ�����_���ҽ��۱� a,ϵͳ����_�ֵ�_�ֵ����ݱ� b where a.���� = b.���" _
            & " and b.ID = (select ID from ϵͳ����_�ֵ�_�ֵ���б� where ���� = 'ְҵ���������ֵ�') and b.���� <> '���ս���¼��' " _
            & " and ϵͳ��� in (select ϵͳ��� from ְҵ�����_���������ݿ� where ��λ���� = '" & Trim(ctxtCompanyName.Text) & "'" _
            & " and (convert(varchar(10),�������,120) >= '" & Format(DTP��ʼʱ��.Value, "yyyy-mm-dd") & "' and convert(varchar(10),�������,120) <= '" & Format(DTP��ֹʱ��.Value, "yyyy-mm-dd") & "') and ���״̬ = 7)")
            
            While Not (ltempRec.EOF Or ltempRec.BOF)
                ltemp = ltempRec("��������")
                lcolItem.Add Left(ltemp, Len(ltemp) - 1)
                ltempRec.MoveNext
            Wend
'            Set lobjRec = Nothing
            lcolFactor.Add lcolItem, "�����Ŀ"
'            lcolFactor.Add lcolInfo, "Σ������"
'            lcolFactor.Add lcolInfo2, "Σ�����"
            
'            lstr = "select b.����,count(*) ���� from dbo.ְҵ�����_�������ͼ a,ְҵ�����_�����Ŀ���ñ� b where a.�����Ŀ=b.���� and ϵͳ��� in(" _
'                & " select ϵͳ��� from ְҵ�����_���������ݿ� where Σ������ = 'Xray'and ��λ���� = '" & Trim(ctxtCompanyName.Text) & "'" _
'                & " and (������� >= '" & DTP��ʼʱ��.Value & "' and ������� <= '" & DTP��ֹʱ��.Value & "') and ���״̬ = 7" _
'                & ") and ������� = '���ϸ�' group by b.����"
'
'            Set lobjRec = dafuncGetData(lstr)
'
'            While Not lobjRec.EOF
'                lcol.Add lobjRec("����")
'                lcol2.Add lobjRec("����")
'                lobjRec.nextmove
'            Wend

'            '�����Ա���ѯ  2015-10-30
'        lstr = "select * from ְҵ�����_����ģ�������Ŀ��(��������,�����Ŀ)values('" & mstr������ & "','" & lobjItem.���� & "') "
'        dafuncGetData (lstr)
 
            
            sub�༭�ܼ챨�� lcolFactor
'            Set lobjRec = Nothing
'            Set ltempRec = Nothing
'            Set lcolFactor = Nothing
            
            '�ڶ�����
            
'            If cgrdList.Rows >= 2 Then
'                Dim lobjRec As Object
'                Dim lcolCompany As Collection
'                Set lobjRec = dafuncGetData("select �������,��λ����,��ַ from ��λ����_��λ��λ��ѯ��ͼ where ��λ���� = '" & Trim(ctxtCompanyName.Text) & "'")
'                If Not lobjRec.EOF Then
'                    Set lcolCompany = New Collection
'                    lcolCompany.Add lobjRec("�������"), "�������"
'                    lcolCompany.Add lobjRec("��λ����"), "��λ����"
'                    lcolCompany.Add lobjRec("��ַ"), "��λ��ַ"
'                End If
'                sub�༭��λ���� mobjRec, lcolCompany
'                lcolID.Add cgrdList.TextMatrix(1, 0)    'ֱ�ӷ����һ��ϵͳ���
'                pobjҵ�����.Sub��ӡ��λ���� "ְҵ�����_��λ����", lcolID, False, True
'            End If
        Case "��ӡ����"
            If cgrdList.rows >= 2 Then
                lcolID.Add cgrdList.TextMatrix(1, 0)    'ֱ�ӷ����һ��ϵͳ���
                pobjҵ�����.Sub��ӡ��λ���� "ְҵ�����_��λ����", lcolID, True, False
            End If
        Case "����"
            If cgrdList.Row < 1 Then
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
                cgrdList.ColDataType(0) = flexDTString
                cgrdList.SaveGrid lstrFile, flexFileExcel, True   '����excelϵͳ���Ϊ����
                'cgrdMain.SaveGrid lstrFile, flexFileTabText, True
                '2012-04-14 �ڵ���
            End If
        Case "�˳�"
            Unload FrmQueryCompany
            Set FrmQueryCompany = Nothing
            Cancel = True
    End Select
    
    Set lobj������� = Nothing
    
    Exit Sub
errHandler:
    sfsub������ "ְҵ������", "FrmQueryCompany", "mobjGUI_BeforeOperate", Err.Number, Err.Description, False
    Cancel = True
    Exit Sub
    Resume
End Sub

Public Function BigNum(num As Integer) As String
    Select Case num
        Case 0
            BigNum = "��"
        Case 1
            BigNum = "һ"
        Case 2
            BigNum = "��"
        Case 3
            BigNum = "��"
        Case 4
            BigNum = "��"
        Case 5
            BigNum = "��"
        Case 6
            BigNum = "��"
        Case 7
            BigNum = "��"
        Case 8
            BigNum = "��"
        Case 9
            BigNum = "��"
        Case Else
            BigNum = "Err"
    End Select
End Function

Private Sub Timer1_Timer()
    Dim lojbRec As Object   '���ݿ�������
    Dim i As Integer
    On Error GoTo errHandler
    
    Timer1.Enabled = False
    '����ʱ������
    DTP��ʼʱ��.Value = DateAdd("M", -1, Now)
    DTP��ֹʱ��.Value = Now
    
errHandler:
    Dim lstrError  As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����ý���", "frmSetMedicalExamTemplate", "Timer1_Timer", 6666, lstrError, False
    MousePointer = 0
    '�ָ�������Բ�����
    Me.Enabled = True
End Sub
