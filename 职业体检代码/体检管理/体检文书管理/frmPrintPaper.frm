VERSION 5.00
Object = "{8099FCC2-0A81-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPrintPaper 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "��ӡ�������"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   11070
   ClipControls    =   0   'False
   Icon            =   "frmPrintPaper.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   11070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CheckBox cchkPreview 
      Caption         =   "��ӡǰԤ��"
      Height          =   285
      Left            =   8760
      TabIndex        =   20
      Top             =   840
      Width           =   1395
   End
   Begin MSComctlLib.ImageList cimg��ťͼ�� 
      Left            =   5400
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.CheckBox cchkPrintAll 
      Caption         =   "ȫ����ӡ"
      Height          =   300
      Left            =   7200
      TabIndex        =   19
      Top             =   840
      Width           =   1095
   End
   Begin VB.ComboBox ccmbPaper 
      Height          =   300
      ItemData        =   "frmPrintPaper.frx":0442
      Left            =   1230
      List            =   "frmPrintPaper.frx":0452
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   885
      Width           =   3615
   End
   Begin VB.Frame cframSearch 
      Appearance      =   0  'Flat
      Caption         =   "��ѯ��죺"
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   0
      TabIndex        =   13
      Top             =   1200
      Width           =   10545
      Begin VB.CheckBox cchk����� 
         Caption         =   "����"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   11
         Top             =   1080
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox cchk����� 
         Caption         =   "����"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   10
         Top             =   1080
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.TextBox ctxt���� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2880
         TabIndex        =   6
         Top             =   720
         Width           =   1260
      End
      Begin VB.ComboBox ccmbUnit 
         Height          =   300
         Left            =   6600
         TabIndex        =   3
         Top             =   360
         Width           =   2655
      End
      Begin VB.CommandButton ccmdLocateUnit 
         Caption         =   "..."
         Height          =   375
         Left            =   9360
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   690
      End
      Begin MSComCtl2.DTPicker cdtpDate 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "yyyy-MM-dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   3
         EndProperty
         Height          =   300
         Left            =   3360
         TabIndex        =   2
         Top             =   375
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   129236992
         CurrentDate     =   36951
         MaxDate         =   73050
         MinDate         =   17899
      End
      Begin VB.TextBox ctxtEndNo 
         Enabled         =   0   'False
         Height          =   300
         Left            =   7440
         TabIndex        =   8
         Top             =   720
         Width           =   1740
      End
      Begin VB.TextBox ctxtStartNo 
         Enabled         =   0   'False
         Height          =   300
         Left            =   5160
         TabIndex        =   7
         Top             =   690
         Width           =   1860
      End
      Begin VB.OptionButton coptChoise 
         Caption         =   "��ϵͳ��ź�������ѯ"
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   705
         Width           =   2175
      End
      Begin VB.OptionButton coptChoise 
         Caption         =   "����λ�����ڲ�ѯ"
         Height          =   210
         Index           =   0
         Left            =   150
         TabIndex        =   1
         Top             =   435
         Value           =   -1  'True
         Width           =   2025
      End
      Begin VB.CommandButton ccmdSearch 
         Caption         =   "��ѯ(F2)"
         Height          =   375
         Left            =   9360
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   680
         Width           =   1050
      End
      Begin VB.Label clblInfo 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   6240
         TabIndex        =   23
         Top             =   1200
         Width           =   120
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   2400
         TabIndex        =   22
         Top             =   720
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   7080
         TabIndex        =   17
         Top             =   750
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ϵͳ��ţ�"
         Height          =   180
         Left            =   4155
         TabIndex        =   16
         Top             =   750
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������ڣ�"
         Height          =   180
         Index           =   0
         Left            =   2355
         TabIndex        =   15
         Top             =   435
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��쵥λ��"
         Height          =   180
         Left            =   5670
         TabIndex        =   14
         Top             =   435
         Width           =   900
      End
   End
   Begin MSComctlLib.Toolbar ctlbTool 
      Align           =   1  'Align Top
      Height          =   615
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   11070
      _ExtentX        =   19526
      _ExtentY        =   1085
      ButtonWidth     =   820
      ButtonHeight    =   926
      Appearance      =   1
      ImageList       =   "cimg��ťͼ��"
      _Version        =   393216
   End
   Begin VSFlex6Ctl.vsFlexGrid cgrdMain 
      Height          =   5055
      Left            =   0
      TabIndex        =   21
      Top             =   2760
      Width           =   10815
      _cx             =   69094020
      _cy             =   69083860
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
   Begin MSComDlg.CommonDialog ccmdFile 
      Left            =   0
      Top             =   500
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ѡ���ʽ��"
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   18
      Top             =   960
      Width           =   900
   End
End
Attribute VB_Name = "frmPrintPaper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'���ߣ��˺�

Private WithEvents mobjGUI As cls����ͨ�ö��� '����ͨ�ö���
Attribute mobjGUI.VB_VarHelpID = -1
Private mobj��켯 As Object
Private mstrϵͳ��Ź̶����� As String

Private mblnInUse As Boolean
Private mblnSys As Boolean
Public Property Get pblnInUse() As Boolean
    pblnInUse = mblnInUse
End Property

Private Sub cchk�����_Click(Index As Integer)
    If mblnSys Then Exit Sub
    ccmdSearch_Click
End Sub

Private Sub ctxtEndNo_GotFocus()
    On Error Resume Next
    If ctxtEndNo = "" Then
'        ctxtEndNo = mstrϵͳ��Ź̶�����
'        ctxtEndNo.SelStart = Len(ctxtEndNo)
'        ctxtEndNo.SelLength = 0
    End If

End Sub

Private Sub ctxtStartNo_GotFocus()
    On Error Resume Next
    If ctxtStartNo = "" Then
'        ctxtStartNo = mstrϵͳ��Ź̶�����
'        ctxtStartNo.SelStart = Len(ctxtStartNo)
'        ctxtStartNo.SelLength = 0
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
    Case vbKeyF2
        ccmdSearch_Click
 
    End Select
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 39 Then
        KeyAscii = 0
    End If

End Sub

Private Sub Form_Load()
    Dim lcolInfo As New Collection
    Dim i As Integer
    
    On Error GoTo errHandler
    '��������Ѿ���ʼ�������ٽ��г�ʼ����
    If mblnInUse Then Exit Sub
    
    MousePointer = 11
    
    '���ô�������ʹ�õı�־��
    mblnInUse = True
'    csbMain.Panels(1) = "�������ڳ�ʼ�������Ժ�..."
    
    Set mobj��켯 = CreateObject("������.clsMedicalExamSet")
   
    '��ʼ������ͨ�ö���ÿ�������Ӧһ������ͨ�ö���ʵ�������ɻ��ã��м��мǣ���
    Set mobjGUI = New cls����ͨ�ö���
    
    '���ù�����������Ҫ�ĸ��ְ�ť��
    Set lcolInfo = New Collection
    With lcolInfo
        .Add "Ԥ��"
        .Add "��ӡ"
        .Add "|"
        .Add "����(&O)111"
        .Add "|"
        .Add "�˳�"
    End With
    
    'Ϊ��Ҫͨ������ͨ�ö�����Ƶĸ��ֿؼ�����ʼֵ��
    With mobjGUI
        Set .Form = Me
        Set .c������ = ctlbTool
        'Set .c״̬�� = csbMain
        '���ý���ͨ�ö����ṩ�ķ������Խ���ؼ����г�ʼ����
        .subInitialize lcolInfo, ""
    End With
    ctlbTool.Buttons(1).Enabled = False
    ctlbTool.Buttons(2).Enabled = False
    
    '���뵥λ����
    Set lcolInfo = pobjҵ�����.���չ������䲾.��λ���Ƽ�
    For i = 1 To lcolInfo.Count
        ccmbUnit.AddItem lcolInfo(i)
    Next i
    
    Dim lobj��� As Object
    '���������󣬻�ȡϵͳ���ǰ��̶����֡�
    Set lobj��� = CreateObject("������.clsMedicalExam")
    mstrϵͳ��Ź̶����� = lobj���.ϵͳ��Ź̶�����
    Set lobj��� = Nothing
    
    '���
    cgrdMain.Rows = 1
    cdtpDate.Value = Format(Now, "yyyy-mm-dd")
    
    If ccmbPaper.ListCount > 0 Then
        ccmbPaper.ListIndex = 0
    End If
    cgrdMain.Editable = False
    
    MousePointer = 0
'    csbMain.Panels(1) = "����ѡ�������ʽ��"
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "���������", "frmPrintPaper", "Form_Load", 6666, lstrError, False
    MousePointer = 0
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    cgrdMain.Width = Me.ScaleWidth - cgrdMain.Left - 60
    cgrdMain.Height = Me.ScaleHeight - cgrdMain.Top - 60
    cframSearch.Width = Me.ScaleWidth - cframSearch.Left - 60
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set mobj��켯 = Nothing
    mblnInUse = False
End Sub

Private Sub ccmbPaper_Click()
    On Error Resume Next
    cframSearch.Enabled = True
    cgrdMain.Editable = True
    cgrdMain.Rows = 1
    ctlbTool.Buttons(1).Enabled = False
    ctlbTool.Buttons(2).Enabled = False
    clblInfo.Caption = ""
    mblnSys = True
    If ccmbPaper.Text = "���ǼǱ�" Or ccmbPaper.Text = "��쵥" Then
        cchk�����(0).Value = 1
        cchk�����(1).Value = 1
    End If
    mblnSys = False
'    csbMain.Panels(1) = "�������ѯ������Ȼ�󰴡���ѯ����ť��"
End Sub

Private Sub ccmbUnit_GotFocus()
    On Error Resume Next
    gfsubShowComboList ccmbUnit
    
End Sub

Private Sub ccmbUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ccmdSearch.SetFocus
    End If
End Sub

Private Sub ccmbUnit_LostFocus()
    On Error GoTo errHandler
    Dim i As Integer
    If Trim(ccmbUnit.Text) = "" Then Exit Sub
    
    '�ж�¼��ĵ�λ�Ƿ����б��д��ڣ�������������б�
    i = gffuncItemIsInComboBox(ccmbUnit, ccmbUnit.Text)
    If i = -1 Then
        '����ccmbUnit
        ccmbUnit.AddItem ccmbUnit.Text
    End If

    Exit Sub
errHandler:
    
End Sub

Private Sub ccmdLocateUnit_Click()
    Dim lobjRec As Object  '��λ��λ���صĽ����¼��
    
    On Error GoTo errHandler
    
    '������λ��λ���档
    Set lobjRec = pobjҵ�����.func��λ��λ
    '��ȡ��λ�ĵ�λ����ʾ�ڡ���λ���ơ�¼����С�
    If Not lobjRec Is Nothing Then
        If lobjRec.RecordCount > 0 Then
            ccmbUnit.Text = IIf(IsNull(lobjRec("��λ����")), "", lobjRec("��λ����"))
            
            '�Ѷ�λ�ĵ�λ���빤�����䲾��
            ccmbUnit_LostFocus
        End If
    End If
    
    Exit Sub
errHandler:
    
End Sub

Private Sub ccmdSearch_Click()
    Dim lobj��켯  As Object
    Dim lstrϵͳ��� As String
    Dim lstrError As String
    
    Dim i As Integer
    On Error GoTo errHandler
    
'    csbMain.Panels(1) = "���ڲ�ѯ���ݿ⣬���Ժ�..."
    MousePointer = 11
    cgrdMain.Rows = 1
    ctlbTool.Buttons(1).Enabled = False
    ctlbTool.Buttons(2).Enabled = False
    
    '�����켯�����ԡ�
    mobj��켯.subClear
    
    If coptChoise(0).Value Then
        '��������ں͵�λ
        mobj��켯.��������� = Format(cdtpDate.Value, "yyyy-mm-dd")
        mobj��켯.��������� = Format(cdtpDate.Value, "yyyy-mm-dd")
        mobj��켯.��λ���� = ccmbUnit.Text
    Else
        '����ʼϵͳ��źͽ���ϵͳ���
        mobj��켯.��ϵͳ��� = ctxtStartNo.Text
        mobj��켯.��ϵͳ��� = ctxtEndNo.Text
        mobj��켯.���� = ctxt����.Text
    End If
    
    If ccmbPaper.Text = "���ǼǱ�" Or ccmbPaper.Text = "��쵥" Then
        '���ǲ�ѯ���ǼǱ�ֻ����δ���������۵ġ�
        mobj��켯.���״̬ = P_LOGIN_STATUS & "," & P_EXAMING_STATUS & "," & P_CONCLUED_STATUS
    ElseIf ccmbPaper.Text = "�������" Then
        '���ǲ�ѯ���������ֻ����һ���������۵ġ�
        mobj��켯.���״̬ = P_ENDED_STATUS
    End If
    
        
    Set lobj��켯 = mobj��켯.Ԫ�ؼ�old("ϵͳ���,�Թܱ��,����,�Ա�,����,��λ����,��쵥��,��������,�������=convert(varchar(10),�������,20),������=isnull(������,'')")

    
    If cchk�����(0).Value = 1 And cchk�����(1).Value = 0 Then
        lobj��켯.Filter = "������='����' or ������=''"
    ElseIf cchk�����(0).Value = 0 And cchk�����(1).Value = 1 Then
        lobj��켯.Filter = "������<>'����' and ������<>''"
    
    End If
    
    If lobj��켯.RecordCount = 0 Then
        'û�ҵ���Ӧ�����Ա
        lstrError = "δ���ҵ����Դ�ӡ��������������Ա���������������������"
        If ccmbPaper.Text = "���ǼǱ�" Or ccmbPaper.Text = "��쵥" Then
            lstrError = "δ���ҵ����Դ�ӡ��������������Ա����δ�������۵ģ����������������������"
        Else
            lstrError = "δ���ҵ����Դ�ӡ��������������Ա�����������۵ģ����������������������"
        End If
        Err.Raise 6666, , lstrError
    Else
        '���뵽�����
        cgrdMain.FormatString = ""
        Set cgrdMain.DataSource = lobj��켯
'        gfsubLoadGridFromRec cgrdMain, lobj��켯, False, "ϵͳ���,����,�Ա�,����,��λ����,��������,�������,������,��Ҫ����,�Ѿ�����"
'        If cgrdMain.Rows > 1 Then
'            cgrdMain.Rows = cgrdMain.Rows - 1
'        End If
        ctlbTool.Buttons(1).Enabled = True
        ctlbTool.Buttons(2).Enabled = True
    End If
    clblInfo.Caption = "��ѯ�����" & cgrdMain.Rows - 1 & "�˴Ρ�"
errHandler:
    If Err <> 0 Then
        lstrError = func������(Err.Number, Err.Description)
        sfsub������ "���������", "frmPrintPaper", "ccmdSearch_Click", 6666, lstrError, False
    End If
    Set lobj��켯 = Nothing
'    csbMain.Panels(1) = ""
    MousePointer = 0
End Sub

Private Sub cdtpDate_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ccmbUnit.SetFocus
    End If
End Sub

Private Sub coptChoise_Click(Index As Integer)
    On Error Resume Next
    cgrdMain.Rows = 1
    ctlbTool.Buttons(1).Enabled = False
    ctlbTool.Buttons(2).Enabled = False
    If coptChoise(0).Value Then
        cdtpDate.Enabled = True
        ccmdLocateUnit.Enabled = True
        ccmbUnit.Enabled = True
        ctxtStartNo.Enabled = False
        ctxtEndNo.Enabled = False
        ctxt����.Enabled = False
        cdtpDate.SetFocus
    Else
        cdtpDate.Enabled = False
        ccmdLocateUnit.Enabled = False
        ccmbUnit.Enabled = False
        ctxtStartNo.Enabled = True
        ctxtEndNo.Enabled = True
        ctxt����.Enabled = True
        mblnSys = True
        cchk�����(0).Value = 1
        cchk�����(1).Value = 1
        mblnSys = False
        
        ctxt����.SetFocus
    End If
End Sub

Private Sub ctxtEndNo_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ccmdSearch.SetFocus
    End If
End Sub

Private Sub ctxtStartNo_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ctxtEndNo.SetFocus
    End If
End Sub


'���ܣ����������ϰ�ť��
Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    Dim lcol���  As Collection
    Dim i As Integer
    On Error GoTo errHandler
    
    Select Case Operate
        Case "Ԥ��"
            Set lcol��� = New Collection
            If cgrdMain.Row < 1 Then
                Err.Raise 6666, , "��ӡԤ��������ѡ��Ҫ��ӡ������¼��"
            Else
                lcol���.Add cgrdMain.TextMatrix(cgrdMain.Row, 0)
            End If
            '��ӡԤ����
            pobjҵ�����.Sub��ӡ���� ccmbPaper.Text, lcol���, False, True
            Cancel = True
            
        Case "��ӡ"
'            csbMain.Panels(1) = "��ӡ" & Trim(ccmbPaper.Text) & "�У����Ժ�..."
            'ȫ����ӡʱ�������еı��
            Set lcol��� = New Collection
            If cchkPrintAll.Value = 1 Then
                For i = 1 To cgrdMain.Rows - 1
                    lcol���.Add cgrdMain.TextMatrix(i, 0)
                Next i
            Else
                If cgrdMain.Row < 1 Then
                    Err.Raise 6666, , "��û��ѡ��ȫ����ӡ��������ѡ��Ҫ��ӡ������¼��"
                Else
                    lcol���.Add cgrdMain.TextMatrix(cgrdMain.Row, 0)
                End If
            End If
            '��ӡ
            If cchkPreview.Value = 1 Then
                If lcol���.Count > 1 Then
                    If Not sffuncMsg("��ӡ������飬���ܽ���Ԥ�������Ƿ������", sfѯ��) Then
                        Err.Raise 6666, , "��ȡ���˴�ӡ������"
                    End If
                End If
                pobjҵ�����.Sub��ӡ���� ccmbPaper.Text, lcol���, True
            Else
                pobjҵ�����.Sub��ӡ���� ccmbPaper.Text, lcol���, False
            End If
            '��ӡ��������������
            If cchkPrintAll.Value = 1 Then
                cgrdMain.Rows = 1
            Else
                cgrdMain.RemoveItem cgrdMain.Row
            End If
'            csbMain.Panels(1) = Trim(ccmbPaper.Text) & "��ӡ��ϣ����������"
            Cancel = True
            
        Case "����"
            Dim lstrFile As String
            ccmdFile.Filter = "Excel�ļ� (*.xls)|*.xls|�ı��ļ� (*.txt)|*.txt"
            ccmdFile.ShowSave
            lstrFile = ccmdFile.FileName
            If lstrFile <> "" Then
                cgrdMain.SaveGrid lstrFile, flexFileTabText, True
            End If
            
    End Select
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "���������", "frmPrintPaper", "mobjGUI_BeforeOperate", 6666, lstrError, False
    
'    csbMain.Panels(1) = ""
    Exit Sub
    Resume
End Sub
