VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Begin VB.Form frm���¸����޸�Ȩ�� 
   Caption         =   "���¸����޸�Ȩ��"
   ClientHeight    =   3660
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7290
   LinkTopic       =   "Form1"
   ScaleHeight     =   3660
   ScaleWidth      =   7290
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton ccmdExit 
      Cancel          =   -1  'True
      Caption         =   "�� ��(&X)"
      Height          =   495
      Left            =   4320
      TabIndex        =   9
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox ctxtHour 
      Enabled         =   0   'False
      Height          =   270
      Left            =   1800
      TabIndex        =   7
      Top             =   2280
      Width           =   1575
   End
   Begin VSFlex8Ctl.VSFlexGrid cgrdInfo 
      Height          =   2175
      Left            =   3480
      TabIndex        =   6
      Top             =   600
      Width           =   3735
      _cx             =   2088769980
      _cy             =   2088767228
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
   Begin VB.CommandButton ccmdPermission 
      Caption         =   "�����޸�Ȩ��"
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton ccmdQuery 
      Caption         =   "�� ѯ"
      Default         =   -1  'True
      Height          =   495
      Left            =   840
      TabIndex        =   4
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox ctxtPatientNo 
      Height          =   270
      Left            =   1440
      TabIndex        =   3
      Top             =   1800
      Width           =   1935
   End
   Begin VB.ComboBox combDoctorNo 
      Height          =   300
      Left            =   1320
      TabIndex        =   2
      Text            =   "Combo3"
      Top             =   1320
      Width           =   2055
   End
   Begin VB.ComboBox combDoctorName 
      Height          =   300
      Left            =   1320
      TabIndex        =   1
      Text            =   "Combo2"
      Top             =   840
      Width           =   2055
   End
   Begin VB.ComboBox combDept 
      Height          =   300
      Left            =   1320
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "��    ��"
      Height          =   180
      Left            =   240
      TabIndex        =   14
      Top             =   360
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "ҽʦ����"
      Height          =   180
      Left            =   240
      TabIndex        =   13
      Top             =   840
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "ҽʦ���"
      Height          =   180
      Left            =   240
      TabIndex        =   12
      Top             =   1320
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "���벡�˱��"
      Height          =   180
      Left            =   240
      TabIndex        =   11
      Top             =   1800
      Width           =   1080
   End
   Begin VB.Label Label2 
      Caption         =   "��ʾ����Ctrl����ѡ������������������ѡ��ʱĬ�϶��б��������������"
      Height          =   375
      Left            =   3480
      TabIndex        =   10
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "�ӳ�ʱ�䣨Сʱ��"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2280
      Width           =   1455
   End
End
Attribute VB_Name = "frm���¸����޸�Ȩ��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'2012-06-28 �ڵ��
'��������������ݡ�
'ҽʦ�޸������Ա�����ʱ�䳬��8Сʱ���������޸ģ�
'���ɹ���Ա��������棬�ҳ���ҽʦҪ�޸ĵ������Ա�������¸���Ȩ�ޡ�


'�޸��ˣ����� 2012.11.29  �޸ķ�Χ���˴���
'˵�������ݲ��˱�Ų�ѯ����ѯ�������벡�˱�š�
'bug�ţ�0000041
Option Explicit

Private pobjDept As Object
Private pobjDoctor As Object
Private pobjPatient As Object
Private PatientList() As String

'Private Sub cchkPatientNoSelect_Click()
'    If cchkPatientNoSelect.Value = 0 Then
'        cgrdInfo.Rows = 1
'    Else
'        Set cgrdInfo.DataSource = pobjPatient
'    End If
'End Sub

Private Sub ccmdExit_Click()
    Unload Me
    Set frm���¸����޸�Ȩ�� = Nothing
End Sub

Private Sub ccmdPermission_Click()
    Dim i As Integer
    Dim deptIndex As Integer
    Dim lobjTmp As Object
    
    deptIndex = combDept.ListIndex + 1
    
    Set lobjTmp = CreateObject("ְҵ������.clsPermissionConfigure")
    If cgrdInfo.SelectedRows = 0 Then
        For i = 1 To cgrdInfo.Rows - 1
            lobjTmp.func���˵����ҽ�������޸� cgrdInfo.TextMatrix(i, 0), deptIndex, "2"
        Next
    Else
        For i = 0 To cgrdInfo.SelectedRows - 1
            lobjTmp.func���˵����ҽ�������޸� cgrdInfo.TextMatrix(cgrdInfo.SelectedRow(i), 0), deptIndex, "2"
        Next
    End If
    subInitPatientNo    '���²�ѯ�����Ա
End Sub

Private Sub ccmdQuery_Click()
    Dim strSQL As String
    Dim deptState As String
    Dim DoctorNo As String
    Dim deptIndex As Integer
    Dim lobjRec As Object
    Dim addIndex As String
    Dim addNum As Integer
    Dim i As Integer
    
'    If cchkPatientNoSelect.Value = 1 Then Exit Sub
'    If cchkPatientNoWrite.Value = 1 Then Exit Sub
    
'    pobjPatient.Filter = ""
'    For i = 0 To pobjPatient.recordcount - 1: PatientList(4, i) = "0": Next
'
'    deptIndex = combDept.ListIndex + 1
'    addIndex = "1"
'    addNum = 0
'    For i = 0 To pobjPatient.recordcount - 1
'        deptState = PatientList(3, i)
'        If Right(Left(deptState, deptIndex), 1) = "3" Then
'            PatientList(4, i) = addIndex
'            addNum = addNum + 1
'        End If
'    Next
'
'    addIndex = "2"
'    addNum = 0
'    combDoctorName_Click
'    DoctorNo = combDoctorNo.Text
'    For i = 0 To pobjPatient.recordcount - 1
'        If PatientList(4, i) = "1" Then
'            Set lobjRec = dafuncGetData("select * from ְҵ�����_�����Ϣ_" & combDept.Text & " where ϵͳ���='" & PatientList(1, i) & "'")
'            If lobjRec.recordcount > 0 Then
'                If lobjRec("���ҽʦ") = DoctorNo Or lobjRec("���ҽʦ") = combDoctorName.Text Then
'                    PatientList(4, i) = addIndex
'                    addNum = addNum + 1
'                End If
'            End If
'        End If
'    Next
'
'    cgrdInfo.Rows = 1 + addNum
'    addNum = 0
'    For i = 0 To pobjPatient.recordcount - 1
'        If PatientList(4, i) = addIndex Then
'            addNum = addNum + 1
'            cgrdInfo.TextMatrix(addNum, 0) = PatientList(1, i)
'            cgrdInfo.TextMatrix(addNum, 1) = PatientList(2, i)
'        End If
'    Next
    subInitPatientNo
End Sub

Private Sub combDept_Click()
    If InStr(combDept.Text, "X��") = 1 Then
        ctxtHour.Text = 48
    Else
        ctxtHour.Text = 8
    End If
'    cchkDept.Value = 1
'    cchkDept.Enabled = False
End Sub

Private Sub combDoctorName_Click()
    'cchkDoctorName.Value = 1
    'cchkDoctorName.Enabled = False
    combDoctorNo.ListIndex = combDoctorName.ListIndex
End Sub

Private Sub combDoctorNo_Click()
    'cchkDoctorNo.Value = 1
    'cchkDoctorNo.Enabled = False
    combDoctorName.ListIndex = combDoctorNo.ListIndex
End Sub

'Private Sub combPatientNo_Click()
'    cchkPatientNoSelect.Value = 1
'    pobjPatient.Filter = "ϵͳ���='" & PatientList(1, combPatientNo.ListIndex) & "'"
'    With cgrdInfo
'        Set .DataSource = pobjPatient
'        .ColHidden(.Cols - 1) = True
'        .AutoSize 0, .Cols - 1, 0, 0
'        .ExplorerBar = flexExSort
'        .DataMode = flexDMFree
'    End With
'End Sub

'Private Sub ctxtPatientNo_LostFocus()
'    Dim i As Integer
'    pobjPatient.Filter = ""
'    For i = 0 To pobjPatient.recordcount - 1
'        If ctxtPatientNo.Text = PatientList(1, i) Then
''            combPatientNo.ListIndex = i
'            Exit For
'        End If
'    Next
'End Sub

Private Sub Form_Load()
    subInitDept
    subInitDoctor
'    subInitPatientNo
    cgrdInfo.SelectionMode = flexSelectionListBox
End Sub

Sub subInitDept()
    Set pobjDept = pobjDict.Fetch("ְҵ���������ֵ�")
    
    pobjDept.movefirst
    combDept.Clear
    While pobjDept.EOF = False
        If Right(pobjDept("����"), 1) = "��" Then
            combDept.AddItem pobjDept("����")
            combDept.ItemData(combDept.NewIndex) = pobjDept("���")
        End If
        pobjDept.movenext
    Wend
    combDept.ListIndex = 0
End Sub

Sub subInitDoctor()
    Dim lobjRec As Object
    Set lobjRec = CreateObject("ְҵ������.clsPermissionConfigure")
    Set pobjDoctor = lobjRec.func��ȡְҵ��������ҽʦ������Ϣ
    Set lobjRec = Nothing
    If pobjDoctor.EOF Or pobjDoctor.bof Then Exit Sub
    pobjDoctor.movefirst
    combDoctorName.Clear
    combDoctorNo.Clear
    While pobjDoctor.EOF = False
        combDoctorName.AddItem pobjDoctor("����")
        If Not pobjDoctor("���") = "gues" Then
            combDoctorName.ItemData(combDoctorName.NewIndex) = pobjDoctor("���")
        
        combDoctorNo.AddItem pobjDoctor("���")
        combDoctorNo.ItemData(combDoctorNo.NewIndex) = pobjDoctor("���")
        End If
        pobjDoctor.movenext
    Wend
    combDoctorName.ListIndex = 0
    combDoctorNo.ListIndex = 0
    
End Sub

Sub subInitPatientNo()
    Dim lobjRec As Object
    Dim i As Integer
    Set lobjRec = CreateObject("ְҵ������.clsPermissionConfigure")
    '�޸��ˣ����� 2012.12.20  ����
    '�޸�˵���������˱��Ϊ��ʱ��ִ�в�ѯ��
    If ctxtPatientNo.Text = "" Then Exit Sub
    '�޸��ˣ����� 2012.12.20  ����
    Set pobjPatient = lobjRec.func��ȡְҵ�������Ա������Ϣ(Trim(ctxtPatientNo.Text))
    Set lobjRec = Nothing
    
'    If pobjPatient.recordcount = 0 Then Exit Sub
'    ReDim PatientList(1 To 4, 0 To pobjPatient.recordcount - 1) 'As String
'    pobjPatient.movefirst
'    combPatientNo.Clear
    If pobjPatient Is Nothing Then
        cgrdInfo.Row = 1
    Else
        Set cgrdInfo.DataSource = pobjPatient
    End If
'    For i = 0 To pobjPatient.recordcount - 1
'        PatientList(1, i) = pobjPatient("ϵͳ���")
'        PatientList(2, i) = pobjPatient("����")
'        PatientList(3, i) = pobjPatient("�������״̬")
'        PatientList(4, i) = "1" '���ڱ����Ӧ�����Ա�Ƿ�ѡ�в���ʾ��cgrdinfo��
''        combPatientNo.AddItem pobjPatient("ϵͳ���")
''        combPatientNo.ItemData(combPatientNo.NewIndex) = i
'        pobjPatient.movenext
'    Next
'    combPatientNo.ListIndex = 0
    cgrdInfo.AutoSize 0, cgrdInfo.Cols - 1, 0, 0
End Sub
