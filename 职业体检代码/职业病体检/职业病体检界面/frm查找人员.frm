VERSION 5.00
Object = "{8099FCC2-0A81-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6.ocx"
Begin VB.Form frm������Ա 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "����"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   5820
   ClipControls    =   0   'False
   Icon            =   "frm������Ա.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   5820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VSFlex6Ctl.vsFlexGrid cgrdPerson 
      Height          =   2295
      Left            =   600
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   4935
      _cx             =   59384321
      _cy             =   59379664
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
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   4
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
   Begin VB.CommandButton ccmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2400
      Width           =   1155
   End
   Begin VB.CommandButton ccmdLocateUnit 
      Caption         =   "��λ(&L)"
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1800
      Width           =   945
   End
   Begin VB.OptionButton coptChoise 
      Caption         =   "����"
      Height          =   300
      Index           =   0
      Left            =   240
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1080
      Value           =   -1  'True
      Width           =   720
   End
   Begin VB.TextBox ctxtName 
      Height          =   300
      Left            =   1560
      TabIndex        =   0
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton ccmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2400
      Width           =   1155
   End
   Begin VB.OptionButton coptChoise 
      Caption         =   "����֤��"
      Height          =   240
      Index           =   1
      Left            =   240
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox ctxtHealthNo 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1560
      TabIndex        =   5
      Top             =   120
      Width           =   2745
   End
   Begin VB.ComboBox ccmbSex 
      Height          =   300
      ItemData        =   "frm������Ա.frx":000C
      Left            =   1560
      List            =   "frm������Ա.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1440
      Width           =   1725
   End
   Begin VB.ComboBox ccmbQueryUnit 
      Height          =   300
      Left            =   1560
      TabIndex        =   2
      Top             =   1800
      Width           =   2775
   End
   Begin VB.TextBox ctxtId 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1560
      TabIndex        =   7
      Top             =   600
      Width           =   2745
   End
   Begin VB.OptionButton coptChoise 
      Caption         =   "���֤��"
      Height          =   240
      Index           =   2
      Left            =   240
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label clblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����б���ѡ�������Ա��˫����꣨�򰴡�ȷ������ť�����أ�"
      ForeColor       =   &H00800000&
      Height          =   180
      Left            =   600
      TabIndex        =   14
      Top             =   2880
      Visible         =   0   'False
      Width           =   5220
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      Height          =   180
      Index           =   1
      Left            =   1080
      TabIndex        =   13
      Top             =   1080
      Width           =   360
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�Ա�"
      Height          =   180
      Index           =   0
      Left            =   1080
      TabIndex        =   12
      Top             =   1440
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��λ"
      Height          =   180
      Index           =   7
      Left            =   1080
      TabIndex        =   11
      Top             =   1800
      Width           =   360
   End
End
Attribute VB_Name = "frm������Ա"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public pstr������� As String '����/���/���顣
Public pstrϵͳ��� As String '���ҳ�����ϵͳ��š�

Private Sub ccmdCancel_Click()
    pstrϵͳ��� = ""
    Unload Me
End Sub

Private Sub ccmdOk_Click()
    Dim lobj��켯 As Object
    Dim lobj�����Ա As Object  '������ȡ�����Ա���������¼��
    Dim lobj��� As Object      '�����Ա�������졣
    Dim lobjRec As Object       'Recordset����켯���󷵻ص�Ԫ�ؼ���
    Dim lstrϵͳ��� As String  '����֤�Ŷ�Ӧ�����ϵͳ��š�
    Dim lstrError As String
    Dim i As Integer
    Dim j As Long
    
    On Error GoTo errHandler
    If cgrdPerson.Visible And cgrdPerson.Row > 0 Then
        '���б���ѡ����Ա����ȷ�����ء�
        cgrdPerson_DblClick
        
    Else
        '������켯����
        Set lobj��켯 = CreateObject("ְҵ������.clsMedicalExamSet")
        lobj��켯.subClear
        If coptChoise(0).Value Or coptChoise(2).Value Then
            '���������֤��
            If coptChoise(0).Value Then
                If Trim(ctxtName.Text) = "" And ccmbSex.Text = "" And Trim(ccmbQueryUnit.Text) = "" Then
                    Err.Raise 6666, , "���������������Ա𡢵�λ���ƣ�"
                End If
                
                '������켯�����������λ�������ԡ�
                With lobj��켯
                    .���� = Trim(ctxtName.Text)
                    .�Ա� = ccmbSex.Text
                    .��λ���� = Trim(ccmbQueryUnit.Text)
                End With
            Else
                '�����֤�Ų�ѯ��
                If Trim(ctxtId.Text) = "" Then
                    Err.Raise 6666, , "������������֤�ţ�"
                End If
                
                '������켯�����������λ�������ԡ�
                With lobj��켯
                    .���֤�� = Trim(ctxtId.Text)
                End With
            
            End If
        ElseIf coptChoise(1).Value Then
            '����֤�š�ϵͳ��š�
            If Trim(ctxtHealthNo.Text) = "" Then
                Err.Raise 6666, , "���������" & coptChoise(1).Caption & "��"
            End If
            If coptChoise(1).Caption = "ϵͳ���" Then
                lobj��켯.��ϵͳ��� = Trim(ctxtHealthNo.Text)
                lobj��켯.��ϵͳ��� = Trim(ctxtHealthNo.Text)
            Else
                '������֤��ת����ϵͳ��š�
                lstrϵͳ��� = pobjҵ�����.Func���ݽ���֤����Ż�ȡ���ϵͳ���(Trim(ctxtHealthNo.Text))
                If lstrϵͳ��� = "" Then
                    Err.Raise 6666, , "������Ľ���֤��û�ж�Ӧ������¼��"
                End If
                lobj��켯.��ϵͳ��� = lstrϵͳ���
                lobj��켯.��ϵͳ��� = lstrϵͳ���
            End If
        End If
        
        If pstr������� = "����" Then
            '��Ҫ���飬���һ�û�н��и���Ǽǡ�
            lobj��켯.�����־ = 1
            lobj��켯.����ϵͳ��� = ""
        End If
        
        '��ȡ���㶨λ���������ϵͳ��š�
        '�޸ģ�2001-12-30��������������������򣩡�
        Set lobjRec = lobj��켯.Ԫ�ؼ�("ϵͳ���,����,��λ����,�Ա�,����=datediff(year,��������,getdate()),�������" & IIf(pstr������� = "����", ",����������", ""), "����,��λ����,������� desc")
        If lobjRec.RecordCount = 0 Then
            'û�ҵ���Ӧ�����Ա
            lstrError = "δ���ҵ����������������������Ա��"
            If pstr������� = "����" Then
                lstrError = lstrError & "�����ǣ�" & Chr(13) & Chr(10) & "(1) ����첻��Ҫ���飬�����ҽʦ����������ʱû������Ҫ���飬�Լ����������" & Chr(13) & Chr(10) & "(2) �������Ѹ���Ǽǹ��ˡ�" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "����취��" & Chr(13) & Chr(10) & "(1) �������������������" & Chr(13) & Chr(10) & "(2) �������������ۣ�����ΪҪ���顣"
            End If
            Err.Raise 6666, , lstrError
        Else
            If lobjRec.RecordCount > 1 Then
                lstrϵͳ��� = ""
                If lobjRec.RecordCount > 100 Then
                    '��ѯ���̫�࣬��ʾ�û���
                    If Not sffuncMsg("������Ĳ�ѯ������Χ̫�󣬲�ѯ�������100����¼���������Ҿ��������Աû�ж�����������С��Χ��" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "����Ҫ����ô���¼��ѡ����", sfѯ��) Then
                        Exit Sub
                    End If
                End If
                '���ҵ�������¼ʱ����list�������뵽Grid�С�
                cgrdPerson.Rows = lobjRec.RecordCount + 1
                cgrdPerson.Cols = lobjRec.Fields.Count
                For j = 0 To cgrdPerson.Cols - 1
                    cgrdPerson.TextMatrix(0, j) = lobjRec.Fields(j).Name
                Next
                i = 1
                Do While Not lobjRec.EOF
                    For j = 0 To cgrdPerson.Cols - 1
                        cgrdPerson.TextMatrix(i, j) = IIf(IsNull(lobjRec(j).Value), "", lobjRec(j).Value)
                    Next
                    lobjRec.MoveNext
                    i = i + 1
                Loop
                cgrdPerson.AutoSize 0, cgrdPerson.Cols - 1
                cgrdPerson.Visible = True
                clblInfo.Visible = True
                cgrdPerson.SetFocus
            Else
                lstrϵͳ��� = lobjRec(0)
            End If
            
        End If
        
        Set lobj�����Ա = Nothing
        Set lobj��� = Nothing
        
        '���ҳɹ������ء�
        If lstrϵͳ��� <> "" Then
            pstrϵͳ��� = lstrϵͳ���
            Unload Me
        Else
            '��Ҫ���б���ѡ�񣬻����������ѯ������
        End If
    End If
    
    Exit Sub
    
errHandler:
    
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����沿��", "frm������Ա", "ccmdOk_Click", 6666, lstrError, False
    
    If coptChoise(0).Value Then
        ctxtName.SetFocus
    ElseIf coptChoise(1).Value Then
        ctxtHealthNo.SetFocus
    Else
        ctxtId.SetFocus
    End If
    
    Exit Sub
    Resume
End Sub

Private Sub cgrdPerson_DblClick()
    On Error Resume Next
    If cgrdPerson.Row < 1 Then Exit Sub
    
    '���ء�
    pstrϵͳ��� = cgrdPerson.TextMatrix(cgrdPerson.Row, 0)
    
    '���б����ʧ��
    If pstrϵͳ��� <> "" Then
        cgrdPerson.Visible = False
    
        Unload Me
    End If
    
End Sub

Private Sub cgrdPerson_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 And cgrdPerson.Row > 0 Then
        cgrdPerson_DblClick
    End If

End Sub

Private Sub cgrdPerson_LostFocus()
    On Error Resume Next
    If ActiveControl.Name = "ccmdOk" And cgrdPerson.Row > 0 Then
    
    Else
        cgrdPerson.Visible = False
        clblInfo.Visible = False
    End If
End Sub

Private Sub coptChoise_Click(Index As Integer)
    On Error GoTo errHandler
    ctxtName.Enabled = False
    ccmbSex.Enabled = False
    ccmbQueryUnit.Enabled = False
    ccmdLocateUnit.Enabled = False
    ctxtId.Enabled = False
    ctxtHealthNo.Enabled = False
    
    If coptChoise(0).Value Then
        'ѡ������������
        ctxtName.Enabled = True
        ccmbSex.Enabled = True
        ccmbQueryUnit.Enabled = True
        ccmdLocateUnit.Enabled = True
        ctxtName.SetFocus
    ElseIf coptChoise(1).Value Then
        'ѡ�����뽡��������š�
        ctxtHealthNo.Enabled = True
        ctxtHealthNo.SetFocus
    ElseIf coptChoise(2).Value Then
        'ѡ���������֤�š�
        ctxtId.Enabled = True
        ctxtId.SetFocus
    End If
    
    Exit Sub
errHandler:
    'sfsub������ "ְҵ�����沿��", "frm������Ա", "coptChoise_Click", Err.Number, Err.Description, False
End Sub

Private Sub ctxtName_GotFocus()
    On Error Resume Next
    With ctxtName
        .SelStart = 0
        .SelLength = Len(Trim(ctxtName.Text))
    End With
End Sub

Private Sub ctxtName_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ccmbSex.SetFocus
    End If
End Sub
Private Sub ccmbSex_GotFocus()
    On Error Resume Next
    If ccmbSex.Text = "" And ccmbSex.ListCount > 0 Then
        ccmbSex.ListIndex = 0
    End If
End Sub
'���ܣ��Զ������б��
Private Sub ccmbQueryUnit_GotFocus()
    On Error GoTo errHandler
    gfsubShowComboList ccmbQueryUnit
    Exit Sub
errHandler:
End Sub

Private Sub ccmbQueryUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ccmdOk.SetFocus
    End If
End Sub

Private Sub ccmbQueryUnit_LostFocus()
    Dim i As Integer
    
    On Error GoTo errHandler
    
    '�ж�¼��ĵ�λ�Ƿ����б��д��ڣ�������������б�
    i = gffuncItemIsInComboBox(ccmbQueryUnit, ccmbQueryUnit.Text)
    
    If i = -1 Then
        '�ӵ�ccmbQueryUnit�С�
        ccmbQueryUnit.AddItem ccmbQueryUnit.Text
    End If
    
    Exit Sub
errHandler:
    
End Sub

Private Sub ccmbSex_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ccmbQueryUnit.SetFocus
    End If
End Sub
'���õ�λ��λ
Private Sub ccmdLocateUnit_Click()
    On Error GoTo errHandler
    Dim lobjRec As Object  '��λ��λ���صĽ����¼��

    '������λ��λ���档
    Set lobjRec = pobjҵ�����.func��λ��λ
    
    '��ȡ��λ�ĵ�λ����ʾ�ڡ���λ���ơ�¼����С�
    If Not lobjRec Is Nothing Then
        If lobjRec.RecordCount > 0 Then
            ccmbQueryUnit.Text = IIf(IsNull(lobjRec("��λ����")), "", lobjRec("��λ����"))
        End If
    End If
    
    '�ѽ���ص���λ¼���
    ccmbQueryUnit.SetFocus
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����沿��", "frm������Ա", "ccmdLocateUnit_Click", 6666, lstrError, False
End Sub


Private Sub ctxtHealthNo_GotFocus()
    On Error Resume Next
    With ctxtHealthNo
        .SelStart = 0
        .SelLength = Len(Trim(ctxtHealthNo.Text))
    End With
End Sub

Private Sub ctxtHealthNo_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ccmdOk.SetFocus
    End If
End Sub
Private Sub ctxtId_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ccmdOk.SetFocus
    End If
End Sub



Private Function func���ݽ���������Ż�ȡϵͳ���(ByVal para����������� As String) As String
    Dim lobj�����Ա  As Object 'clsPersonExamed.
    Dim lobj��� As Object      'clsMedicalExam
    Dim lstrϵͳ��� As String
    
    func���ݽ���������Ż�ȡϵͳ��� = ""
    
    '��ȡ�������һ������¼��
    '���������Ա����
    Set lobj�����Ա = CreateObject("ְҵ������.clsPersonExamed")
    lobj�����Ա.����������� = para�����������
    Set lobj��� = lobj�����Ա.Func��ȡ�������һ�����
    If Not lobj��� Is Nothing Then
        lstrϵͳ��� = lobj���.ϵͳ���
    Else
        Err.Raise 6666, , "�������Ա��û���ڱ���������������޷��������Ǽǡ���ѡ�����Ǽǡ�"
    End If
            
    func���ݽ���������Ż�ȡϵͳ��� = lstrϵͳ���
    
End Function

Private Sub Form_Load()
    Dim lcolInfo  As Collection
    Dim i As Long
    On Error Resume Next
    
    If pstr������� = "���" Then
       coptChoise(1).Caption = "����֤��"
    Else
       coptChoise(1).Caption = "ϵͳ���"
    End If
    '�ӵ��չ������Ѳ��л�ȡ����¼����ĵ�λ���ơ�
    Set lcolInfo = pobjҵ�����.���չ������䲾.��λ���Ƽ�
    For i = 1 To lcolInfo.Count
        ccmbQueryUnit.AddItem lcolInfo(i)
    Next

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    cgrdPerson.Visible = False
    clblInfo.Visible = False
End Sub
