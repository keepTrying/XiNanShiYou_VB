VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Object = "{2327075A-FED2-4AEF-82B8-DD0C2B1AC8E1}#1.2#0"; "dyCatchPhoto.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   ScaleHeight     =   6555
   ScaleWidth      =   8535
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "�����Ա"
      Height          =   495
      Left            =   4560
      TabIndex        =   12
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton Command9 
      Caption         =   "���"
      Height          =   495
      Left            =   4560
      TabIndex        =   10
      Top             =   240
      Width           =   1455
   End
   Begin VSFlex6DAOCtl.vsFlexGrid cgrdResult 
      Height          =   2175
      Left            =   2280
      TabIndex        =   8
      Top             =   4080
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   3836
      _ConvInfo       =   1
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
      AllowUserResizing=   0
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
   End
   Begin dyCatchPhoto.ctlCatchPhoto ccpMain 
      Height          =   3495
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   6165
      BackColor       =   0
      FontSize        =   11.25
   End
   Begin VB.CommandButton Command8 
      Caption         =   "�����Ŀ"
      Height          =   495
      Left            =   6240
      TabIndex        =   6
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton Command7 
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   5
      Top             =   3000
      Width           =   3135
   End
   Begin VB.CommandButton Command6 
      Caption         =   "������"
      Height          =   495
      Left            =   6240
      TabIndex        =   4
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "���ҽʦ"
      Height          =   495
      Left            =   6240
      TabIndex        =   3
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "��켯"
      Height          =   495
      Left            =   4560
      TabIndex        =   2
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "����ģ��"
      Height          =   495
      Left            =   6240
      TabIndex        =   1
      Top             =   2400
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Height          =   2175
      Left            =   240
      ScaleHeight     =   2115
      ScaleWidth      =   1635
      TabIndex        =   0
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Label clblInfo 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   2280
      TabIndex        =   11
      Top             =   3720
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��ȡ�������Ա��Ƭ"
      Height          =   180
      Left            =   240
      TabIndex        =   9
      Top             =   3840
      Width           =   1620
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
    On Error GoTo errHandler
    
    '��ʼ�����ݷ��ʶ���(���ӱ���)��
'    dasubInitialize "Provider=MSDataShape.1;Data Provider=SQLOLEDB.1;Password=welcome;Persist Security Info=True;User ID=user26;Initial Catalog=����2001;Data Source=YANGCHUN"
'    If Not umfuncУ�����("5555", "") Then
'        sffuncMsg "У�����ʧ��5555��", sf����
'    End If
        
    '��ʼ�����ݷ��ʶ���(Tdcserver)��
    dasubInitialize "Provider=MSDataShape.1;Data Provider=SQLOLEDB.1;Password=welcome;Persist Security Info=True;User ID=user26;Initial Catalog=����2001;Data Source=Tdcserver"
    If Not umfuncУ�����("5555", "") Then
        sffuncMsg "У�����ʧ��5555��", sf����
    End If
        
    '��ʼ�����ݷ��ʶ���(Testserver)��
'    dasubInitialize "Provider=MSDataShape.1;Data Provider=SQLOLEDB.1;Password=dyfy;Persist Security Info=True;User ID=sa;Initial Catalog=���������°����ݿ�;Data Source=TESTSERVER"
'    If Not umfuncУ�����("0008", "") Then
'        sffuncMsg "У�����ʧ��0008��", sf����
'    End If
        

    Exit Sub
errHandler:
    sfsub������ "����", "", "Form_load", Err.Number, Err.Description
    Exit Sub
    Resume
End Sub
Private Sub Command9_Click()
    Dim lobj��� As clsMedicalExam
    Dim lobjRec As Object
    Dim lstrTemp As String
    
    Set lobj��� = New clsMedicalExam
    With lobj���
        '��������¼��
        lstrTemp = lobj���.Func����ϵͳ���
        .ϵͳ��� = lstrTemp
        
        '�����������ԡ�
        .������ = P_EXAM_FIRST '���졣
        .�����Ա.���� = "����"
        .�����Ա.��λ���� = "������С��"
        '...
        
        .����.������ = "��ҵ��Ա����"
        Debug.Print .����.�Թܱ����ĸ
        
        Debug.Assert 1 = 2
        
        '����Թܱ����ĸ�Ƿ��ѷ��䡣
        If .����.�Թܱ����ĸ = "" Then
            .����.�Թܱ����ĸ = "A"
        End If
        .�����Ա.���� = ""
        Debug.Assert 1 = 2
        
        .Sub�������Ǽ���Ϣ
        
        '�˶Կ��У���������Ϣ�������Ա��Ϣ���������Ϣ�����Ƿ���������Ӧ��¼��
        Debug.Assert 1 = 2
        
    End With
    
    With lobj���
        '����ϵͳ���Ϊ���д��ڵġ�
        .ϵͳ��� = "123401010402002"
        
        '������ԣ��������ݿ�˶��Ƿ�һ�£���
        '...
        Debug.Assert 1 = 2
        
        '���Է�����
        lstrTemp = .func��ȡϵͳ��ŵ�ǰһ����("00000103280021")
        Debug.Assert 1 = 2
        lstrTemp = .func��ȡϵͳ��ŵĺ�һ����("00000103280021")
        Debug.Assert 1 = 2
        
        lstrTemp = .Func����ϵͳ���
        '��顰������_�����ˮ�ű��Ƿ�ݼӡ�
        Debug.Assert 1 = 2
        
        .sub�˻�ϵͳ��� lstrTemp
        
        '��顰������_�����ˮ�ű��Ƿ�ָ���
        Debug.Assert 1 = 2
        
        
    End With
End Sub



Private Sub Command1_Click()
    On Error GoTo errHandler
   '���������Ա��
    Dim lobjPerson As clsPersonExamed
    Dim lobjRec As Object
    Dim lcolInfo As New Collection
    
    Set lobjPerson = New clsPersonExamed
    
    With lobjPerson
        '���������Ա���ԡ�
        .����������� = .Func���佡���������(lcolInfo)
    
        .���� = "�ź�"
        .������ݺ��� = "510223450608120"
        .��λ���� = "������"
                
        Debug.Assert 1 = 2
        
        .Sub����
        
        Debug.Assert 1 = 2
        
        '��ȡ��Ƭ��
        Picture1.Picture = .��Ƭ
        
        Debug.Assert 1 = 2
        
        '��ȡ�������һ����졣
        Set lobjRec = .Func��ȡ�������һ�����
        
        '��ʾ��ѯ�����
        gfsubLoadGridFromRec cgrdResult, lobjRec
    End With

    Exit Sub
errHandler:
    sfsub������ "����1", "Form1", "Command1_Click", Err.numer, Err.Description, False
    Exit Sub
    Resume
End Sub

Private Sub Command2_Click()
    On Error GoTo errHandler
    Dim lobj���� As ClsMedicalExamSheet
    
    Set lobj���� = New ClsMedicalExamSheet
    With lobj����
        .ϵͳ��� = "11111200103200004"
        
        Debug.Assert 1 = 2
        '������ԡ�
        '...
        
        '����ѡ������
        .������ = "��ҵ��Ա���"
        Debug.Assert 1 = 2
        
        '���������
        .Sub������� "0101", "100", "1234", Date
                
        Debug.Assert 1 = 2
        
        '�����������
        .Sub���������
        
    End With
    
    Exit Sub
errHandler:
    sfsub������ "����1", "Form1", "Command2_Click", Err.Number, Err.Description, False
End Sub

Private Sub Command3_Click()
    On Error GoTo errHandler
    Dim lobj����ģ�� As clsMedicalExamTemplate
    
    Set lobj����ģ�� = New clsMedicalExamTemplate
    With lobj����ģ��
        '�������ԡ�
        .������ = "�����Ѹ����԰�"
        .���� = 9
        .��쵥���� = "�����Թ���쵥"
        .�Թ���ĸ��� = "B"
        .�Ƿ񸴲����� = True
        .�շѱ�׼ = "�����Ѹ����԰��շ�"
        
        .Sub��Ӹ�����Ŀ "��������", True
        .Sub��Ӹ�����Ŀ "Ƭ��", True
        
        
        .Sub��������Ŀ "0001"
        .Sub��������Ŀ "0002"
                
        .Sub��������� 3435
        
        Debug.Assert 1 = 2
        .Sub����ģ��
        
        Debug.Assert 1 = 2
        .Sub����ģ�� "�����Ѹ����԰�"
        .���� = 10
        .Sub����ģ��
        
        Debug.Assert 1 = 2
        .Subɾ��ģ��
        
    End With
    Exit Sub
errHandler:
    sfsub������ "����1", "Form1", "Command3_Click", Err.Number, Err.Description, False
End Sub

Private Sub Command4_Click()
    Dim lobj��켯 As clsMedicalExamSet
    Dim lobjRec As Object
    
    Set lobj��켯 = New clsMedicalExamSet
    
    With lobj��켯
        .��������� = "2001-3-01"
        .��������� = "2001-4-01"
    
        '.��λ���� = "111"
    
        '.���Թܱ�� = "A:0001"
        '.���Թܱ�� = "A:0010"
    
        '��ȡ��Ҫ����ġ�
        '.�����־ = 1
    
        '��ȡ��Ҫ���鵫��δ����ġ�
        .�����־ = 1
        .����ϵͳ��� = ""
    End With
        
    '��ȡ��ѯ�����
    Set lobjRec = lobj��켯.Ԫ�ؼ�
    
    '��ʾ��ѯ�����
    clblInfo = "��������ڣ�" & lobj��켯.��������� & "������" & lobj��켯.��������� & "����Ҫ���鵫��δ���������¼"
    
    gfsubLoadGridFromRec cgrdResult, lobjRec
    
End Sub

Private Sub Command5_Click()
    '�������ҽʦ��
    Dim lobj���ҽʦ As ClsMedicalExaminer
    Dim lblnCan As Boolean
    Dim lcolInfo As Collection
    
    Set lobj���ҽʦ = New ClsMedicalExaminer
    With lobj���ҽʦ
        .��� = "5555"
        
        '��ȡ���������Ŀ��
        Set lcolInfo = .���������Ŀ
        
        Debug.Assert 1 = 2
'        .Sub��������Ŀ "0101"
'        .Sub��������Ŀ "0102"
'        .Sub��������Ŀ "0201"
'        .Subɾ�������Ŀ "0201"
'        .Sub��������Ŀ "0202"
'
'        Set lcolInfo = .���������Ŀ
       
        lblnCan = .func�Ƿ������Ŀ("0101")
        Debug.Assert 1 = 2
        
        lblnCan = .func�Ƿ������Ŀ("0201")
        Debug.Assert 1 = 2
        
        Set lcolInfo = .Func��ȡ����ָ�������Ͽ����������Ŀ("11111200103200001", "����")
        Debug.Assert 1 = 2
    End With

End Sub

Private Sub Command6_Click()
    '���������ۡ�
    Dim lobj������ As ClsMedicalExamConclusion
    Dim lobj���������� As ClsConclusionFilter
    Dim lcolInfo As Collection
    Dim lbln���� As Boolean
    Set lobj���������� = New ClsConclusionFilter
    With lobj����������
        .ID = 3406
        .��� = 2
        .SubAddFilter 1, "0001", "��", "=", "�쳣"
        .SubAddFilter 2, "0001", "��", "=", "�쳣"
        
        .subSave
        Debug.Assert 1 = 2

        .SubRemoveFilter 2
        .subSave
        Debug.Assert 1 = 2

        lbln���� = .Func�ж��Ƿ���������("11111200103200001")
        Debug.Assert 1 = 2
        
    End With
    
    Set lobj������ = New ClsMedicalExamConclusion
    With lobj������
        .ID = 3406
        Set lobj���������� = .�ж�����(1)
        Debug.Assert 1 = 2
        
        .Subɾ���������� 1
        Set lcolInfo = .�����ж�����
        Debug.Assert 1 = 2
        
        lbln���� = .Func�ж��Ƿ���±�����("11111200103200001")
        Debug.Assert 1 = 2
    End With


End Sub

Private Sub Command7_Click()
    '�������������
    Dim lobj������ As clsManageMedicalExam
    Dim lobjRec As Object
    Dim lcolInfo As Collection
    Dim lstrTemp As String
    
    Set lobj������ = New clsManageMedicalExam
    With lobj������
        Set lobjRec = .���չ������䲾
        '����ȡ�����ԡ�
        Debug.Assert 1 = 2
        
        Set lobjRec = .����վ����
        '����ȡ�����ԡ�
        Debug.Assert 1 = 2
    
        Set lcolInfo = .����ҵ������
        '����ȡ�����ԡ�
        Debug.Assert 1 = 2
        
        Set lobjRec = .������츽����Ŀ
        '����ȡ�����ԡ�
        Debug.Assert 1 = 2
        
        Set lobjRec = .��������������
        '����ȡ�����ԡ�
        Debug.Assert 1 = 2
      
        Set lcolInfo = .��������շѱ�׼
        '����ȡ�����ԡ�
        Debug.Assert 1 = 2
      
        lstrTemp = .ҵ������("�Ƿ��շ�")
        '����ȡ�����ԡ�
        Debug.Assert 1 = 2
        .Sub�޸�ҵ������ "�Ƿ��շ�", "��"
        .Sub�޸�ҵ������ "�Ƿ�����", "��"
        .Sub�޸�ҵ������ "��������", "30"
        .Sub�޸�ҵ������ "�Ƿ��ӡ��쵥", "��"
        
        '�����У�ҵ�����ñ��Ƿ����޸ġ�
        Debug.Assert 1 = 2
        
        '�����Ŀ[��Ŀ����,¼�����,��������,���ݳ���,ö��ֵ]
        Set lcolInfo = New Collection
        With lcolInfo
            .Add "���֤��", "��Ŀ����"
            .Add "���֤��", "¼�����"
            .Add 3, "��������"
            .Add "20", "���ݳ���"
            .Add "", "ö��ֵ"
        End With
        .Sub������츽����Ŀ 1, lcolInfo
        '�����У������Ա������Ŀ���ñ�
        Debug.Assert 1 = 2
        
        '�޸���Ŀ��
        Set lcolInfo = New Collection
        With lcolInfo
            .Add "�Ա�", "��Ŀ����"
            .Add "�Ա�", "¼�����"
            .Add 3, "��������"
            .Add "6", "���ݳ���"
            .Add "�Ա��ֵ�", "ö��ֵ"
        End With
        .Sub������츽����Ŀ 2, lcolInfo, "�Ա�"
        '�����У������Ա������Ŀ���ñ�
        Debug.Assert 1 = 2
        
        'ɾ����Ŀ��
        .Sub������츽����Ŀ 3, lcolInfo, "�Ա�"
        '�����У������Ա������Ŀ���ñ�
        Debug.Assert 1 = 2
        
        Set lobjRec = .Func��ȡ���޸ĵ�����¼("", "", "")
        '����ȡ�����ݡ�
        Debug.Assert 1 = 2
        
        Set lobjRec = .Func��ȡ��������ȷ��������¼("", "", "")
        '����ȡ�����ݡ�
        Debug.Assert 1 = 2
        Set lobjRec = .Func��ȡ��Ҫ���������¼()
        
        '����ȡ�����ݡ�
        Debug.Assert 1 = 2
        Set lobjRec = .Func��ȡ���½��۵�δȷ��������¼("", "", "", "")
        lstrTemp = .Func���ݽ���֤����Ż�ȡ���ϵͳ���("")
        
        Debug.Assert 1 = 2
        
        'û�е�λ����ӿڣ��˷�����ʱ���ܲ��ԡ�
        'Set lobjRec = .func��λ��λ()
        
        Dim lobj��� As clsMedicalExam
        
        Set lobj��� = New clsMedicalExam
        lstrTemp = lobj���.Func����ϵͳ���
        lobj���.ϵͳ��� = lstrTemp
        lobj���.������ = P_EXAM_FIRST
        lobj���.����.������ = "��ҵ��Ա���"
        lobj���.�����Ա.���� = "���"
        lobj���.�����Ա.��λ���� = "�����"
        lobj���.�����Ա.������ݺ��� = "510223470812110"
        lobj���.�����Ա.��Ƭ = Picture1.Picture
        
        lobj������.Sub���Ǽ� lobj���
        '���������ݡ�
        Debug.Assert 1 = 2
        
        Dim lcolResult As Collection
        Dim lcolItem As Collection
        
        Set lcolInfo = New Collection
        lcolInfo.Add "11111200103200001"
        lcolInfo.Add "11111200103200002"
        
        '[�����Ŀ�������]
        Set lcolResult = New Collection
        Set lcolItem = New Collection
        lcolItem.Add "0001", "�����Ŀ"
        lcolItem.Add "����", "�����"
        lcolResult.Add lcolItem, lcolItem("�����Ŀ")
        Set lcolItem = New Collection
        lcolItem.Add "0002", "�����Ŀ"
        lcolItem.Add "����", "�����"
        lcolResult.Add lcolItem, lcolItem("�����Ŀ")
        
        .Sub��д����� lcolInfo, lcolResult, "1234", Date
        '���������ݡ�
        Debug.Assert 1 = 2
        
        .Subȷ�������� "11111200103200002", "�Ҹ�", "�������", "", "", False
        '���������ݡ�
        Debug.Assert 1 = 2
        
        .Subȡ�������� "11111200103200001"
        '���������ݡ�
        Debug.Assert 1 = 2
        
        '.Sub��ӡ����
        
    End With

End Sub

Private Sub Command8_Click()
    Dim lobj�����Ŀ As ClsTestItem
    Dim lobjRec As Object
    
    Set lobj�����Ŀ = New ClsTestItem
    With lobj�����Ŀ
        .���� = "0010"
        .���� = "������"
        .ȱʡֵ = "20"
        .������ = 4
        .���� = "����"
        .ö����Դ = ""
        .subSave
        Debug.Assert 1 = 2
        
        .subDelete
        Debug.Assert 1 = 2
    End With
    
    '���������Ŀ��
    Dim lobj�����Ŀ�� As clsTestItemSet
    Set lobj�����Ŀ�� = New clsTestItemSet
    With lobj�����Ŀ��
        '.������ = 2
        .���� = "����"
        '.�����Ŀ���� = "0001"
        Set lobjRec = .�����Ŀ
        Debug.Assert 1 = 2
    End With
    
End Sub

