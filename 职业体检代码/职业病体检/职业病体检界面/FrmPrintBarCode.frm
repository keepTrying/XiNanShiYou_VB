VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Begin VB.Form FrmPrintBarCode 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��ӡ������"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6645
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   6645
   StartUpPosition =   1  '����������
   Begin VSFlex8Ctl.VSFlexGrid cgrdSysNum 
      Height          =   2655
      Left            =   240
      TabIndex        =   8
      Top             =   360
      Width           =   1815
      _cx             =   2088766593
      _cy             =   2088768075
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
      AllowUserResizing=   1
      SelectionMode   =   0
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
      FormatString    =   "�����������    "
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
   Begin VB.CommandButton PrintBarCode 
      Caption         =   "��  ӡ"
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox PrintNumber 
      Height          =   270
      Left            =   3720
      TabIndex        =   1
      Text            =   "1"
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton ccmdExit 
      Caption         =   "��  ��"
      Height          =   495
      Left            =   4200
      TabIndex        =   0
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "��"
      Height          =   255
      Left            =   3840
      TabIndex        =   7
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label clblSysNo_Last 
      Caption         =   "12345678901234"
      Height          =   255
      Left            =   4200
      TabIndex        =   6
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label clblSysNo_First 
      Caption         =   "12345678901234"
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "��"
      Height          =   255
      Left            =   4440
      TabIndex        =   4
      Top             =   645
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "��ӡ"
      Height          =   255
      Left            =   3120
      TabIndex        =   3
      Top             =   645
      Width           =   375
   End
End
Attribute VB_Name = "FrmPrintBarCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'2012-02-25 �ڵ�� ���Ӵ�ӡ�����봰�����Ӧ����
'�����ӡ���������Զ�������Ҫ��ӡ���������
'��������ɹ���ͬϵͳ��ţ�������5λ���Ϊ��1����һ��������û����һλ

Option Explicit
'ע��-->�������� �� ���ϵͳ��� ������ȫ��ͬ��ֻ���ڴ����в�ͬλ�ã����Ʋ�ͬ��
Private pstrϵͳ��� As String           '���������ݿ��е�ϵͳ��ţ�Ҳ�ǽ�������ӡ����������
Public pstrNumbers As Integer           '��¼��һ��д��Ĵ�ӡ�������������ڽ���ϵͳ����˻صĲ���
Private pstr�Ƿ��˻�ϵͳ��� As Boolean '����ӡʧ�ܻ�û�д�ӡ�����˻�ϵͳ��ţ�ֵΪtrue������Ϊfalse��

Private Sub ccmdExit_Click()
    Dim i As Integer
    Dim lobjTmp As Object
    Set lobjTmp = CreateObject("ְҵ������.clsmedicalexam")
    
    'û�д�ӡ�����˻ص�ǰ���ɵ�ϵͳ���
    If pstr�Ƿ��˻�ϵͳ��� = True Then
        For i = cgrdSysNum.Rows - 1 To 1 Step -1
            lobjTmp.Func�˻�ְҵ�����ϵͳ��� (cgrdSysNum.Cell(flexcpText, i, 0))
            cgrdSysNum.RemoveItem (i)
        Next i
    End If
    
    '�˳�����մ��ڶ���
    Unload Me
    Set FrmPrintBarCode = Nothing
End Sub

Private Sub Form_Load()
    Dim lobjTmp As Object
    On Error GoTo errHandler
    
    Set lobjTmp = CreateObject("ְҵ������.clsMedicalExam")
    pstr�Ƿ��˻�ϵͳ��� = True
    pstrNumbers = Val(PrintNumber.Text)
    PrintNumber.TabIndex = 1
    
    'form_loadʱ��Ĭ����ʾ��һ��ϵͳ���
    pstrϵͳ��� = lobjTmp.Func����ְҵ�����ϵͳ���
    clblSysNo_First.Caption = pstrϵͳ���
    cgrdSysNum.AddItem pstrϵͳ���, 1
    sub���ɲ���ʾ����� (1)
    
    Exit Sub
    
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����沿��", "FrmPrintBarCode", "Form_Load", 6666, lstrError, False
End Sub

Private Sub PrintBarCode_Click()
    Dim i As Integer

    On Error GoTo errHandler
    
    '������ӡһ�ι��󣬾Ͳ����޸Ĵ�ӡ������Ҳ���ܼ�����ӡ��[���Ǻܺ�����趨�ɣ�]
    PrintBarCode.Enabled = False
    PrintNumber.Enabled = False
    'For i = 1 To cgrdSysNum.Rows - 1
        'sub��ӡ������������ (cgrdSysNum.TextMatrix(i, 0))
    'Next i
    'pstr�Ƿ��˻�ϵͳ��� = False
    
'2012-04-05 ��¶
'��ӡ�������
    Dim para�������� As Collection
    Set para�������� = New Collection
    For i = 1 To cgrdSysNum.Rows - 1
        para��������.Add (cgrdSysNum.TextMatrix(i, 0))
    Next i
    sub��ӡ����������� para��������
    pstr�Ƿ��˻�ϵͳ��� = False
'2012-04-05 ��¶
    
    Exit Sub

'��������ӡ�������ʾ���˻�ϵͳ��š�
errHandler:
    Dim lobjTmp As Object
    Set lobjTmp = CreateObject("ְҵ������.clsmedicalexam")
    sfsub������ "ְҵ������", "FrmPrintBarCode", "PrintBarCode_Click", Err.Number, Err.Description, False
    For i = cgrdSysNum.Rows - 1 To 1 Step -1
        lobjTmp.Func�˻�ְҵ�����ϵͳ��� (cgrdSysNum.Cell(flexcpText, i, 0))
        cgrdSysNum.RemoveItem (i)
    Next i
    Exit Sub
End Sub

Private Sub PrintNumber_Change()
    Dim i, lobjInt, IfContinue As Integer
    Dim lobjTmp As Object
    Set lobjTmp = CreateObject("ְҵ������.clsmedicalexam")
    
    '�ж������ʽ
    If IsNumeric(PrintNumber.Text) = False Or CLng(Val(PrintNumber.Text)) <= 0 Then
        MsgBox ("��������Ϊ����0������")
        Exit Sub
    End If
    
    '�������̫��ı�ţ��ܱ�ų���9999�Ļ����
    '���ڣ�ÿ�������Ա�� ����9999�Ŀ����ԱȽ�С������û������������޸ģ�����������ʾ��
    If Val(PrintNumber.Text) > 9000 Then IfContinue = MsgBox("һ��������ô���ţ�����ſ��ܲ��㡣Ҫ������", vbYesNo)
    If IfContinue = vbNo Then
        PrintNumber.Text = CStr(pstrNumbers)
        Exit Sub
    End If
    
    '������������˵��֮ǰ�����ϵͳ���û�д�ӡ������Ҫ�˻أ����������±�š�
    For i = Val(pstrNumbers) To 2 Step -1
        lobjTmp.Func�˻�ְҵ�����ϵͳ��� (cgrdSysNum.Cell(flexcpText, i, 0))
        cgrdSysNum.RemoveItem (i)
    Next i
    sub���ɲ���ʾ����� (Val(PrintNumber.Text))
    
End Sub

'���ô����ɵ������������ʾ�ڴ����б���
Sub sub���ɲ���ʾ�����(ByVal paraPrintNum As Integer)
    Dim i As Integer
    Dim lobjTmp As Object
    
    On Error GoTo errHandler
    
    Set lobjTmp = CreateObject("ְҵ������.clsmedicalexam")
    For i = 2 To paraPrintNum
        cgrdSysNum.AddItem lobjTmp.Func����ְҵ�����ϵͳ���, cgrdSysNum.Rows
    Next i
    clblSysNo_Last.Caption = cgrdSysNum.Cell(flexcpText, cgrdSysNum.Rows - 1, 0)
    pstrNumbers = Val(PrintNumber.Text)
    Exit Sub
    
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ������", "frmPrintBarCode", "sub���ɲ���ʾ�����", 6666, lstrError, True
End Sub


