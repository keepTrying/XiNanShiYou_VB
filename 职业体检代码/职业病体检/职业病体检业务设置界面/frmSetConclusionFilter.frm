VERSION 5.00
Object = "{8099FCC2-0A81-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6.ocx"
Object = "{F48D4DEC-1198-11D5-91BE-0050BA06B70C}#5.9#0"; "¼��ؼ�.ocx"
Begin VB.Form frmSetConclusionFilter 
   Caption         =   "�������ж���������"
   ClientHeight    =   7800
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   11070
   Icon            =   "frmSetConclusionFilter.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   11070
   Begin VB.CommandButton ccmdExit 
      Caption         =   "�˳�(&X)"
      Height          =   400
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2040
      Width           =   1275
   End
   Begin VB.CommandButton ccmdUpdate 
      Caption         =   "�޸�(&M)"
      Enabled         =   0   'False
      Height          =   426
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   840
      Width           =   1275
   End
   Begin VB.CommandButton ccmdAdd 
      Caption         =   "����(&A)"
      Height          =   426
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Width           =   1275
   End
   Begin VB.CommandButton ccmdDelete 
      Caption         =   "ɾ��(&D)"
      Enabled         =   0   'False
      Height          =   426
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "������¼��һ�������۵��ж�����"
      ClipControls    =   0   'False
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   10815
      Begin VB.CommandButton ccmdRemoveRow 
         Caption         =   "ɾ����(&R)"
         Enabled         =   0   'False
         Height          =   427
         Left            =   9360
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3000
         Width           =   1275
      End
      Begin VB.ComboBox ccmbOperator 
         Height          =   300
         Left            =   5500
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton ccmdOk 
         Caption         =   "����(&S)"
         Enabled         =   0   'False
         Height          =   427
         Left            =   9360
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1800
         Width           =   1275
      End
      Begin VB.CommandButton ccmdCancel 
         Caption         =   "ȡ��(&C)"
         Enabled         =   0   'False
         Height          =   427
         Left            =   9360
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2400
         Width           =   1275
      End
      Begin ¼��ؼ�.ctlInputGrid cgrdInput 
         Height          =   1935
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   3413
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Cols            =   3
         Rows            =   1
         Count           =   0
         Rows            =   1
         Cols            =   3
      End
      Begin ¼��ؼ�.ctlInputFrame cifFilter 
         Height          =   1725
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   3043
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   ""
         Rows            =   2
         Cols            =   17
         DistanceofRow   =   0
         BorderStyle     =   0
         FormatString    =   "���,1,0,2,������,1,3,8,����,1,12,8,���,2,0,1,�����Ŀ,2,2,8,�ж�����,2,11,2,�ж�ֵ,2,14,4"
         Count           =   7
         titleInputBox0001=   "���"
         statusinfoInputBox0001=   ""
         lengthInputBox0001=   2
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
         ����InputBox0001=   "���"
         ȱʡֵInputBox0001=   ""
         ����ȱʡֵInputBox0001=   ""
         ����InputBox0001=   4
         MaxInputBox0001 =   ""
         MinInputBox0001 =   ""
         VisibleInputBox0001=   -1  'True
         PermitNullInputBox0001=   -1  'True
         TriggerstrInputBox0001=   ""
         EnableInputBox0001=   0   'False
         �����ѡInputBox0001=   0   'False
         titleInputBox0002=   "������"
         statusinfoInputBox0002=   ""
         lengthInputBox0002=   8
         orderInputBox0002=   2
         valueInputBox0002=   ""
         datatypeInputBox0002=   3
         colInputBox0002 =   3
         rowInputBox0002 =   1
         PassWordCharInputBox0002=   0   'False
         ����InputBox0002=   0   'False
         ����������ֵInputBox0002=   0   'False
         ���������СֵInputBox0002=   0   'False
         �ֵ�����InputBox0002=   "�������ֵ�"
         ��ʾ�ֵ��ֶ�InputBox0002=   "����"
         �����ֵ��ֶ�InputBox0002=   "InnerID"
         ����InputBox0002=   "������"
         ȱʡֵInputBox0002=   ""
         ����ȱʡֵInputBox0002=   ""
         ����InputBox0002=   50
         MaxInputBox0002 =   ""
         MinInputBox0002 =   ""
         VisibleInputBox0002=   -1  'True
         PermitNullInputBox0002=   0   'False
         TriggerstrInputBox0002=   ""
         CheckInDictInputBox0002=   -1  'True
         �����ѡInputBox0002=   0   'False
         titleInputBox0003=   "����"
         statusinfoInputBox0003=   ""
         lengthInputBox0003=   8
         orderInputBox0003=   3
         valueInputBox0003=   ""
         datatypeInputBox0003=   3
         colInputBox0003 =   12
         rowInputBox0003 =   1
         PassWordCharInputBox0003=   0   'False
         ����InputBox0003=   0   'False
         ����������ֵInputBox0003=   0   'False
         ���������СֵInputBox0003=   0   'False
         �ֵ�����InputBox0003=   ""
         ��ʾ�ֵ��ֶ�InputBox0003=   ""
         �����ֵ��ֶ�InputBox0003=   ""
         ����InputBox0003=   "����"
         ȱʡֵInputBox0003=   ""
         ����ȱʡֵInputBox0003=   ""
         ����InputBox0003=   100
         MaxInputBox0003 =   ""
         MinInputBox0003 =   ""
         VisibleInputBox0003=   -1  'True
         PermitNullInputBox0003=   -1  'True
         TriggerstrInputBox0003=   ""
         �����ѡInputBox0003=   0   'False
         titleInputBox0004=   "���"
         statusinfoInputBox0004=   ""
         lengthInputBox0004=   1
         orderInputBox0004=   4
         valueInputBox0004=   ""
         datatypeInputBox0004=   2
         colInputBox0004 =   0
         rowInputBox0004 =   2
         PassWordCharInputBox0004=   0   'False
         ����InputBox0004=   0   'False
         ����������ֵInputBox0004=   0   'False
         ���������СֵInputBox0004=   -1  'True
         �ֵ�����InputBox0004=   ""
         ��ʾ�ֵ��ֶ�InputBox0004=   ""
         �����ֵ��ֶ�InputBox0004=   ""
         ����InputBox0004=   "���"
         ȱʡֵInputBox0004=   ""
         ����ȱʡֵInputBox0004=   ""
         ����InputBox0004=   4
         MaxInputBox0004 =   ""
         MinInputBox0004 =   "1"
         VisibleInputBox0004=   -1  'True
         PermitNullInputBox0004=   0   'False
         TriggerstrInputBox0004=   ""
         �����ѡInputBox0004=   0   'False
         titleInputBox0005=   "�����Ŀ"
         statusinfoInputBox0005=   ""
         lengthInputBox0005=   8
         orderInputBox0005=   5
         valueInputBox0005=   ""
         datatypeInputBox0005=   3
         colInputBox0005 =   2
         rowInputBox0005 =   2
         PassWordCharInputBox0005=   0   'False
         ����InputBox0005=   0   'False
         ����������ֵInputBox0005=   0   'False
         ���������СֵInputBox0005=   0   'False
         �ֵ�����InputBox0005=   "�����Ŀ�ֵ�"
         ��ʾ�ֵ��ֶ�InputBox0005=   "����"
         �����ֵ��ֶ�InputBox0005=   "����"
         ����InputBox0005=   "�����Ŀ"
         ȱʡֵInputBox0005=   ""
         ����ȱʡֵInputBox0005=   ""
         ����InputBox0005=   50
         MaxInputBox0005 =   ""
         MinInputBox0005 =   ""
         VisibleInputBox0005=   -1  'True
         PermitNullInputBox0005=   0   'False
         TriggerstrInputBox0005=   ""
         CheckInDictInputBox0005=   -1  'True
         �����ѡInputBox0005=   0   'False
         titleInputBox0006=   "�ж�����"
         statusinfoInputBox0006=   ""
         lengthInputBox0006=   2
         orderInputBox0006=   6
         valueInputBox0006=   ""
         datatypeInputBox0006=   0
         colInputBox0006 =   11
         rowInputBox0006 =   2
         PassWordCharInputBox0006=   0   'False
         ����InputBox0006=   0   'False
         ����������ֵInputBox0006=   0   'False
         ���������СֵInputBox0006=   0   'False
         �ֵ�����InputBox0006=   ""
         ��ʾ�ֵ��ֶ�InputBox0006=   ""
         �����ֵ��ֶ�InputBox0006=   ""
         ����InputBox0006=   "�ж�����"
         ȱʡֵInputBox0006=   "="
         ����ȱʡֵInputBox0006=   "="
         ����InputBox0006=   10
         MaxInputBox0006 =   ""
         MinInputBox0006 =   ""
         VisibleInputBox0006=   -1  'True
         PermitNullInputBox0006=   0   'False
         TriggerstrInputBox0006=   ""
         �����ѡInputBox0006=   0   'False
         titleInputBox0007=   "�ж�ֵ"
         statusinfoInputBox0007=   ""
         lengthInputBox0007=   4
         orderInputBox0007=   7
         valueInputBox0007=   ""
         datatypeInputBox0007=   3
         colInputBox0007 =   14
         rowInputBox0007 =   2
         PassWordCharInputBox0007=   0   'False
         ����InputBox0007=   0   'False
         ����������ֵInputBox0007=   0   'False
         ���������СֵInputBox0007=   0   'False
         �ֵ�����InputBox0007=   ""
         ��ʾ�ֵ��ֶ�InputBox0007=   ""
         �����ֵ��ֶ�InputBox0007=   ""
         ����InputBox0007=   "�ж�ֵ"
         ȱʡֵInputBox0007=   ""
         ����ȱʡֵInputBox0007=   ""
         ����InputBox0007=   50
         MaxInputBox0007 =   ""
         MinInputBox0007 =   ""
         VisibleInputBox0007=   -1  'True
         PermitNullInputBox0007=   0   'False
         TriggerstrInputBox0007=   ""
         �����ѡInputBox0007=   0   'False
         ErrColor        =   12648447
      End
      Begin ¼��ؼ�.ctlInputDictGrid cidgMain 
         Height          =   3495
         Left            =   6120
         TabIndex        =   13
         Top             =   240
         Visible         =   0   'False
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   6165
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
   End
   Begin VSFlex6Ctl.vsFlexGrid cgrdMain 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   9255
      _cx             =   23740357
      _cy             =   23728503
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
      BackColor       =   16777215
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
      Rows            =   1
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   "^��� |^������                     |^����                                 "
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "      (3) һ�������������������г������Σ���ʾ�ý��������ֶ������жϷ�����"
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   16
      Top             =   7440
      Width           =   6840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "      (2) һ�������۵�һ���ж����������ɶ����������ɣ�����������һ�д���һ������������������֮���ǲ��ҹ�ϵ��"
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   15
      Top             =   7200
      Width           =   10080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "˵����(1) ���������۵��ж�������ҽʦ����д��������Ա���������Ŀ�Ľ����ϵͳ��ݴ������Զ��ó������ۡ�"
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   6960
      Width           =   10080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���е��������жϷ��飺"
      Height          =   180
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2160
   End
End
Attribute VB_Name = "frmSetConclusionFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'���ߣ��

Private WithEvents mobj����ͨ�ö��� As cls����ͨ�ö���
Attribute mobj����ͨ�ö���.VB_VarHelpID = -1

Private mobj���������� As Object 'clsConclusionFilter

Private Sub cgrdMain_AfterSort(ByVal Col As Long, Order As Integer)
    On Error Resume Next
    cgrdMain.Row = 0
    ccmdDelete.Enabled = False
    ccmdUpdate.Enabled = False
    
End Sub

'���,1,0,2,������,1,3,8,����,1,12,8,���,2,0,1,�����Ŀ,2,2,8,�ж�����,2,11,3,�ж�ֵ,2,15,3
Private Sub Form_Load()
    Dim lobjRec As Object
    Dim lcolInfo As New Collection
    Dim lcolItem As Collection
    Dim i As Long
    
    On Error GoTo errHandler
    
    '��ȡ����ʾ�����������ж��������������
    Set lobjRec = pobjҵ�����.��������������
    gfsubLoadGridFromRec cgrdMain, lobjRec
    If cgrdMain.Rows > 1 Then cgrdMain.Rows = cgrdMain.Rows - 1
    
    '����ȫ�ֱ���mobj����ͨ�ö���
    Set mobj����ͨ�ö��� = New cls����ͨ�ö���
    
    With mobj����ͨ�ö���
        Set .Form = Me
        Set .c¼��� = cifFilter
        Set .c��¼�� = cgrdInput
        Set .c�ֵ�� = cidgMain
        .pint��ϸ��Ϣ��ʼ��� = 4
        
        .subInitialize lcolInfo, ""
    End With
    
    '����ȫ�ֱ�����mobj��������������
    Set mobj���������� = CreateObject("ְҵ������.clsConclusionFilter")
        
    '��ȡ���п�ѡ���ж�������[���ţ�˵��]��
    Set lcolInfo = mobj����������.�ж�����ö��
    
    '��ȡ�������ֵ䡣
    Set lobjRec = pobjDict.Fetch("�������ֵ���ͼ")
    
    '����¼��塰�����ۡ�¼�����ֵ����ݡ�
    Set cifFilter.InfoCollection(2).DictRecordSet = lobjRec
    
    '��ʼ�����¼�����ж������ֵ䡣
    ccmbOperator.Clear
    For i = 1 To lcolInfo.Count
        ccmbOperator.AddItem lcolInfo(i)("����")
    Next
    
    '��ȡ���������Ŀ��
    Dim lobj�����Ŀ�� As Object
    Set lobj�����Ŀ�� = CreateObject("ְҵ������.clsTestItemSet")
    Set lobjRec = lobj�����Ŀ��.�����Ŀ
    
    '����¼��塰�����Ŀ��¼�����ֵ����ݡ�
    Set cifFilter.InfoCollection(5).DictRecordSet = lobjRec
    
    cifFilter.Enabled = False
    cgrdInput.Enabled = False
    ccmdOk.Enabled = False
    ccmdCancel.Enabled = False
    ccmdRemoveRow.Enabled = False
    
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����ý���", "frmSetConclusionFilter", "Form_Load", 6666, lstrError, False
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 39 Or Chr(KeyAscii) = "��" Then
        '���������롰'���͡�������
        KeyAscii = 0
    End If

End Sub
Private Sub ccmbOperator_Click()
    On Error Resume Next
    cifFilter.ItemText(5) = ccmbOperator.Text
End Sub

Private Sub ccmbOperator_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        cifFilter.SetFocus
        cifFilter.ItemSetfocus 6
    ElseIf KeyCode = vbKeyTab Then
        cifFilter.SetFocus
        cifFilter.ItemSetfocus 6
    ElseIf KeyCode = vbKeyTab And (Shift And vbShiftMask = vbShiftMask) Then
        cifFilter.ItemSetfocus 5
    End If
End Sub

Private Sub ccmbOperator_LostFocus()
    On Error Resume Next
    ccmbOperator.Visible = False
End Sub

Private Sub ccmdRemoveRow_Click()
    Dim lblnSuc As Boolean 'ɾ���Ƿ�ɹ���
    On Error GoTo errHandler
    
    'ɾ������ġ�
    If cgrdInput.Row > 0 And Val(cifFilter.ItemText(3)) <> 0 Then
        mobj����������.subRemoveFilter cifFilter.ItemText(3)
    
        'ɾ�����������ϵġ�
        mobj����ͨ�ö���.subOperate optDELETE, lblnSuc
    Else
        sffuncMsg "����ѡ��Ҫɾ�����У�"
    End If
    
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����ý���", "frmSetConclusionFilter", "ccmdRemoveRow_Click", 6666, lstrError, False
End Sub

Private Sub cifFilter_ItemGetFocus(Index As Integer)
    On Error GoTo errHandler
    If Index = 5 Then
        ccmbOperator.Visible = True
        ccmbOperator.SetFocus
    End If
    Exit Sub
errHandler:
End Sub
Private Sub cifFilter_ItemLostFocus(Index As Integer)
    On Error Resume Next
    If Index = 1 Then
        mobj����ͨ�ö���_ItemLostFocus Index, "������", cifFilter.ItemText(Index), cifFilter.ItemTrueText(Index), False
    End If
End Sub


Private Sub ccmdCancel_Click()
    On Error GoTo errHandler
    
    '�����޸ģ�������¼������ʾcgrdMain��ǰ�����ݣ������½������¼������
    If cgrdMain.Row > 0 Then
        cgrdMain_Click
    Else
        subClear
    End If
    
    '����"subReset"�ָ����档
    subReset
    
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����ý���", "frmSetConclusionFilter", "ccmdCancel_Click", 6666, lstrError, False
End Sub

Private Sub ccmdDelete_Click()
    On Error GoTo errHandler
    'ѯ�ʡ�
    If sffuncMsg("��ȷ��Ҫɾ����" & cgrdMain.TextMatrix(cgrdMain.Row, 2) & "�����ж�������", sfѯ��) Then
    
        'ɾ����ǰѡ�ű�ŷ���������ж�������
        mobj����������.subDelete
        
        '��ս��档
        subClear
        cgrdMain.RemoveItem cgrdMain.Row
        
        ccmdDelete.Enabled = False
        ccmdUpdate.Enabled = False
    End If
    
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����ý���", "frmSetConclusionFilter", "ccmdDelete_Click", 6666, lstrError, False
End Sub

Private Sub ccmdOk_Click()
    Dim lblnIsExist As Boolean '��־��ǰ�����Ƿ����½���
    Dim llngRow As Long        'cgrdMain����ǰѡ�е��кš�
    
    On Error GoTo errHandler
    
    If mobj����������.ID <> 0 And mobj����������.�ж�����.Count > 0 Then
        lblnIsExist = mobj����������.�Ƿ��Ѵ���
        '���浱ǰ�����޸Ļ�������������
        mobj����������.���� = cifFilter.ItemText(2)
        mobj����������.subSave
    
        '�����½�����cgrdMain����ʾ�½����������������������
        '���򣬰���¼����Ϣ�޸�cgrdMain��ǰ�У�
        If lblnIsExist Then
            '�޸��С�
            llngRow = cgrdMain.Row
            cgrdMain.TextMatrix(llngRow, 1) = cifFilter.ItemTrueText(1)
            cgrdMain.TextMatrix(llngRow, 2) = cifFilter.ItemText(1)
            cgrdMain.TextMatrix(llngRow, 3) = cifFilter.ItemText(2)
        Else
            cifFilter.ItemText(0) = mobj����������.���
            '����С�
            cgrdMain.AddItem cifFilter.ItemText(0) & vbTab & cifFilter.ItemTrueText(1) & vbTab & cifFilter.ItemText(1) & vbTab & cifFilter.ItemText(2)
            If cgrdMain.Row > 0 Then
                cgrdMain_Click
            End If
                
        End If
    Else
        Err.Raise 6666, , "ϵͳ�޷����棡��Ϊ��" & Chr(13) & Chr(10) & "����ѡ�������ۡ�¼���ж�������������¼�����һ����������һ��Ҫ�����һ��س�������֤¼�������������������������С�"
    End If
    
    '�ָ����档
    subReset
    
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����ý���", "frmSetConclusionFilter", "ccmdOk_Click", 6666, lstrError, False
End Sub

Private Sub ccmdUpdate_Click()
    On Error GoTo errHandler
    '����¼�������ã�������ֻ��"ȷ��"��ȡ�����˳���ť���ã������Զ���"������"¼���
    subBeginEdit
    
    '�����۲����޸ġ�
    cifFilter.ItemEnable(1) = False
    cifFilter.ItemEnable(2) = True
    
    '�����Զ��䵽"����"¼���
    cifFilter.ItemSetfocus 3
    
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����ý���", "frmSetConclusionFilter", "ccmdUpdate_Click", 6666, lstrError, False
End Sub

Private Sub cgrdMain_Click()

    On Error GoTo errHandler
    If cgrdMain.Row > 0 Then
        ccmdDelete.Enabled = True
    
        subShowHistory
    
        ccmdUpdate.Enabled = True
    Else
        ccmdDelete.Enabled = False
        ccmdUpdate.Enabled = False
    End If
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����ý���", "frmSetConclusionFilter", "cgrdMain_Click", 6666, lstrError, False
End Sub
Private Sub subShowHistory()
    Dim llng��� As Long
    Dim lcol���� As Collection
    Dim lcolFilterItem As Variant
    
    Dim lcol������Ϣ As Collection
    
    Dim lcol��ϸ���� As Collection
    Dim lcolRow As Collection
    Dim lcolItem As Collection
    
    On Error GoTo errHandler
    
    If cgrdMain.Row < 1 Then Exit Sub
    cgrdInput.Rows = 1
    cifFilter.ClearContent
    
    '����"mobj����������"������"ID"��"���"��
    llng��� = cgrdMain.TextMatrix(cgrdMain.Row, 0)
    mobj����������.subClear
    mobj����������.��� = llng���
    mobj����������.ID = cgrdMain.TextMatrix(cgrdMain.Row, 1)
    
    '��ȡ��ǰ���������������
    Set lcol���� = mobj����������.�ж�����
    
    '���ݶ���������¼������ʾ�������۵��ж�������
    Set lcol������Ϣ = New Collection
    Set lcolItem = New Collection
    lcolItem.Add llng���, "��ʾ����"
    lcolItem.Add llng���, "��������"
    lcol������Ϣ.Add lcolItem, "���"
    Set lcolItem = New Collection
    lcolItem.Add mobj����������.����������, "��ʾ����"
    lcolItem.Add mobj����������.ID, "��������"
    lcol������Ϣ.Add lcolItem, "������"
    Set lcolItem = New Collection
    lcolItem.Add mobj����������.����, "��ʾ����"
    lcolItem.Add mobj����������.����, "��������"
    lcol������Ϣ.Add lcolItem, "����"
    
    '��������ͨ��¼��ؼ���Ҫ����뼯���С�
    Set lcol��ϸ���� = New Collection
    For Each lcolFilterItem In lcol����
        Set lcolRow = New Collection
        Set lcolItem = New Collection
        lcolItem.Add lcolFilterItem("���"), "��ʾ����"
        lcolItem.Add lcolFilterItem("���"), "��������"
        lcolRow.Add lcolItem, "���"
        Set lcolItem = New Collection
        lcolItem.Add lcolFilterItem("�����Ŀ����"), "��ʾ����"
        lcolItem.Add lcolFilterItem("�����Ŀ"), "��������"
        lcolRow.Add lcolItem, "�����Ŀ"
        Set lcolItem = New Collection
        lcolItem.Add lcolFilterItem("�ж�����"), "��ʾ����"
        lcolItem.Add lcolFilterItem("�ж�����"), "��������"
        lcolRow.Add lcolItem, "�ж�����"
        Set lcolItem = New Collection
        lcolItem.Add lcolFilterItem("��׼ֵ"), "��ʾ����"
        lcolItem.Add lcolFilterItem("��׼ֵ"), "��������"
        lcolRow.Add lcolItem, "�ж�ֵ"
        
        lcol��ϸ����.Add lcolRow
    Next
    
    '��cgrdInput����ʾ�����Ѵ��ڵ�������
    Set mobj����ͨ�ö���.pcol������Ϣ = lcol������Ϣ
    Set mobj����ͨ�ö���.pcol��ϸ��Ϣ = lcol��ϸ����
    mobj����ͨ�ö���.sub�Ѽ�����������¼���
    
    '�޸ġ�ɾ����ť���á�
    ccmdUpdate.Enabled = True
    ccmdDelete.Enabled = True
    
    Exit Sub
errHandler:
    sfsub������ "ְҵ�����ý���", "frmSetConclusionFilter", "cgrdMain_Click", Err.Number, Err.Description, True
End Sub

Private Sub ccmdAdd_Click()
    On Error GoTo errHandler
    
    '���ö��mobj����������.��š�Ϊ����ֵ0����ʾ��ʼ������
    cgrdMain.Row = 0
    mobj����������.��� = 0
    
    '���¼������
    subClear
    
    '���ý�����ֻ��"ȷ��"��"ȡ��"��"�˳�"��ť���ã�����¼�������á�
    subBeginEdit
    
    '�����ۿ���¼�롣
    cifFilter.ItemEnable(1) = True
    cifFilter.ItemEnable(2) = True
    
    '���㵽�������ۡ���
    cifFilter.ItemSetfocus 1
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����ý���", "frmSetConclusionFilter", "ccmdAdd_Click", 6666, lstrError, False
End Sub

Private Sub ccmdExit_Click()
    On Error Resume Next
    Set mobj����ͨ�ö���.Form = Nothing
    Unload Me
End Sub

'���ܣ����¼������
Private Sub subClear()
    On Error Resume Next
    cifFilter.ClearContent
    cifFilter.InfoCollection.ClearHistory
    Set cgrdInput.InfoCollection = cifFilter.InfoCollection
End Sub

'���ܣ����ý����Ͽؼ���״̬��׼��¼�롣
Private Sub subBeginEdit()
    On Error GoTo errHandler

    '������ֻ��¼������ȷ����ȡ�����˳����á�
    cgrdMain.Enabled = False
    ccmdAdd.Enabled = False
    ccmdDelete.Enabled = False
    ccmdUpdate.Enabled = False
    ccmdExit.Enabled = True
    
    '¼�������á�
    Frame1.Enabled = True
    cifFilter.Enabled = True
    cidgMain.Enabled = True
    cgrdInput.Enabled = True
    ccmdOk.Enabled = True
    ccmdCancel.Enabled = True
    ccmdRemoveRow.Enabled = True
    
    Exit Sub
errHandler:
    sfsub������ "ְҵ�����ý���", "frmSetConclusionFilter", "subBeginEdit", Err.Number, Err.Description, True
End Sub

'���ܣ��ָ������Ͽؼ�״̬������¼�롣
Private Sub subReset()
    On Error Resume Next
    cgrdMain.Row = 0
    
    '������ֻ���������������˳����á�
    cgrdMain.Enabled = True
    ccmdAdd.Enabled = True
    If cgrdMain.Row > 0 Then
        ccmdDelete.Enabled = True
        ccmdUpdate.Enabled = True
    Else
        ccmdDelete.Enabled = False
        ccmdUpdate.Enabled = False
    End If
    ccmdExit.Enabled = True
        
    '¼���������á�
    Frame1.Enabled = False
    cifFilter.Enabled = False
    cidgMain.Enabled = False
    cgrdInput.Enabled = False
    ccmdOk.Enabled = False
    ccmdCancel.Enabled = False
    ccmdRemoveRow.Enabled = False
    
    mobj����������.��� = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set mobj����ͨ�ö��� = Nothing
    Set mobj���������� = Nothing
    
End Sub

'���ܣ�¼����������ϣ����롰mobj�������������������С�
Private Sub mobj����ͨ�ö���_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    Dim lstrError  As String
    Dim i As Long
    
    On Error GoTo errHandler
    '¼��һ��������ϣ���������С�
    Select Case Operate
    Case "���"
        If cifFilter.ItemsError.Count > 0 Then
            Err.Raise 6666, , "¼�뻹�д����������ɫ�����ݡ�"
        End If
        '���������������Ƿ��Ѵ��ڸ���š�
        For i = 1 To cgrdInput.Rows - 1
            If cgrdInput.TextMatrix(i, 3) = cifFilter.ItemTrueText(3) Then
                lstrError = "��Ų������ظ���" & Chr(13) & Chr(10)
                Exit For
            End If
        Next
        '�޸ģ�2001-8-24���ж�ֵ�Ƿ�Ϸ�����
        Select Case cifFilter.Box1("�ж�����").Text
        Case "<", ">", "<=", ">="
            '�ж�ֵ���������֡�
            If Not IsNumeric(cifFilter.Box1("�ж�ֵ").Text) Then
                lstrError = lstrError & "�ж�����Ϊ< (��>, <=, >=) ʱ���ж�ֵ����������ֵ�͡�"
            End If
        End Select
        If lstrError <> "" Then
            Err.Raise 6666, , "¼�����ݲ�����ϵͳ�涨���޷���ӣ�" & Chr(13) & Chr(10) & lstrError
        End If
        
        '���룺��ţ������Ŀ�������Ŀ���ƣ��ж��������ж�ֵ��
        mobj����������.subAddFilter cifFilter.ItemTrueText(3), cifFilter.ItemTrueText(4), cifFilter.ItemText(4), cifFilter.ItemText(5), cifFilter.ItemText(6)
        
    Case "�޸�"
        '�޸ģ�2001-8-24���ж�ֵ�Ƿ�Ϸ�����
        Select Case cifFilter.Box1("�ж�����").Text
        Case "<", ">", "<=", ">="
            '�ж�ֵ���������֡�
            If Not IsNumeric(cifFilter.Box1("�ж�ֵ").Text) Then
                Err.Raise 6666, , "¼�����ݲ�����ϵͳ�涨���޷���ӣ�" & Chr(13) & Chr(10) & "�ж�����Ϊ< (��>, <=, >=) ʱ���ж�ֵ����������ֵ�͡�"
            End If
        End Select
        '��ɾ���ɵġ�
        If cgrdInput.Row > 0 Then
            mobj����������.subRemoveFilter cifFilter.ItemHistory(cgrdInput.Row)(4)
        End If
        '���������������Ƿ��Ѵ��ڸ���š�
        For i = 1 To cgrdInput.Rows - 1
            If cgrdInput.TextMatrix(i, 3) = cifFilter.ItemTrueText(3) And i <> cgrdInput.Row Then
                Err.Raise 6666, , "��Ų������ظ���"
            End If
        Next
        '����ӡ�
        mobj����������.subAddFilter cifFilter.ItemTrueText(3), cifFilter.ItemTrueText(4), cifFilter.ItemText(4), cifFilter.ItemText(5), cifFilter.ItemText(6)

    Case "�˳�"
        Set mobj����ͨ�ö���.Form = Nothing
        Unload Me
    End Select
    
    Exit Sub
errHandler:
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����ý���", "frmSetConclusionFilter", "mobj����ͨ�ö���_BeforeOperate", 6666, lstrError, False
    cifFilter.ItemSetfocus 3
    Cancel = True
End Sub

Private Sub mobj����ͨ�ö���_ItemLostFocus(ByVal Index As Integer, ByVal ���� As String, ByVal ���� As String, ByVal �������� As String, ByVal IsError As Boolean)
    Dim i As Long
    Dim Row As Integer
    
    On Error GoTo errHandler
    If ���� = "" Or (ActiveControl.Name <> "cifFilter") Then Exit Sub
    
    Select Case ����
    Case "������"
        If mobj����������.ID <> �������� Then
            mobj����������.ID = ��������
        End If
    Case "����"
        mobj����������.���� = ��������
    Case "���"
    End Select

    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����ý���", "frmSetConclusionFilter", "mobj����ͨ�ö���_ItemLostFocus", 6666, lstrError, False
    cifFilter.ItemSetfocus Index
End Sub
