VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Begin VB.Form FrmPrintBarCodeAgain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "���´�ӡ��������"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   5280
   StartUpPosition =   1  '����������
   Begin VSFlex8Ctl.VSFlexGrid cgrdSysNum 
      Height          =   2895
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   1815
      _cx             =   2088766593
      _cy             =   2088768498
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
      Cols            =   1
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   "�ش���������  "
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
   Begin VB.CommandButton ccmdExit 
      Caption         =   "��  ��"
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton ccmdPrint 
      Caption         =   "��  ӡ"
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   2400
      TabIndex        =   2
      Top             =   840
      Width           =   2415
   End
End
Attribute VB_Name = "FrmPrintBarCodeAgain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'2012-02-27 �ڵ�� ���´�ӡ����
'�������������԰�Ctrl��ѡ�ж�����´�ӡ�������Ա��Ȼ������������ϴ�ӡ

Private Sub ccmdExit_Click()
    Unload Me
    Set FrmPrintBarCodeAgain = Nothing
End Sub

Private Sub ccmdPrint_Click()
    Dim i As Integer
    On Error GoTo errHandler
    
    'For i = 1 To cgrdSysNum.Rows - 1
        'sub��ӡ������������ (cgrdSysNum.TextMatrix(i, 0))
    'Next i
    
'2012-04-05 ��¶
'��ӡ�����������
    'Dim para�������� As New Collection
    'For i = 1 To cgrdSysNum.Rows - 1
        'para�������� = para�������� & (cgrdSysNum.TextMatrix(i, 0)) & ","
    'Next i
    'sub��ӡ����������� (para��������)
    
    Dim para�������� As Collection
    Set para�������� = New Collection
    For i = 1 To cgrdSysNum.Rows - 1
        para��������.Add (cgrdSysNum.TextMatrix(i, 0))
    Next i
    sub��ӡ����������� para��������
    
    pstr�Ƿ��˻�ϵͳ��� = False
'2012-04-05 ��¶
    Exit Sub
    
errHandler:
    sfsub������ "ְҵ������", "FrmPrintBarCodeAgain", "ccmdPrint_Click", Err.Number, Err.Description, False
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    
    sub��ʾ�ش������
    If cgrdSysNum.Rows = 1 Then ccmdPrint.Enabled = False
    Label1.Caption = "���´�ӡ�������:" & (cgrdSysNum.Rows - 1)
    Exit Sub
    
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����沿��", "FrmPrintBarCodeAgain", "Form_Load", 6666, lstrError, False
End Sub

Sub sub��ʾ�ش������()
    Dim i, indextmp As Integer
    On Error GoTo errHandler
    
    With frmRegisterManage.cgrdMain
        For i = 0 To .SelectedRows - 1
            cgrdSysNum.AddItem .Cell(flexcpText, .SelectedRow(i), 0)
        Next i
    End With
    Exit Sub
    
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ������", "frmPrintBarCodeAgain", "sub��ʾ�ش������", 6666, lstrError, True
End Sub

