VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmQueryCompanyLocation 
   Caption         =   "��λ��λ��ѯ"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   6465
   StartUpPosition =   3  '����ȱʡ
   Begin VSFlex8Ctl.VSFlexGrid clqlist 
      Height          =   3495
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   6135
      _cx             =   10821
      _cy             =   6165
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
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
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4800
      TabIndex        =   0
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "˵��������ֱ��˫��ѡ�е�λ"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   4080
      Width           =   3015
   End
End
Attribute VB_Name = "frmQueryCompanyLocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'���ܣ���ѯͳ�Ƶĵ�λ�����ֶ����룬ֱ��ѡ��λ by moujun  2015-11-2

Option Explicit
Private chickname As String


'������ѡ��ȷ������������ı���
Private Sub clqlist_Click()
'    Dim indX As String
'    Dim indY As String
'    indX = clqlist.MouseRow
'    indY = clqlist.MouseCol
'    If indX < 0 Or indY < 0 Then
'        Exit Sub
'    ElseIf indX >= 0 And indX < clqlist.rows And indY >= 0 And indY < clqlist.cols Then
    chickname = clqlist.TextMatrix(clqlist.Row, 1)
    clqlist.SelectionMode = flexSelectionByRow     'ѡ�����ɫ����
'    End If
    Command1.Visible = True
    Command1.Enabled = True
    
End Sub


'˫����ֱ��ѡ�������ı���
Private Sub clqlist_DblClick()
   chickname = clqlist.TextMatrix(clqlist.Row, 1)
   clqlist.SelectionMode = flexSelectionByRow
    FrmQueryCompany.ctxtCompanyName.Text = chickname     '��λͳ�Ƶ�λ����
    FrmQueryStatis.ctxt��λ����.Text = chickname        '��ѯͳ�Ƶ�λ����
    frmQueryCompanyLocation.Hide
End Sub

Private Sub Command1_Click()
    FrmQueryCompany.ctxtCompanyName.Text = chickname  '��λͳ�Ƶ�λ����
    FrmQueryStatis.ctxt��λ����.Text = chickname  '��ѯͳ�Ƶ�λ����
    frmQueryCompanyLocation.Hide
End Sub

'�������
Private Sub Form_Load()
    Dim i As Integer
    Dim cname As String
    Dim lobjRec As Object
    Dim sql As String
'    Dim sql�����ʱ�� As String
'    Dim sql��ѯ��� As String
    sql = "select ��λ���� from ְҵ�����_���������ݿ� group by ��λ����"
    Set lobjRec = dafuncGetData(sql)
    clqlist.Clear
    For i = 1 To lobjRec.RecordCount
        cname = lobjRec("��λ����")
        lobjRec.MoveNext
    Next
'    cgrdList.Clear
    Set clqlist.DataSource = lobjRec
'    Set mobjRec = lobjRec
    clqlist.AutoSize 0, clqlist.cols - 1, 0, 0
    clqlist.ExplorerBar = flexExSort
    clqlist.DataMode = flexDMFree
    clqlist.Col = 0
    Exit Sub
'    Command1.Enabled = False
'    Command1.Visible = False
End Sub
