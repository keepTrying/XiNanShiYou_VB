VERSION 5.00
Object = "{8099FCC2-0A81-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm��ӡ���ж��ʵ� 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "��ӡ���ж��ʵ�"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   10905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.OptionButton coptType 
      Caption         =   "�����շ�"
      Height          =   375
      Index           =   1
      Left            =   6240
      TabIndex        =   18
      Top             =   720
      Width           =   1095
   End
   Begin VB.OptionButton coptType 
      Caption         =   "һ���Խɷ�"
      Height          =   375
      Index           =   0
      Left            =   4680
      TabIndex        =   17
      Top             =   720
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.CheckBox cchkAll 
      Caption         =   "ȫѡ"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   1080
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CommandButton ccmdExit 
      Caption         =   "�ر�(&C)"
      Height          =   375
      Left            =   9720
      TabIndex        =   15
      Top             =   720
      Width           =   1000
   End
   Begin VB.CommandButton ccmdPrint 
      Caption         =   "��ӡ(&P)"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8640
      TabIndex        =   11
      Top             =   720
      Width           =   1000
   End
   Begin VB.CommandButton ccmdQuery 
      Caption         =   "��ѯ(&Q)"
      Height          =   375
      Left            =   7560
      TabIndex        =   10
      Top             =   720
      Width           =   1000
   End
   Begin VB.ComboBox ccmb������ 
      Height          =   300
      Left            =   960
      TabIndex        =   9
      Top             =   720
      Width           =   3375
   End
   Begin VB.TextBox ctxtƱ�ݺ� 
      Height          =   375
      Index           =   1
      Left            =   7680
      TabIndex        =   7
      Top             =   240
      Width           =   1935
   End
   Begin VB.TextBox ctxtƱ�ݺ� 
      Height          =   375
      Index           =   0
      Left            =   5400
      TabIndex        =   5
      Top             =   240
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker cdtp��ֹ���� 
      Height          =   300
      Left            =   2880
      TabIndex        =   1
      Top             =   240
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   529
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   25296897
      CurrentDate     =   36951
   End
   Begin MSComCtl2.DTPicker cdtp��ʼ���� 
      Height          =   300
      Left            =   720
      TabIndex        =   2
      Top             =   240
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   529
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   25296897
      CurrentDate     =   36951
   End
   Begin VSFlex6Ctl.vsFlexGrid cgrdDetail 
      Height          =   6420
      Left            =   120
      TabIndex        =   12
      Top             =   1440
      Width           =   10605
      _cx             =   163793170
      _cy             =   163785788
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
      BackColorAlternate=   16437167
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
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
      TabBehavior     =   1
      OwnerDraw       =   0
      Editable        =   -1  'True
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
   Begin VB.Label clblTotal 
      AutoSize        =   -1  'True
      Caption         =   "2000"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1680
      TabIndex        =   14
      Top             =   7920
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "�ܽ��(Ԫ)��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   13
      Top             =   7920
      Width           =   1545
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "��������"
      Height          =   180
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   720
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��"
      Height          =   180
      Index           =   1
      Left            =   7440
      TabIndex        =   6
      Top             =   240
      Width           =   180
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Ʊ�ݺţ�"
      Height          =   180
      Left            =   4680
      TabIndex        =   4
      Top             =   240
      Width           =   720
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��"
      Height          =   180
      Index           =   0
      Left            =   2520
      TabIndex        =   3
      Top             =   240
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "���ڣ�"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   540
   End
End
Attribute VB_Name = "frm��ӡ���ж��ʵ�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrID As String
Private Sub cchkAll_Click()
    Dim i As Long

    For i = 1 To cgrdDetail.Rows - 1
        cgrdDetail.Cell(flexcpChecked, i, 2) = IIf(cchkAll.Value = 1, flexChecked, flexUnchecked)
    Next
    
    sub��ʾ�ܶ�
End Sub

Private Sub sub��ʾ�ܶ�()
    Dim i As Long
    Dim ldblTotal As Double
    Dim lIndex As Long
    
    For i = 0 To cgrdDetail.Cols - 1
        If cgrdDetail.TextMatrix(0, i) = "���" Then
            lIndex = i
            Exit For
        End If
    Next
    For i = 1 To cgrdDetail.Rows - 1
        If cgrdDetail.Cell(flexcpChecked, i, 2) = flexChecked Then
            ldblTotal = Format(ldblTotal + cgrdDetail.ValueMatrix(i, lIndex), "0.00")
        End If
    Next
    clblTotal = ldblTotal
End Sub
Private Sub ccmdExit_Click()
    Unload Me
End Sub

Private Sub ccmdPrint_Click()
    Dim lobjRec As Object
    Dim lCount As Long
    Dim lcolItem As Collection
    Dim lcolInfo As Collection
    Dim lobj��ӡ As Object
    Dim i As Long
    On Error GoTo errhandler
    
    '�����ݵ�����ʱ��
    If mstrID <> "" Then dafuncGetData "delete temp_���ж��ʵ� where ID='" & mstrID & "'"
    
    Set lobjRec = dafuncGetData("select newid()")
    mstrID = lobjRec(0)
    lCount = 0
    For i = 1 To cgrdDetail.Rows - 1
        If cgrdDetail.Cell(flexcpChecked, i, 2) = flexChecked Then
            dafuncGetData "insert into temp_���ж��ʵ�(ID,�շ�����,�շ���Ŀ���) values('" & mstrID & "','" & cgrdDetail.TextMatrix(i, 0) & "','" & cgrdDetail.TextMatrix(i, 1) & "')"
            lCount = lCount + 1
        End If
    Next
    If lCount = 0 Then
        sffuncMsg "����Ҫ��ӡ�ļ�¼ǰ�򹴣�"
    Else
        Dim lstrFilter As String
        lstrFilter = IIf(IsNull(cdtp��ʼ����.Value), "", "���ڣ�" & cdtp��ʼ����.Value) & IIf(IsNull(cdtp��ֹ����.Value), "", " ����" & cdtp��ֹ����.Value) & IIf(ctxtƱ�ݺ�(0) <> "" Or ctxtƱ�ݺ�(1) <> "", "     Ʊ�ݺţ�" & ctxtƱ�ݺ�(0) & " �� " & ctxtƱ�ݺ�(1), "") & IIf(ccmb������ = "", "", " �������У�" & ccmb������) & IIf(coptType(0).Value, "   һ���Խɷ�", "   �����շ�")
        
        Set lcolInfo = New Collection
        Set lcolItem = New Collection
        lcolItem.Add "ID", "����"
        lcolItem.Add mstrID, "ֵ"
        lcolInfo.Add lcolItem
        
        Set lcolItem = New Collection
        lcolItem.Add "����", "����"
        lcolItem.Add lstrFilter, "ֵ"
        lcolInfo.Add lcolItem
        
        Set lobj��ӡ = CreateObject("ͨ��ˮ�������ӡ.cls����")
        lobj��ӡ.funcPrintReport "���ж��ʵ�", lcolInfo, App.Path, True
        
    End If
    Exit Sub
errhandler:
    sfsub������ "�շѽ��沿��", "frm��ӡ���ж��ʵ�", "ccmdPrint_Click", Err.Number, Err.Description, False
    
End Sub

Private Sub ccmdQuery_Click()
    Dim lobjRec As Object
    Dim i As Long
    On Error GoTo errhandler
    Set lobjRec = dafuncGetData("exec �շѹ���_��ѯ���ж��ʵ� '" & IIf(IsNull(cdtp��ʼ����.Value), "", cdtp��ʼ����.Value) & "','" & IIf(IsNull(cdtp��ֹ����.Value), "", cdtp��ֹ����.Value) & "','" & ctxtƱ�ݺ�(0) & "','" & ctxtƱ�ݺ�(1) & "','" & ccmb������ & "','" & IIf(coptType(0).Value, "һ��", "����") & "'")
    cgrdDetail.FormatString = ""
    Set cgrdDetail.DataSource = lobjRec
    
    cgrdDetail.Editable = True

    cgrdDetail.ColHidden(0) = True
    cgrdDetail.ColHidden(1) = True
    
    For i = 1 To cgrdDetail.Rows - 1
        cgrdDetail.Cell(flexcpChecked, i, 2) = IIf(cchkAll.Value = 1, flexChecked, flexUnchecked)
    Next
    cgrdDetail.ColWidth(0) = 1200
    sub��ʾ�ܶ�
    If cgrdDetail.Rows > 0 Then
        
        ccmdPrint.Enabled = True
    Else
        ccmdPrint.Enabled = False
    End If
    Exit Sub
errhandler:
    sfsub������ "�շѽ��沿��", "frm��ӡ���ж��ʵ�", "ccmdQuery_Click", Err.Number, Err.Description, False
    
End Sub

Private Sub cgrdDetail_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    sub��ʾ�ܶ�
End Sub

Private Sub cgrdDetail_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 2 Then Cancel = True
    
End Sub

Private Sub coptType_Click(Index As Integer)
    If coptType(0).Value Then
        ccmb������.Enabled = True
    Else
        ccmb������.Enabled = False
        ccmb������.ListIndex = -1
    End If
    
End Sub

Private Sub Form_Load()
    Dim lobjRec As Object
    On Error GoTo errhandler
    '��ȡ�������ʺ�.
    Set lobjRec = dafuncGetData("select ������+' '+�ʺ� from �շѹ���_���п��������ñ�")
    ccmb������.Clear
    ccmb������.AddItem ""
    Do While Not lobjRec.EOF
        ccmb������.AddItem lobjRec(0)
        
        lobjRec.MoveNext
    Loop
    
    cdtp��ʼ����.Value = Format(Now, "yyyy/mm/dd")
    cdtp��ֹ����.Value = Format(Now, "yyyy/mm/dd")
    cgrdDetail.Editable = True
    Exit Sub
errhandler:
    sfsub������ "�շѽ��沿��", "frm��ӡ���ж��ʵ�", "Form_Load", Err.Number, Err.Description, False
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    If mstrID <> "" Then dafuncGetData "delete temp_���ж��ʵ� where ID='" & mstrID & "'"
End Sub
