VERSION 5.00
Object = "{8099FCC2-0A81-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frm������ 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "������"
   ClientHeight    =   7620
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   11070
   ClipControls    =   0   'False
   Icon            =   "frm������.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   11070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton ccmdQuery 
      Caption         =   "��  ѯ"
      Height          =   375
      Left            =   8520
      TabIndex        =   21
      Top             =   720
      Width           =   1095
   End
   Begin VB.ComboBox clstName 
      Height          =   300
      Left            =   6360
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   720
      Width           =   1815
   End
   Begin MSComctlLib.ImageList cimg��ťͼ�� 
      Left            =   7080
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   6960
      Width           =   10995
      Begin VB.TextBox ctxt�ϼ� 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """��""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   375
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   120
         Width           =   2415
      End
      Begin VB.Timer ctmr��ʱ 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   5640
         Top             =   120
      End
      Begin VB.TextBox ctxt�˷����� 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   120
         Width           =   1905
      End
      Begin VB.TextBox ctxt�ܽ�� 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """��""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   2
         EndProperty
         Height          =   330
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ܽ��"
         Height          =   180
         Index           =   1
         Left            =   7680
         TabIndex        =   10
         Top             =   240
         Width           =   540
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "�շ�����"
         Height          =   240
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���"
         Height          =   180
         Index           =   0
         Left            =   3360
         TabIndex        =   5
         Top             =   240
         Width           =   360
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      Caption         =   "������ϸ��Ϣ"
      ForeColor       =   &H80000008&
      Height          =   6195
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   10980
      Begin VB.OptionButton coptType 
         Caption         =   "����Ʊ��"
         Height          =   255
         Index           =   1
         Left            =   3960
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton coptType 
         Caption         =   "�շѼ�¼"
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VSFlex6Ctl.vsFlexGrid cgrdMain 
         Height          =   5565
         Left            =   75
         TabIndex        =   1
         Top             =   540
         Width           =   7935
         _cx             =   25376428
         _cy             =   25372248
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
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
         BackColorAlternate=   16777215
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   3
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   -1  'True
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
      Begin VSFlex6Ctl.vsFlexGrid cgrdDetail 
         Height          =   5565
         Left            =   8040
         TabIndex        =   13
         Top             =   540
         Width           =   2895
         _cx             =   162862066
         _cy             =   162866776
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
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
         BackColorAlternate=   16777215
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   3
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   -1  'True
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   1
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
      Begin VB.Label clblMark 
         BackColor       =   &H00C0FFFF&
         Height          =   195
         Left            =   5640
         TabIndex        =   12
         Top             =   240
         Width           =   195
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   5880
         TabIndex        =   11
         Top             =   240
         Width           =   360
      End
   End
   Begin MSComctlLib.Toolbar ctlb������ 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11070
      _ExtentX        =   19526
      _ExtentY        =   1005
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "cimg��ťͼ��"
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker cdtp��ֹ���� 
      Height          =   300
      Left            =   3480
      TabIndex        =   15
      Top             =   720
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   529
      _Version        =   393216
      Format          =   51773441
      CurrentDate     =   36951
   End
   Begin MSComCtl2.DTPicker cdtp��ʼ���� 
      Height          =   300
      Left            =   1320
      TabIndex        =   16
      Top             =   720
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   529
      _Version        =   393216
      Format          =   51773441
      CurrentDate     =   36951
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "�շ�Ա��"
      Height          =   180
      Left            =   5520
      TabIndex        =   20
      Top             =   720
      Width           =   720
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���ڷ�Χ"
      Height          =   180
      Left            =   360
      TabIndex        =   18
      Top             =   720
      Width           =   720
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��"
      Height          =   180
      Left            =   3000
      TabIndex        =   17
      Top             =   720
      Width           =   180
   End
   Begin VB.Menu cmnuView 
      Caption         =   "ϵͳ(&S)"
      Begin VB.Menu cmnuItemView 
         Caption         =   "�˳�(&X)"
         Index           =   5
      End
   End
End
Attribute VB_Name = "frm������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public pblnInUse As Boolean

Private WithEvents mobjGUI As cls����ͨ�ö���
Attribute mobjGUI.VB_VarHelpID = -1

Private mstr��λ��� As String

'��ѯ����
Private mstr�վݺ� As String

Private mobjQueryResult As Object
Private mcolIndex As Collection

Private mstrSQL As String  '�����ַ���
  

Private Sub ccmdQuery_Click()
    sub��ѯ����ʾ��¼
End Sub

'���ܣ����������Ϣ���е�һ�У�ˢ����ʾ��ѡ��ļ�¼
'���룺��
'�������
'���أ���
'ע�������
'���ߣ�����
'����ʱ�䣺2001-3-29
Private Sub cgrdMain_Click()
    On Error GoTo errHandler
    sub��ʾһ��������ϸ
    If coptType(1).Value Then
        ctlb������.Buttons(5).Visible = True
    Else
        ctlb������.Buttons(5).Visible = False
    End If
    Exit Sub
errHandler:
    Call sfsub������("�շѽ��沿��", "frm�շѹ���", "cgrdMain_Click", Err.Number, Err.Description, False)
End Sub

'���ܣ��ڷ�����Ϣ�������β��ְ���
'���룺��
'�������
'���أ���
'ע�������
'���ߣ�����
'����ʱ�䣺2001-3-29
Private Sub cgrdMain_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyHome Or KeyCode = vbKeyEnd Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
        KeyCode = 0
    End If
End Sub


Private Sub cgrdMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
End Sub


Private Sub cmnuItemView_Click(Index As Integer)
    On Error GoTo errHandler
    Select Case Index
    Case 1 '��ѯ
        sub��ѯ����ʾ��¼
    Case 2 'ˢ��
        sub��ʾ��¼
    Case 5
        Unload Me
    End Select

    Exit Sub
errHandler:
    Call sfsub������("�շѽ��沿��", "frm�շѹ���", "cmnuItemView_Click", Err.Number, Err.Description, False)
End Sub

Private Sub coptType_Click(Index As Integer)
    On Error GoTo errHandler
    
    sub��ʾ��¼
    
    Exit Sub
errHandler:
    Call sfsub������("�շѽ��沿��", "frm�շѹ���", "coptType_Click()", Err.Number, Err.Description, False)
    
End Sub

Private Sub Form_Load()
    If pblnInUse Then Exit Sub
    Dim lcol��������ť As Collection
    Dim lLen As Integer
    On Error GoTo errHandler
    pblnInUse = True                              'ָʾ����������

    '��ʼ��������
    Set mobjGUI = New cls����ͨ�ö���
    Set mobjGUI.Form = Me
    Set mobjGUI.c������ = ctlb������
    Set lcol��������ť = New Collection
    lcol��������ť.Add "�Ŷ�(&Q)105"
    lcol��������ť.Add "|"
    lcol��������ť.Add "����(&T)110"
    lcol��������ť.Add "��������(&J)111"
    lcol��������ť.Add "ȡ������(&H)109"
    lcol��������ť.Add "|"
    lcol��������ť.Add "�˳�"
    mobjGUI.subInitialize lcol��������ť, ""
    
    Dim lobjRec As Object, i As Integer
    Dim lobjRec1 As Object
    
    clstName.Clear
    Set lobjRec = dafuncGetData("select ���,���� from ϵͳ����_Ա��������Ϣ��ͼ order by ���")
    For i = 1 To lobjRec.RecordCount
        Set lobjRec1 = dafuncGetData("select * from ϵͳ����_�û�����Ȩ�ޱ� where �û����='" & lobjRec(0) & "' and Ȩ����='�շѹ���_ֱ���շ�'")
        If lobjRec1.RecordCount > 0 Then
            clstName.AddItem lobjRec(0) & " " & lobjRec(1)
        End If
        lobjRec.MoveNext
    Next
    If clstName.ListCount > 0 Then
        clstName.ListIndex = 0
    Else
        MsgBox "��ǰû�����þ����շ�Ȩ�޵���Ա��", vbInformation, "ϵͳ��ʾ"
    End If
    
    'Ĭ����ʾ����������շѼ�¼��
    cdtp��ʼ����.Value = Format(Date, "yyyy-mm-dd")
    cdtp��ֹ����.Value = Format(Date, "yyyy-mm-dd")
    
    sub��ѯ����ʾ��¼
    
    coptType_Click 0
    Exit Sub
errHandler:
    Call sfsub������("�շѽ��沿��", "frm�շѹ���", "Form_Load", Err.Number, Err.Description, False)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Frame2.Width = Me.ScaleWidth - Frame2.Left - 60
    Frame4.Width = Me.ScaleWidth - Frame4.Left - 60
    Frame4.Top = Me.ScaleHeight - Frame4.Height - 120
    
    Frame2.Height = Frame4.Top - Frame2.Top - 60
    cgrdMain.Width = Frame2.Width * 0.7
    cgrdMain.Height = Frame2.Height - cgrdMain.Top - 60
    
    cgrdDetail.Left = cgrdMain.Left + cgrdMain.Width + 60
    cgrdDetail.Width = Frame2.Width - cgrdDetail.Left - 60
    cgrdDetail.Height = cgrdMain.Height
    
End Sub

'���ܣ��رմ��壬����pblninuseΪFalse
'���룺��
'�������
'���أ���
'ע�������
'���ߣ�����
'����ʱ�䣺2001-3-29
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    pblnInUse = False
    Set mobjGUI = Nothing
    Set mobjQueryResult = Nothing
End Sub


Private Sub sub��ѯ����ʾ��¼()
    Dim i As Integer
    Dim lint�˷Ѽ�¼�� As Long
    Dim lInt As Long            '����ѭ������
    
    
    On Error GoTo errhandle
    cgrdDetail.Rows = 1
    
    '��ѯ�շѼ�¼.
    Set mobjQueryResult = pobj�շѹ���.func�����ܽ����ѯ(mstr�վݺ�, Left(clstName.Text, InStr(clstName.Text, " ") - 1), cdtp��ʼ����.Value, cdtp��ֹ����.Value)
    
    sub��ʾ��¼
    
    Exit Sub
errhandle:
    Call sfsub������("�շѽ��沿��", "frm�շѹ���", "sub��ѯ����ʾ��¼()", Err.Number, Err.Description, True)
End Sub

Private Sub sub��ʾ��¼()
    Dim i As Long
    On Error GoTo errhandle
    
    cgrdDetail.Rows = 1
    
    If coptType(0).Value Then
        mobjQueryResult.Filter = "��ʶ=1"
    ElseIf coptType(1).Value Then
        mobjQueryResult.Filter = "��ʶ=3"
    End If
'    mobjQueryResult.Sort = "�շ�����,�շѱ��"
    mobjQueryResult.Sort = "�շѱ��"
    Set cgrdMain.DataSource = mobjQueryResult
    
    Set mcolIndex = New Collection
    For i = 0 To cgrdMain.Cols - 1
        mcolIndex.Add i, cgrdMain.TextMatrix(0, i)
    Next
    
    cgrdMain.ColHidden(mcolIndex("�շ�״̬")) = True
    cgrdMain.ColHidden(mcolIndex("��ʶ")) = True
    
    
'    Call subѡ��ͬ������
    If cgrdMain.Row > 0 Then
        sub��ʾһ��������ϸ
    End If
    '��ʾ�ϼơ�
    Dim dblTotal As Double
    For i = 1 To cgrdMain.Rows - 1
        dblTotal = Format(dblTotal + cgrdMain.ValueMatrix(i, mcolIndex("���")), "0.00")
        
        '��ʾ���ϼ�¼����ɫ��
        If cgrdMain.TextMatrix(i, mcolIndex("�շ�״̬")) = 3 Then
            cgrdMain.Cell(flexcpBackColor, i, 0, i, cgrdMain.Cols - 1) = clblMark.BackColor
        End If
            
    Next
    ctxt�ϼ� = Format(dblTotal, "0.00")
    cgrdMain.AutoSize 0, cgrdMain.Cols - 1
    Exit Sub
errhandle:
    Call sfsub������("�շѽ��沿��", "frm�շѹ���", "sub��ʾ��¼()", Err.Number, Err.Description, True)
    
End Sub


Private Sub mobjGUI_Operate(ByVal Operate As String, Cancel As Boolean)
    On Error GoTo errhandle
    Select Case Operate
        Case "�Ŷ�"
            frm�Ŷ�����.clstName.ListIndex = clstName.ListIndex
            frm�Ŷ�����.ccmdSave.Enabled = False
            frm�Ŷ�����.Show 1
        Case "����"
            frm����.clblName.Visible = True
            frm����.clstName.Visible = True
            frm����.Show 1
        Case "��������"
            frm��������.Show 1
        Case "ȡ������"
            If cgrdMain.Row = 0 Then
                MsgBox "��ѡ��Ҫȡ����Ʊ����Ϣ��", vbInformation, "ϵͳ��ʾ"
                Exit Sub
            End If
            If MsgBox("��ȷ��Ҫ��ѡ�еķ�Ʊ�ָ�Ϊ������", vbQuestion + vbYesNo + vbDefaultButton2, "ϵͳ��ʾ") = vbYes Then
                pobj�շѹ���.subȡ�����Ϸ�����Ϣ cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("�շѱ��"))
                'ˢ�½��档
                sub��ѯ����ʾ��¼
            End If
    End Select
    Exit Sub
errhandle:
    sffuncMsg Operate & "���ɹ���" & Err.Description, sf����
End Sub

'���ܣ����Զ�δ�ӡ�˷���Ϣ
'ʱ�䣺2002/08/02
'���ߣ��켽��
Private Sub subѡ��ͬ������()
Dim i As Integer
Dim lcur��� As Currency
Dim lbln As Boolean
Dim lInt As Long
    ctxt�ܽ��.Text = "0.00 Ԫ"
    ctxt�˷�����.Text = ""
    
    lInt = cgrdMain.RowSel
    If lInt > 0 Then
         If cgrdMain.TextMatrix(lInt, mcolIndex("����")) >= 0 Then
            lbln = True
         Else
            lbln = False
         End If
    End If
    With cgrdMain
    If .Rows > 1 Then
        For i = 1 To .Rows - 1
            .IsSelected(i) = False
        Next i
        
        For i = 1 To .Rows - 1
        
            If (UCase(.TextMatrix(i, 0)) = UCase(.TextMatrix(.Row, 0))) Then
                .IsSelected(i) = True
            End If
        Next i
        
        For i = 1 To .Rows - 1
            If .IsSelected(i) = True Then
                .TopRow = i
                Exit For
            End If
        Next i
                
        lcur��� = 0
        For i = 1 To .Rows - 1
            If .IsSelected(i) = True Then
               lcur��� = lcur��� + CCur(.TextMatrix(i, mcolIndex("����"))) * Val(.TextMatrix(i, mcolIndex("����")))
            End If
        Next i
        ctxt�ܽ��.Text = CStr(lcur���) & " Ԫ"
        ctxt�˷�����.Text = .TextMatrix(.Row, 0)
    Else
        ctxt�ܽ��.Text = "0.00 Ԫ"
        ctxt�˷�����.Text = ""
    End If
    End With
    
    
   
    With cgrdMain
    If .Rows > 1 Then
        For i = 1 To .Rows - 1
            .IsSelected(i) = False
        Next i
        
        For i = 1 To .Rows - 1
            If lbln = True Then
                If (UCase(.TextMatrix(i, 0)) = UCase(.TextMatrix(.Row, 0))) And .TextMatrix(i, mcolIndex("����")) >= 0 Then
                    .IsSelected(i) = True
                End If
            Else
                If (UCase(.TextMatrix(i, 0)) = UCase(.TextMatrix(.Row, 0))) And .TextMatrix(i, mcolIndex("����")) < 0 Then
                    .IsSelected(i) = True
                End If
            End If
        Next i
        
        For i = 1 To .Rows - 1
            If .IsSelected(i) = True Then
                .TopRow = i
                Exit For
            End If
        Next i
                
        lcur��� = 0
        For i = 1 To .Rows - 1
            If .IsSelected(i) = True Then
               lcur��� = lcur��� + CCur(.TextMatrix(i, mcolIndex("����"))) * Val(.TextMatrix(i, mcolIndex("����")))
            End If
        Next i
        ctxt�ܽ��.Text = CStr(lcur���) & " Ԫ"
        ctxt�˷�����.Text = .TextMatrix(.Row, 0)
    Else
        ctxt�ܽ��.Text = "0.00 Ԫ"
        ctxt�˷�����.Text = ""
    End If
    End With
End Sub

Private Sub sub��ʾһ��������ϸ()
    Dim lstrNo As String
    Dim lobjRec As Object
    On Error GoTo errHandler
    If cgrdMain.Row < 1 Then Exit Sub
    
    lstrNo = cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("�շѱ��"))
    Set lobjRec = pobj�շѹ���.func��ѯ������ϸ(lstrNo)
    cgrdDetail.FormatString = ""
    Set cgrdDetail.DataSource = lobjRec
    cgrdDetail.AutoResize = True
    cgrdDetail.MergeCol(0) = True
    cgrdDetail.MergeCells = flexMergeRestrictColumns
    ctxt�˷�����.Text = cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("�շѱ��"))
    ctxt�ܽ��.Text = cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("���"))
    Exit Sub
errHandler:
    Call sfsub������("�շѽ��沿��", "frm�շѹ���", "sub��ʾһ��������ϸ()", Err.Number, Err.Description, True)
End Sub

