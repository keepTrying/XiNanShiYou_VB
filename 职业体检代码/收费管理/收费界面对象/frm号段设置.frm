VERSION 5.00
Object = "{8099FCC2-0A81-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6.ocx"
Begin VB.Form frm�Ŷ����� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��Ʊ�Ŷ�����"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton ccmdExit 
      Caption         =   "��  ��"
      Height          =   375
      Left            =   5160
      TabIndex        =   10
      Top             =   240
      Width           =   1095
   End
   Begin VB.ComboBox clstName 
      Height          =   300
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   240
      Width           =   1815
   End
   Begin VB.TextBox ctxtBegin 
      Height          =   270
      Left            =   5400
      MaxLength       =   10
      TabIndex        =   3
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox ctxtEnd 
      Height          =   270
      Left            =   5400
      MaxLength       =   10
      TabIndex        =   2
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton ccmdSave 
      Caption         =   "��  ��"
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton ccmdCancel 
      Caption         =   "ȡ  ��"
      Height          =   375
      Left            =   6120
      TabIndex        =   0
      Top             =   2640
      Width           =   1095
   End
   Begin VSFlex6Ctl.vsFlexGrid cgrdMain 
      Height          =   5895
      Left            =   360
      TabIndex        =   4
      ToolTipText     =   "˫����ѡ���ĺŶ���Ϣ���Խ����޸�"
      Top             =   1080
      Width           =   3375
      _cx             =   52172609
      _cy             =   52177054
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
      SelectionMode   =   0
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
      FormatString    =   "ID|���|ֹ��|�Ƿ�����"
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
   Begin VB.Label clblCurNo 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   5280
      TabIndex        =   12
      Top             =   4680
      Width           =   90
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "��ǰƱ�ţ�"
      Height          =   180
      Left            =   4200
      TabIndex        =   11
      Top             =   4680
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�շ�Ա��"
      Height          =   180
      Left            =   360
      TabIndex        =   9
      Top             =   240
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "�����úŶΣ�"
      Height          =   180
      Left            =   360
      TabIndex        =   8
      Top             =   720
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "��ţ�"
      Height          =   180
      Left            =   4560
      TabIndex        =   7
      Top             =   1200
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "ֹ�ţ�"
      Height          =   180
      Left            =   4560
      TabIndex        =   6
      Top             =   1800
      Width           =   540
   End
End
Attribute VB_Name = "frm�Ŷ�����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public pblnInUse As Boolean

Dim mlngID As Long      '��ǰ�޸ĵĺŶε�ID

Private Sub ccmdCancel_Click()
    mlngID = 0
    ctxtBegin = ""
    ctxtEnd = ""
End Sub

Private Sub ccmdExit_Click()
    Unload Me
End Sub

Private Sub ccmdSave_Click()
    Dim i As Integer
    Dim llngBegin As Long, llngEnd As Long
    Dim lobjRec As Object
    
    On Error GoTo errhandler
    
    ctxtBegin = Trim(ctxtBegin)
    ctxtEnd = Trim(ctxtEnd)
    If ctxtBegin = "" Then
        MsgBox "��Ų���Ϊ�գ�", vbInformation, "ϵͳ��ʾ"
        ctxtBegin.SetFocus
        Exit Sub
    End If
    If ctxtEnd = "" Then
        MsgBox "ֹ�Ų���Ϊ�գ�", vbInformation, "ϵͳ��ʾ"
        ctxtEnd.SetFocus
        Exit Sub
    End If
    llngBegin = CLng(ctxtBegin)
    llngEnd = CLng(ctxtEnd)
    If CLng(ctxtEnd) < llngBegin Then
        MsgBox "ֹ�ű���С����ţ�", vbInformation, "ϵͳ��ʾ"
        ctxtEnd.SetFocus
        Exit Sub
    End If
    '���÷�Χ�Ƿ�����ϵͳ¼��ĺŶη�Χ��
    Set lobjRec = dafuncGetData("select * from �շѹ���_�շ�Ա�Ŷ���Ϣ�� where (���<='" & llngBegin & "' and ֹ��>='" & llngBegin & "' or ���<='" & llngEnd & "' and ֹ��>='" & llngEnd & "' or ���>'" & llngBegin & "' and ֹ��<'" & llngEnd & "') and ID<>" & mlngID)
    If lobjRec.RecordCount Then
        MsgBox "����ӵĺŶη�Χ���ѷ�������շ�Ա�������շ�Ա�ĺŶη�Χ�ص���������ӣ�", vbInformation, "ϵͳ��ʾ"
        ctxtBegin.SetFocus
        Exit Sub
    End If
    If mlngID = 0 Then
        dafuncGetData "insert into �շѹ���_�շ�Ա�Ŷ���Ϣ�� (�û����,���,ֹ��,�Ƿ�����) values('" & Mid(clstName.Text, InStr(clstName.Text, " ") + 1) & "','" & llngBegin & "','" & llngEnd & "','��')"
    Else
        dafuncGetData "update �շѹ���_�շ�Ա�Ŷ���Ϣ�� set ���='" & llngBegin & "',ֹ��='" & llngEnd & "' where ID=" & mlngID
    End If
    MsgBox "����ɹ���", vbInformation, "ϵͳ��ʾ"
    ctxtBegin = ""
    ctxtEnd = ""
    mlngID = 0
    'ˢ�±��
    clstName_Click
    Exit Sub
errhandler:
    sfsub������ "�շѽ��沿��", "frmҵ������", "ccmdSave_Click", Err.Number, Err.Description, False
    Exit Sub
End Sub

Private Sub cgrdMain_DblClick()
    If cgrdMain.Row = 0 Then Exit Sub
    If cgrdMain.Cell(flexcpText, cgrdMain.Row, 3) = "��" Then
        MsgBox "�úŶ��Ѿ�ʹ����ϣ������޸ģ�", vbInformation, "ϵͳ��ʾ"
    Else
        mlngID = CLng(cgrdMain.Cell(flexcpText, cgrdMain.Row, 0))
        ctxtBegin = cgrdMain.Cell(flexcpText, cgrdMain.Row, 1)
        ctxtEnd = cgrdMain.Cell(flexcpText, cgrdMain.Row, 2)
    End If
End Sub

Private Sub clstName_Click()
    Dim lobjRec As Object, i As Integer

    On Error GoTo errhandler
    
    cgrdMain.FormatString = ""
    Set lobjRec = dafuncGetData("select ID,���,ֹ��,�Ƿ����� from �շѹ���_�շ�Ա�Ŷ���Ϣ�� where �û����='" & Mid(clstName.Text, InStr(clstName.Text, " ") + 1) & "' order by ID desc")
    Set cgrdMain.DataSource = lobjRec
    cgrdMain.ColHidden(0) = True
        
        
    Set lobjRec = dafuncGetData("select ��ǰֵ from ϵͳ����_ϵͳ������ɼ�¼�� where ҵ������='�շѹ���" & Mid(clstName.Text, InStr(clstName.Text, " ") + 1) & "' and �������='�վݺ�'")
    If lobjRec.RecordCount = 0 Then
        clblCurNo = ""
    Else
        clblCurNo = lobjRec(0)
    End If
    Exit Sub
errhandler:
    sfsub������ "�շѽ��沿��", "frmҵ������", "clstName_Click", Err.Number, Err.Description, False
    Exit Sub
    Resume
End Sub
Private Sub ctxtBegin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then ctxtEnd.SetFocus
End Sub

Private Sub ctxtBegin_LostFocus()
    ctxtBegin = Trim(ctxtBegin)
    If ctxtBegin <> "" Then
        If Not IsNumeric(ctxtBegin) Then
            MsgBox "��ű���������ȷ�����֣�", vbInformation, "ϵͳ��ʾ"
            ctxtBegin.SetFocus
        ElseIf CLng(ctxtBegin) <= 0 Then
            MsgBox "��ű����Ǵ���0��������", vbInformation, "ϵͳ��ʾ"
            ctxtBegin.SetFocus
        ElseIf ctxtEnd <> "" Then
            If CLng(ctxtEnd) < CLng(ctxtBegin) Then MsgBox "ֹ�ű���С����ţ�", vbInformation, "ϵͳ��ʾ"
        End If
    End If
End Sub

Private Sub ctxtEnd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then ccmdSave.SetFocus
End Sub

Private Sub ctxtEnd_LostFocus()
    ctxtEnd = Trim(ctxtEnd)
    If ctxtEnd <> "" Then
        If Not IsNumeric(ctxtEnd) Then
            MsgBox "ֹ�ű���������ȷ�����֣�", vbInformation, "ϵͳ��ʾ"
            ctxtEnd.SetFocus
        ElseIf CLng(ctxtEnd) <= 0 Then
            MsgBox "ֹ�ű����Ǵ���0��������", vbInformation, "ϵͳ��ʾ"
            ctxtEnd.SetFocus
        ElseIf ctxtBegin <> "" Then
            If CLng(ctxtEnd) < CLng(ctxtBegin) Then MsgBox "ֹ�ű���С����ţ�", vbInformation, "ϵͳ��ʾ"
        End If
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo errhandler
    
    If pblnInUse Then Exit Sub
    pblnInUse = True
    
    '��ȡ��ǰƱ�ݺš�
    Dim lobjRec As Object, i As Integer
    Dim lobjRec1 As Object
    
    mlngID = 0
    clstName.Clear
    Set lobjRec = dafuncGetData("select ���,���� from ϵͳ����_Ա��������Ϣ��ͼ order by ���")
    For i = 1 To lobjRec.RecordCount
        Set lobjRec1 = dafuncGetData("select * from ϵͳ����_�û�����Ȩ�ޱ� where �û����='" & lobjRec(0) & "' and Ȩ����='�շѹ���_ֱ���շ�'")
        If lobjRec1.RecordCount > 0 Then
            clstName.AddItem lobjRec(1) & " " & lobjRec(0)
        End If
        lobjRec.MoveNext
    Next
    If clstName.ListCount > 0 Then
        clstName.ListIndex = 0
    Else
        MsgBox "��ǰû�����þ����շ�Ȩ�޵���Ա����������Ʊ�ݺŶΣ�����Ϊ�շ�Ա�����շ�Ȩ�ޣ�", vbInformation, "ϵͳ��ʾ"
    End If
    
    Exit Sub
errhandler:
    sfsub������ "�շѽ��沿��", "frmҵ������", "Form_Load", Err.Number, Err.Description, False
    Exit Sub
    Resume
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    pblnInUse = False
End Sub
