VERSION 5.00
Object = "{8099FCC2-0A81-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm������� 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "�������"
   ClientHeight    =   7860
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   11070
   Icon            =   "frm�������.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   11070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CheckBox cchkPreview 
      Caption         =   "��ӡǰԤ��"
      Height          =   375
      Left            =   8520
      TabIndex        =   14
      Top             =   120
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.CheckBox cchkType 
      BackColor       =   &H00C0FFFF&
      Caption         =   "�Ѵ�ӡ"
      Height          =   255
      Index           =   1
      Left            =   8040
      TabIndex        =   13
      Top             =   960
      Width           =   1335
   End
   Begin VB.CheckBox cchkType 
      BackColor       =   &H80000009&
      Caption         =   "δ��ӡ"
      Height          =   255
      Index           =   0
      Left            =   6840
      TabIndex        =   12
      Top             =   960
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.TextBox ctxt��ע 
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   7320
      Width           =   8415
   End
   Begin VB.TextBox ctxt�������� 
      Height          =   375
      Left            =   7560
      TabIndex        =   3
      Top             =   6840
      Width           =   1935
   End
   Begin VB.TextBox ctxt�������� 
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   6840
      Width           =   1815
   End
   Begin VB.TextBox ctxt������ 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   6840
      Width           =   2055
   End
   Begin VB.ListBox clstUnit 
      Height          =   5280
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   3015
   End
   Begin VSFlex6Ctl.vsFlexGrid cgrdMain 
      Height          =   5415
      Left            =   3240
      TabIndex        =   6
      Top             =   1320
      Width           =   7725
      _cx             =   25310522
      _cy             =   25306447
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
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   "���   |����    |�Ա�    |����    |��λ����     |��������    |��ҵ���    |ְҵ    |�������   | ������ "
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
   Begin MSComctlLib.Toolbar C������ 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   11070
      _ExtentX        =   19526
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "cimg��ťͼ��"
      _Version        =   393216
      Begin MSComctlLib.ImageList cimg��ťͼ�� 
         Left            =   1200
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   15
         ImageHeight     =   15
         MaskColor       =   12632256
         _Version        =   393216
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ע��"
      Height          =   180
      Index           =   3
      Left            =   120
      TabIndex        =   11
      Top             =   7440
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�������ڣ�"
      Height          =   180
      Index           =   2
      Left            =   6600
      TabIndex        =   10
      Top             =   6960
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����ʱ��(��)��"
      Height          =   180
      Index           =   1
      Left            =   3360
      TabIndex        =   9
      Top             =   6960
      Width           =   1260
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����ţ�"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   6960
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����������"
      Height          =   180
      Index           =   1
      Left            =   3360
      TabIndex        =   7
      Top             =   1080
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��λ���ƣ�"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   900
   End
End
Attribute VB_Name = "frm�������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mobjGUI As cls����ͨ�ö��� '���������õĽ���ͨ�ö�
Attribute mobjGUI.VB_VarHelpID = -1



'���ܣ����Ʋ������뵥ӡ�ţ�����س���
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        SendKeys Chr(9)
    ElseIf KeyCode = 39 Then
        KeyCode = 0
    End If
    

End Sub
Private Sub cchkType_Click(Index As Integer)
    On Error GoTo errhandler
    subRefresh
    
    Exit Sub
errhandler:
    sfsub������ "����֤������", "frm�������", "cchkType_Click", Err.Number, Err.Description, False
    
End Sub

Private Sub clstUnit_Click()
    Dim i As Long
    Dim lobjRec As Object
    On Error GoTo errhandler
    
    '��ʾѡ�е�λ�ĵ�����Ա��
    Set lobjRec = pobj������.func��ȡ������Ա(clstUnit.List(clstUnit.ListIndex))
    cgrdMain.FormatString = ""
    Set cgrdMain.DataSource = lobjRec
    cgrdMain.ColHidden(0) = True
    If cgrdMain.Rows > 1 Then
        ctxt������.Text = cgrdMain.TextMatrix(1, 1)
        ctxt��������.Text = cgrdMain.TextMatrix(1, cgrdMain.Cols - 4)
        ctxt��������.Text = cgrdMain.TextMatrix(1, cgrdMain.Cols - 3)
        ctxt��ע.Text = cgrdMain.TextMatrix(1, cgrdMain.Cols - 2)
    End If
    
    Exit Sub
errhandler:
    sfsub������ "����֤������", "frm�������", "clstUnit_Click", Err.Number, Err.Description, False
    
End Sub

Private Sub Form_Load()
        
    On Error GoTo errhandler
    
    Dim lcol��������ť As New Collection '�������ϵİ�ť��ʼ�����ϡ�
    Set mobjGUI = New cls����ͨ�ö���
    Set mobjGUI.Form = Me
    Set mobjGUI.C������ = C������
    lcol��������ť.Add "ˢ��"
    lcol��������ť.Add "|"
    lcol��������ť.Add "��ӡ"
    lcol��������ť.Add "|"
    lcol��������ť.Add "�˳�"
    mobjGUI.subInitialize lcol��������ť, ""
    
    '��ȡ�����е�����Ա�ĵ�λ��
    subRefresh
    
    ctxt��������.Text = Format(Date, "yyyy-mm-dd")
    Exit Sub
errhandler:
    sfsub������ "����֤������", "frm�������", "Form_Load", Err.Number, Err.Description, False
    
End Sub
Private Sub subRefresh()
    Dim lobjRec As Object
    Dim lstr״̬���� As String
    
    On Error GoTo errhandler
    
    cgrdMain.Rows = 1
    
    If cchkType(0).Value = 1 And cchkType(1).Value = 0 Then
        lstr״̬���� = "(״̬='δ��ӡ' or isnull(������,'')='')"
    ElseIf cchkType(0).Value = 0 And cchkType(1).Value = 1 Then
        lstr״̬���� = "״̬='�Ѵ�ӡ'"
    End If
    Set lobjRec = pobj������.func��ȡ���뵥λ(lstr״̬����)
    clstUnit.Clear
    Do While Not lobjRec.EOF
        clstUnit.AddItem lobjRec(0).Value
        lobjRec.MoveNext
    Loop
    If clstUnit.ListCount > 0 Then clstUnit.ListIndex = 0

    Exit Sub
errhandler:
    sfsub������ "����֤������", "frm�������", "subRefresh", Err.Number, Err.Description, True
End Sub



Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    On Error GoTo errhandler
    Select Case Operate
    Case "ˢ��"
        subRefresh
    Case "��ӡ"
        Dim i As Long
        Dim lcolInfo As Collection
        Dim lstr������  As String
        
        If cgrdMain.Rows = 1 Then
            MsgBox "�����ݿɴ�ӡ��", vbOKOnly + vbExclamation, "ϵͳ��ʾ"
            Exit Sub
        End If
        If ctxt��������.Text = "" Then
            MsgBox "���������ʱ�ޣ�", vbOKOnly + vbExclamation, "ϵͳ��ʾ"
            ctxt��������.SetFocus
            Exit Sub
        End If
        If ctxt��������.Text = "" Then
            MsgBox "������������ڣ�", vbOKOnly + vbExclamation, "ϵͳ��ʾ"
            ctxt��������.SetFocus
            Exit Sub
        End If
        '��������ŵ���Ϣ��
        For i = 1 To cgrdMain.Rows - 1
            If cgrdMain.TextMatrix(i, 1) = "" Then
                If lstr������ = "" Then
                    lstr������ = pobj������.func���ɵ�����(cgrdMain.TextMatrix(i, 0))
                End If
                dafuncGetData "update ����֤����_��֤������Ϣ�� set ������='" & lstr������ & "',��������=" & ctxt��������.Text & ",��������='" & ctxt��������.Text & "',��ע='" & ctxt��ע.Text & "' where ϵͳ���='" & cgrdMain.TextMatrix(i, 0) & "'"
            Else
                lstr������ = cgrdMain.TextMatrix(i, 1)
            End If
        Next
        pobj������.sub��ӡ����֪ͨ lstr������
    
    End Select
    
    Exit Sub
errhandler:
    sfsub������ "����֤������", "frm�������", "mobjGUI_BeforeOperate", Err.Number, Err.Description, False
End Sub
