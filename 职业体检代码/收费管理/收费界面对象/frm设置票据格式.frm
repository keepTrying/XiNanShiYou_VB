VERSION 5.00
Object = "{8099FCC2-0A81-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm����Ʊ�ݸ�ʽ 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "����Ʊ�ݸ�ʽ"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   10800
   ClipControls    =   0   'False
   Icon            =   "frm����Ʊ�ݸ�ʽ.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   10800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5490
      Left            =   0
      TabIndex        =   8
      Top             =   840
      Width           =   10815
      Begin VB.Frame Frame10 
         Appearance      =   0  'Flat
         Caption         =   "Ʊ����Ϣ"
         ForeColor       =   &H80000008&
         Height          =   5130
         Left            =   6720
         TabIndex        =   10
         Top             =   120
         Width           =   3972
         Begin VB.OptionButton copt���� 
            Caption         =   "��"
            Height          =   252
            Index           =   1
            Left            =   1680
            TabIndex        =   19
            Top             =   3360
            Value           =   -1  'True
            Width           =   492
         End
         Begin VB.OptionButton copt���� 
            Caption         =   "��"
            Height          =   252
            Index           =   0
            Left            =   960
            TabIndex        =   18
            Top             =   3360
            Width           =   612
         End
         Begin VB.CommandButton Ccmd��� 
            Caption         =   "�����"
            Height          =   396
            Left            =   2640
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   3720
            Width           =   1125
         End
         Begin VB.ComboBox ccmb��Ӧҵ�� 
            Height          =   276
            ItemData        =   "frm����Ʊ�ݸ�ʽ.frx":0442
            Left            =   960
            List            =   "frm����Ʊ�ݸ�ʽ.frx":0444
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   2100
            Width           =   2868
         End
         Begin VB.TextBox cinbƱ�ݸ�ʽ�ļ����� 
            Height          =   372
            Left            =   120
            MaxLength       =   24
            TabIndex        =   6
            Top             =   4200
            Width           =   3708
         End
         Begin VB.TextBox cinbƱ������ 
            Height          =   360
            Left            =   960
            MaxLength       =   25
            TabIndex        =   1
            Top             =   960
            Width           =   2832
         End
         Begin VB.TextBox cinbƱ�ݱ�� 
            Enabled         =   0   'False
            Height          =   360
            Left            =   960
            MaxLength       =   2
            TabIndex        =   0
            Top             =   360
            Width           =   1290
         End
         Begin VB.ComboBox ccmbƱ������ 
            Height          =   276
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1560
            Width           =   2868
         End
         Begin VB.TextBox ctxt������� 
            Height          =   360
            Left            =   960
            MaxLength       =   2
            TabIndex        =   4
            Top             =   2640
            Width           =   1245
         End
         Begin MSComDlg.CommonDialog CcmnƱ�� 
            Left            =   3720
            Top             =   4680
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������Ŀ"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   17
            Top             =   3360
            Width           =   720
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ʊ������"
            Height          =   180
            Left            =   120
            TabIndex        =   16
            Top             =   1560
            Width           =   720
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��Ӧҵ��"
            Height          =   180
            Left            =   120
            TabIndex        =   15
            Top             =   2100
            Width           =   720
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ʽ�ļ���"
            Height          =   180
            Left            =   120
            TabIndex        =   14
            Top             =   3960
            Width           =   900
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���"
            Height          =   180
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   360
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ʊ������"
            Height          =   180
            Left            =   120
            TabIndex        =   12
            Top             =   960
            Width           =   720
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�������"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   11
            Top             =   2760
            Width           =   720
         End
      End
      Begin VSFlex6Ctl.vsFlexGrid cgrdMain 
         Height          =   5115
         Left            =   120
         TabIndex        =   9
         Top             =   150
         Width           =   6465
         _cx             =   165162124
         _cy             =   165159742
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
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   20
         Cols            =   7
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
   End
   Begin MSComctlLib.ImageList cimg��ťͼ�� 
      Left            =   5760
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar ctlb���� 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   10800
      _ExtentX        =   19050
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "cimg��ťͼ��"
      _Version        =   393216
   End
End
Attribute VB_Name = "frm����Ʊ�ݸ�ʽ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public pblnInUse As Boolean

Private WithEvents mobjGUI As cls����ͨ�ö���
Attribute mobjGUI.VB_VarHelpID = -1

Private Sub Ccmd���_Click()
    On Error GoTo errHandler
    Dim lstrFileName As String
    CcmnƱ��.InitDir = App.Path
    CcmnƱ��.ShowOpen
    lstrFileName = funcGetFileName(CcmnƱ��.filename)
    If Len(lstrFileName) > 14 Then
        sffuncMsg "�Բ����ļ����ƹ�������������14���֣�", sf����
    Else
        cinbƱ�ݸ�ʽ�ļ�����.Text = lstrFileName
    End If
    Exit Sub
errHandler:
    sfsub������ "�շѽ��沿��", "frm����Ʊ�ݸ�ʽ", "Ccmd���_Click", Err.Number, Err.Description, False
    
End Sub

Private Function funcGetFileName(filename As String) As String
    Dim lintOffset As Integer
    Dim lstrResult As String
    
    lstrResult = filename
    lintOffset = InStr(lstrResult, "\")
    If lintOffset = 0 Then
        funcGetFileName = filename
        Exit Function
    End If
    
    Do While lintOffset <> 0
        lstrResult = Mid(lstrResult, lintOffset + 1)
        lintOffset = InStr(lstrResult, "\")
    Loop
    funcGetFileName = lstrResult
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 39 Then
        KeyAscii = 0
    ElseIf KeyAscii = 13 Then
        SendKeys Chr(vbKeyTab)
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    
    If pblnInUse = True Then Exit Sub
    
    pblnInUse = True

    '��ʼ��������
    Dim lcol��������ť As Collection
    Set lcol��������ť = New Collection
    Set mobjGUI = New cls����ͨ�ö���
    Set mobjGUI.Form = Me
    Set mobjGUI.c������ = ctlb����

    lcol��������ť.Add "���"
    lcol��������ť.Add "ɾ��"
    lcol��������ť.Add "|"
    lcol��������ť.Add "����"
    lcol��������ť.Add "|"
    lcol��������ť.Add "�˳�"
    
    mobjGUI.subInitialize lcol��������ť, ""
    
    Dim lobjRec As Object
    Set lobjRec = dafuncGetData("select ���,Ʊ�����ͱ��,Ʊ������,Ʊ�ݸ�ʽ�ļ�����,��Ӧҵ��,�������,��Ŀ���� from �շѹ���_Ʊ��������Ϣ�� order by ���")
    Set cgrdMain.DataSource = lobjRec
    cgrdMain.Row = 0
    cgrdMain.ColHidden(1) = True
    
    '��ʼ��Ʊ������
    Set lobjRec = dafuncGetData("select * from �շѹ���_Ʊ�������ֵ���ͼ")
    If (Not lobjRec.EOF) And (Not lobjRec.BOF) Then
        Do While (Not lobjRec.EOF)
            ccmbƱ������.AddItem lobjRec.Fields("����").Value
            ccmbƱ������.ItemData(ccmbƱ������.NewIndex) = lobjRec.Fields("innerId").Value
            lobjRec.MoveNext
        Loop
        lobjRec.MoveFirst
    End If
    If ccmbƱ������.ListCount > 0 Then
        ccmbƱ������.ListIndex = 0
    End If
    
    
    ccmb��Ӧҵ��.AddItem "һ��", 0
    ccmb��Ӧҵ��.AddItem "����", 1
    Exit Sub
errHandler:
    sfsub������ "�շѽ��沿��", "frm����Ʊ�ݸ�ʽ", "Form_Load", Err.Number, Err.Description, False
    Exit Sub
    Resume
End Sub

Private Sub Form_Unload(Cancel As Integer)
    pblnInUse = False
End Sub

Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    On Error GoTo errHandler
    Select Case Operate
    Case "���"
        Cancel = True
        cinbƱ�ݱ��.Text = ""
        cinbƱ������.Text = ""
        ctxt�������.Text = 4
        cinbƱ�ݸ�ʽ�ļ�����.Text = ""
        cinbƱ������.SetFocus
    
    Case "����"
        Cancel = True
        subValidate
         
        pobj�շѹ���.sub����Ʊ������ cinbƱ�ݱ��.Text, cinbƱ������.Text, cinbƱ�ݸ�ʽ�ļ�����.Text, ccmbƱ������.ItemData(ccmbƱ������.ListIndex), ccmb��Ӧҵ��.Text, Val(ctxt�������.Text), IIf(copt����(0).Value, "��", "��")
        
        Dim lobjRec As Object
        Set lobjRec = dafuncGetData("select ���,Ʊ�����ͱ��,Ʊ������,Ʊ�ݸ�ʽ�ļ�����,��Ӧҵ��,�������,��Ŀ���� from �շѹ���_Ʊ��������Ϣ�� order by ���")
        Set cgrdMain.DataSource = lobjRec
        cgrdMain.Row = 0
        cgrdMain.ColHidden(1) = True
        mobjGUI_BeforeOperate "���", True
        
    Case "ɾ��"
        Cancel = True
    End Select
    
    Set lobjRec = Nothing
    Exit Sub
    
errHandler:
    sfsub������ "�շѽ��沿��", "frm����Ʊ�ݸ�ʽ", "mobjGUI_BeforeOperate", Err.Number, Err.Description, False
        
End Sub

Private Sub subValidate()
    '�ж��û�¼���Ƿ����
    If cinbƱ�ݸ�ʽ�ļ�����.Text = "" Or cinbƱ������.Text = "" Or ccmb��Ӧҵ��.Text = "" Or ccmbƱ������.Text = "" Then
        Err.Raise 6666, , "¼�벻������������¼�룡"
    End If
    If ctxt�������.Text = "" Then
        Err.Raise 6666, , "���������������"
    End If
    If Not IsNumeric(ctxt�������.Text) Then
        If ctxt�������.Text < 1 Then
            Err.Raise 6666, , "�������¼��Ƿ���"
        End If
    End If
    
    '�жϸü�¼�Ƿ����
    Dim linttemp As Integer
    If cgrdMain.Rows > 0 Then
        For linttemp = 1 To cgrdMain.Rows - 1
            If linttemp <> cgrdMain.Row Then
                If cgrdMain.Cell(flexcpText, linttemp, 1) = ccmbƱ������.ItemData(ccmbƱ������.ListIndex) And _
                    cgrdMain.Cell(flexcpText, linttemp, 3) = cinbƱ�ݸ�ʽ�ļ�����.Text And _
                    cgrdMain.Cell(flexcpText, linttemp, 4) = ccmb��Ӧҵ��.Text And _
                    cgrdMain.Cell(flexcpText, linttemp, 5) = ctxt�������.Text Then
                    Err.Raise 6666, , "��Ʊ��������Ϣ�Ѵ��ڣ����޸ģ�"
                End If
            End If
        Next
    End If
End Sub



Private Sub cgrdMain_Click()
    Dim lobjRec As Object
    Dim lstr��� As String
    
    On Error GoTo errHandler
    
    If cgrdMain.RowSel = 0 Then
        Exit Sub
    End If
    Set lobjRec = dafuncGetData("select * from �շѹ���_Ʊ�������ֵ���ͼ")
    cinbƱ�ݱ��.Text = cgrdMain.Cell(flexcpText, cgrdMain.RowSel, 0)
    cinbƱ������.Text = cgrdMain.Cell(flexcpText, cgrdMain.RowSel, 2)
    cinbƱ�ݸ�ʽ�ļ�����.Text = cgrdMain.Cell(flexcpText, cgrdMain.RowSel, 3)
    ctxt�������.Text = cgrdMain.Cell(flexcpText, cgrdMain.RowSel, 5)
    If (Not lobjRec.EOF) And (Not lobjRec.BOF) Then
        lobjRec.MoveFirst
        Do While (Not lobjRec.EOF)
            lstr��� = lobjRec("innerID").Value
            If lstr��� = cgrdMain.Cell(flexcpText, cgrdMain.RowSel, 1) Then
               ccmbƱ������.Text = lobjRec("����")
               Exit Do
            Else
                lobjRec.MoveNext
            End If
            If lobjRec.EOF Then
                MsgBox "����޸���Ʊ�������ֵ�����������ø���Ŀ��Ʊ�����ͣ�����ĿƱ�����������⣡", vbExclamation, "ϵͳ��ʾ"
                Exit Do
            End If
       Loop
    End If
    
    If cgrdMain.Cell(flexcpText, cgrdMain.RowSel, 4) = "һ��" Then
        ccmb��Ӧҵ��.ListIndex = 0
    Else
        ccmb��Ӧҵ��.ListIndex = 1
    End If
    
    If cgrdMain.Cell(flexcpText, cgrdMain.RowSel, 6) = "��" Then
        copt����(0).Value = True
    Else
        copt����(1).Value = True
    End If
    
    cinbƱ������.SetFocus
    
    Set lobjRec = Nothing
    Exit Sub
    
errHandler:
    sfsub������ "�շѽ��沿��", "frm����Ʊ�ݸ�ʽ", "cgrdMain_Click", Err.Number, Err.Description, False
    
End Sub


