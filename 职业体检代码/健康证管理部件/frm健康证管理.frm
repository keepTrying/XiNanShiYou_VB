VERSION 5.00
Object = "{8099FCC2-0A81-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frm����֤���� 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ʳƷ"
   ClientHeight    =   7575
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   11715
   ClipControls    =   0   'False
   Icon            =   "frm����֤����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   11715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.OptionButton copt���� 
      Caption         =   "����"
      Height          =   255
      Left            =   4200
      TabIndex        =   12
      Top             =   840
      Width           =   855
   End
   Begin VB.OptionButton coptʳƷ 
      Caption         =   "ʳƷ"
      Height          =   255
      Left            =   3360
      TabIndex        =   11
      Top             =   840
      Value           =   -1  'True
      Width           =   735
   End
   Begin VB.TextBox ctxtϵͳ��� 
      Height          =   270
      Left            =   8640
      TabIndex        =   9
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox ctxtNum 
      Height          =   270
      Left            =   6120
      TabIndex        =   6
      Text            =   "10"
      Top             =   840
      Width           =   495
   End
   Begin VB.CheckBox cchkType 
      BackColor       =   &H00FFDFFE&
      Caption         =   "����"
      Height          =   255
      Index           =   2
      Left            =   2280
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CheckBox cchkType 
      BackColor       =   &H00D1F7FE&
      Caption         =   "�Ѵ�ӡ"
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.CheckBox cchkType 
      BackColor       =   &H00D2FCCF&
      Caption         =   "δ��ӡ"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VSFlex6Ctl.vsFlexGrid cgrdMain 
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   10815
      _cx             =   25315972
      _cy             =   25307929
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
      BackColorAlternate=   16777215
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
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   "���   |����    |�Ա�    |����    |��λ����     |����    |ְҵ    |�������   | ������ |����֤��"
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
   Begin MSComctlLib.Toolbar C������ 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   11715
      _ExtentX        =   20664
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
      Caption         =   "�����(����/����)��"
      Height          =   180
      Index           =   1
      Left            =   6720
      TabIndex        =   10
      Top             =   840
      Width           =   1890
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "��"
      Height          =   180
      Left            =   2160
      TabIndex        =   7
      Top             =   8280
      Width           =   180
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ѡ��ǰ�棺"
      Height          =   180
      Index           =   0
      Left            =   5280
      TabIndex        =   5
      Top             =   840
      Width           =   900
   End
   Begin VB.Label clblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������"
      Height          =   180
      Left            =   10800
      TabIndex        =   4
      Top             =   840
      Width           =   540
   End
   Begin VB.Menu cmnuView 
      Caption         =   "�鿴(&V)"
      Begin VB.Menu cmnuItemView 
         Caption         =   "��ѯ(&Q)"
         Index           =   1
      End
      Begin VB.Menu cmnuItemView 
         Caption         =   "ˢ��(&R)"
         Index           =   2
      End
      Begin VB.Menu cmnuItemView 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu cmnuItemView 
         Caption         =   "�˳�(&Esc)"
         Index           =   4
      End
   End
   Begin VB.Menu cmnuInput 
      Caption         =   "¼��(&I)"
      Visible         =   0   'False
      Begin VB.Menu cmnuItemInput 
         Caption         =   "����(&N)"
         Index           =   1
      End
      Begin VB.Menu cmnuItemInput 
         Caption         =   "�޸�(&E)"
         Index           =   2
      End
      Begin VB.Menu cmnuItemInput 
         Caption         =   "ɾ��(&D)"
         Index           =   3
      End
   End
   Begin VB.Menu cmnuPrint 
      Caption         =   "��ӡ(&p)"
      Begin VB.Menu cmnuItemPrint 
         Caption         =   "����֤(&Z)"
         Index           =   1
      End
      Begin VB.Menu cmnuItemPrint 
         Caption         =   "����֪ͨ"
         Index           =   2
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frm����֤����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public pblnInUse As Boolean

Private WithEvents mobjGUI As cls����ͨ�ö��� '���������õĽ���ͨ�ö�
Attribute mobjGUI.VB_VarHelpID = -1

'��ѯ������
Private mstrϵͳ��� As String
Private mstr���� As String
Private mstr��λ As String
Private mstr������ڴ� As String
Private mstr������ڵ� As String
Private mstr���� As String
Private mstr��֤��λ As String

Private mobjRec As Object

Private mcolIndex As Collection

Private Sub cchkType_Click(Index As Integer)
    subRefresh
End Sub

Private Sub cgrdMain_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 1 Then Cancel = True
End Sub

Private Sub cmnuItemInput_Click(Index As Integer)

    Dim lobj��� As cls���
    
    On Error GoTo errhandler
    Select Case Index
    Case 1 '����
        frm���¼��.pstrϵͳ��� = ""
        frm���¼��.Show 1, Me
'        frm���¼��.Move Me.Left, Me.Top
        'ˢ�½��档
        subRefresh
    
    Case 2 '�޸�
        If cgrdMain.Row = 0 Then
            MsgBox "��ѡ��Ҫ�޸ĵ������Ա��", vbOKOnly + vbExclamation, "ϵͳ��ʾ"
            Exit Sub
        End If
        
        frm���¼��.pstrϵͳ��� = cgrdMain.TextMatrix(cgrdMain.Row, 0)
        frm���¼��.Show 1, Me
        
        'ˢ�½��档
        subRefresh
    
    Case 3 'ɾ��
        Set lobj��� = New cls���
        If cgrdMain.Row = 0 Then
            MsgBox "��ѡ��Ҫɾ���������Ա��", vbOKOnly + vbExclamation, "ϵͳ��ʾ"
            Exit Sub
        End If
        If MsgBox("��ȷ��Ҫɾ����" & cgrdMain.TextMatrix(cgrdMain.Row, 1) & "��������¼��", vbYesNo + vbQuestion, "ϵͳѯ��") = vbNo Then
            Exit Sub
        End If
        lobj���.ϵͳ��� = cgrdMain.TextMatrix(cgrdMain.Row, 0)
        lobj���.subɾ��
        cgrdMain.RemoveItem cgrdMain.Row
    
    End Select
    
    Exit Sub
errhandler:
    sfsub������ "����֤������", "frm����֤����", "cmnuItemInput_Click", Err.Number, Err.Description, False
    
End Sub

Private Sub cmnuItemPrint_Click(Index As Integer)
    Dim i As Long
    Dim lobj��� As cls���
    Dim lstrSet As String
    Dim lobjRec As Object
    Dim lobjCN As Object
    On Error GoTo errhandler
    
    Select Case Index
    Case 1 '����֤
        Dim lcolInfo As Collection
        Dim lstrCN As String
        Dim lbln�������� As Boolean
        Dim lstr���ǰ׺ As String
        Dim lstr��������� As String
        '����ҵ�����ã��ж��Ƿ���Ҫ�Զ����ɽ���֤�š�
        '˵����ʡ�����Ʊ���ʹ�ô�����Ľ���֤��
'        lstrSet = pobj������.ҵ������("����֤������")
'        lstrSet = "��"
'        If lstrSet = "��" Or pobj������.ҵ������("�ֹ����뽡��֤��") = "��" Then
        
            '�û����뽡��֤�ŵ���ʼ�š�
'            lstrCN = InputBox("�����뽡��֤����ʼ��", "����")
'            If lstrCN = "" Then
'                Exit Sub
'            End If
            
            '�ж����뽡��֤���Ƿ�Ϊ���֡�
'            Do While Not (IsNumeric(lstrCN))
'                If MsgBox("������Ľ���֤�Ÿ�ʽ���ԡ��Ƿ��������룿", vbYesNo, "ϵͳ��ʾ") = vbYes Then
'                    lstrCN = InputBox("�����뽡��֤����ʼ��", "����")
'                Else
'                    Exit Sub
'                End If
'            Loop
            '��鿨�ų��ȣ�12λΪ�°濨�ţ�����Ϊ�ϰ濨�ţ�Ҫ�ж�����ʹ�õ��������
'            Dim lobjCheck As New clsCheck
'
'            If Len(lstrCN) <> 12 Then
''                If lobjCheck.funcCheckExpireDate() Then
''                    Err.Raise 6666, , "��ǰϵͳ����ʶ�����ֿ���"
''                End If
'                'δ���ڣ��жϿ��Ƿ�Ϸ���
'                Dim lobjEncrypt As Object
'                Set lobjEncrypt = CreateObject("fycarddes.clsDataEncrypt")
'                If Not lobjEncrypt.funcCheckJkzCardno(lstrCN) Then
'                    Err.Raise 6666, , "ϵͳ�޷�ʶ�����ſ�����ȷ�����Ƿ����𻵣�"
'                End If
'            Else
'                '�°濨�ţ��жϿ��Ƿ�Ϸ���
'                If Not funcCheckCardno(lstrCN) Then
'                    Err.Raise 6666, , "ϵͳ�޷�ʶ�����ſ�����ȷ��������ָ���ĸ�ʽ��"
'                End If
'                '�жϿ����Ƿ񳬳���˾�Ĵ�ӡ��Χ
'                If Not lobjCheck.funcCheckMaxNo(lstrCN) Then
'                    Err.Raise 6666, , "ϵͳ�޷�ʶ�����ſ����ÿ�����ָ����Ӧ�̷��еĿ���"
'                End If
'            End If
'            lstrCN = Left(lstrCN, Len(lstrCN) - 2)  'ȥ��У��λ
'        Else
'            'ϵͳ�Զ����ɽ���֤��
'            lstrCN = ""
'        End If
        
        lstrCN = ""
        '��ȡѡ�е�ϵͳ��ţ�����������
        Set lcolInfo = New Collection
        For i = 1 To cgrdMain.Rows - 1
            If cgrdMain.Cell(flexcpChecked, i, 1) = flexChecked Then
                Set lobj��� = New cls���
                lobj���.ϵͳ��� = cgrdMain.TextMatrix(i, 0)
                lobj���.���� = cgrdMain.TextMatrix(i, mcolIndex("Ѫ��"))
                If lobj���.���� = "����" Then
                    Err.Raise 6666, , "������Ա���ܴ�ӡ����֤���벻Ҫѡ�е�����Ա��"
                End If
                Set lobjCN = dafuncGetData("EXEC ����֤����_���ɽ���֤���")
                lstrCN = lobjCN(0)
                Set lobjCN = Nothing
                Set lobjCN = dafuncGetData("SELECT ���������� FROM ϵͳ����_ϵͳ�������ñ�")
                lstr��������� = lobjCN(0)
                
                Select Case lstr���������
                    Case "1"
                        lstr���ǰ׺ = "A"
                    Case "2"
                        lstr���ǰ׺ = "B"
                    Case "3"
                        lstr���ǰ׺ = "C"
                    Case Else
                        lstr���ǰ׺ = "D"
                End Select
                
                lobj���.����֤�� = lstr���ǰ׺ & Right(lstrCN, 7) '�����

'                lobj���.����֤�� = lobj���.���ϵͳ���
                
                If lobj���.���� = "ʳƷ����" And lobj���.���֤�� = "" Then   'ʳƷ֤��֤�Ŵ���ڡ����֤�š��У������ر�
                    lobj���.���֤�� = Mid(pobj������.func���ɽ���֤��(), 2)     'ϵͳ���ɵĺ���6λ��ֻҪ��5λ
                End If
                
                '��������ϵͳ�������ļ�¼��û�з�֤���ںͷ�֤��λ��
                If lobj���.��֤���� = "" Then
                    lobj���.��֤���� = Format(Date, "yyyy-mm-dd")
                End If
                If lobj���.��Ч���� = "" Then
                    lobj���.��Ч���� = Format(DateAdd("d", -1, DateAdd("yyyy", 1, Date)), "yyyy-mm-dd")
                End If
                If lobj���.��֤��λ = "" Then
                    lobj���.��֤��λ = um����վ��
                End If
                                
                lcolInfo.Add lobj���
                
                
                '����֤���Զ�������
'                lstrCN = Format(Val(lstrCN) + 1, String(Len(lstrCN), "0"))
                
            End If
        Next
        
        If lcolInfo.Count = 0 Then
            MsgBox "��ѡ��Ҫ��ӡ�������Ա�������ϴ򹴣���", vbOKOnly + vbExclamation, "ϵͳ��ʾ"
            Exit Sub
        End If
        
'        pobj������.sub��ӡ����֤ lcolInfo
        Dim frm As Form
        Set frm = frmPrintPVCCard
        Set frm.Cards = lcolInfo
        frm.Show 1
        Unload frm
        'ˢ�½��档
        subRefresh
        
    Case 2 '����֪ͨ��
        'frm�������.Show 1, Me
        
    End Select

    Exit Sub
errhandler:
    sfsub������ "����֤������", "frm����֤����", "cmnuItemPrint_Click", Err.Number, Err.Description, False
End Sub

Private Sub cmnuItemView_Click(Index As Integer)
    On Error GoTo errhandler
    
    Select Case Index
    Case 1 '��ѯ
        frm��ѯ.Show 1, Me
        
        If frm��ѯ.pblnOk Then
            mstr���� = frm��ѯ.pstrName
            mstrϵͳ��� = frm��ѯ.pstrNo
            mstr������ڴ� = frm��ѯ.pstrStartDate
            mstr������ڵ� = frm��ѯ.pstrEndDate
            mstr��λ = frm��ѯ.pstrUnit
            'mstr���� = frm��ѯ.pstrType
            mstr���� = IIf(coptʳƷ.Value, "ʳƷ����", "")
            mstr��֤��λ = frm��ѯ.pstr��֤��λ
            
            subRefresh
        End If

    Case 2 'ˢ�¡�
        subRefresh
        
    Case 4
        Unload frm���¼��
        Unload Me
    End Select

    Exit Sub
errhandler:
    sfsub������ "����֤���沿��", "frm����֤����", "cmnuItemView_Click", Err.Number, Err.Description, False
    Exit Sub
    Resume
End Sub


Private Sub copt����_Click()
    mstr���� = ""
    subRefresh
End Sub

Private Sub coptʳƷ_Click()
    mstr���� = "ʳƷ����"
    subRefresh
End Sub

Private Sub ctxtNum_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    Dim llngNum As Long
    On Error GoTo errhandler
    llngNum = Val(ctxtNum.Text)
    If llngNum > cgrdMain.Rows - 1 Then
        llngNum = cgrdMain.Rows - 1
    End If
    For i = 1 To cgrdMain.Rows - 1
        cgrdMain.Cell(flexcpChecked, i, 1) = flexUnchecked
    Next
    For i = 1 To llngNum
        cgrdMain.Cell(flexcpChecked, i, 1) = flexChecked
    Next
    Exit Sub
errhandler:
End Sub

Private Sub ctxtϵͳ���_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long, j As Long
    Dim lobjRec As Object
    On Error GoTo errhandler
    If KeyCode = 13 And ctxtϵͳ��� <> "" Then
        For i = 1 To cgrdMain.Rows - 1
            If cgrdMain.TextMatrix(i, mcolIndex("���ϵͳ���")) = ctxtϵͳ��� Then
                cgrdMain.TopRow = i
                cgrdMain.Row = i
                Exit Sub
            End If
        Next
        
        '����������û���ҵ������š��ӿ����ҡ�
        Set lobjRec = pobj������.func��������ѯ("", "", "", "", "", "", "", "", ctxtϵͳ���)
        If lobjRec.RecordCount > 0 Then
            cgrdMain.Rows = cgrdMain.Rows + 1
            i = cgrdMain.Rows - 1
            cgrdMain.Cell(flexcpChecked, i, 1) = flexUnchecked
            For j = 0 To cgrdMain.Cols - 1
                cgrdMain.TextMatrix(i, j) = IIf(IsNull(lobjRec(j)), "", lobjRec(j))
            Next
            cgrdMain.AutoSize 0, cgrdMain.Cols - 1
            
            '��ʾ��ɫ��
            If lobjRec!���� = "������֤" Then
                If lobjRec!״̬ = "δ��ӡ" Then
                    cgrdMain.Cell(flexcpBackColor, i, 0, i, cgrdMain.Cols - 1) = cchkType(0).BackColor
                Else
                    cgrdMain.Cell(flexcpBackColor, i, 0, i, cgrdMain.Cols - 1) = cchkType(1).BackColor
                End If
            Else
                cgrdMain.Cell(flexcpBackColor, i, 0, i, cgrdMain.Cols - 1) = cchkType(2).BackColor
            End If
            cgrdMain.ColWidth(1) = 1000
            
            '����ϵͳ��š�
            cgrdMain.ColHidden(0) = True
            
            cgrdMain.TopRow = i
            cgrdMain.Row = i
            clblInfo.Caption = "������" & cgrdMain.Rows - 1
        End If
        
        ctxtϵͳ���.Text = ""
    End If
    Exit Sub
errhandler:
    sfsub������ "����֤������", "frm����֤����", "ctxtϵͳ���_KeyDown", Err.Number, Err.Description, False
    
End Sub

Private Sub Form_Load()

    On Error GoTo errhandler

    If pblnInUse Then Exit Sub
    pblnInUse = True
    
    Dim lcol��������ť As New Collection '�������ϵİ�ť��ʼ�����ϡ�
    Set mobjGUI = New cls����ͨ�ö���
    Set mobjGUI.Form = Me
    Set mobjGUI.C������ = C������
    lcol��������ť.Add "ˢ��"
    lcol��������ť.Add "|"
    lcol��������ť.Add "���"
    lcol��������ť.Add "�޸�"
    lcol��������ť.Add "ɾ��"
    lcol��������ť.Add "|"
    lcol��������ť.Add "��ӡ֤(&Z)107"
    lcol��������ť.Add "����֪ͨ(&T)107"
    lcol��������ť.Add "|"
    lcol��������ť.Add "�˳�"
    mobjGUI.subInitialize lcol��������ť, ""
    '������ʹ�ñ���ϵͳ¼֤
    C������.Buttons(3).Visible = False
    C������.Buttons(4).Visible = False
    C������.Buttons(5).Visible = False
    C������.Buttons(6).Visible = False
    
    'Ȩ���жϡ�
'    If Not umfuncУ���û�Ȩ��("����֤����_¼��") Then
'        C������.Buttons(3).Visible = False
'        C������.Buttons(4).Visible = False
'        If Not umfuncУ���û�Ȩ��("����֤����_ɾ��") Then
'            C������.Buttons(5).Visible = False
'            C������.Buttons(6).Visible = False
'            cmnuInput.Visible = False
'        Else
'            cmnuItemInput(1).Visible = False
'            cmnuItemInput(2).Visible = False
'        End If
'    Else
'        If Not umfuncУ���û�Ȩ��("����֤����_ɾ��") Then
'            C������.Buttons(5).Visible = False
'            cmnuItemInput(3).Visible = False
'        End If
'    End If
    If Not umfuncУ���û�Ȩ��("����֤����_��ӡ") Then
        C������.Buttons(7).Visible = False
        C������.Buttons(8).Visible = False
        C������.Buttons(9).Visible = False
        cmnuPrint.Visible = False
    End If
    
    C������.Buttons(8).Visible = False
    
    '��ȡ������ܵ�δ��ӡ����¼��
    mstr������ڴ� = Format(DateAdd("d", 1 - DatePart("w", Now, vbMonday), Now) - 7, "yyyy-mm-dd")
    mstr������ڵ� = Format(Now, "yyyy-mm-dd")
    mstr���� = "ʳƷ����"
    
    subRefresh
    
    Exit Sub
errhandler:
    sfsub������ "����֤������", "frm����֤����", "Form_Load", Err.Number, Err.Description, False
End Sub

'���ܣ����ݲ�ѯ������ʾ��ѯ�����
Public Sub subRefresh()
    
    Dim lstr״̬���� As String
    Dim i As Long
    lstr״̬���� = ""
    
    '���ȵ������ݡ�
    On Error Resume Next
    dafuncGetData "exec ����֤����_��������������Ա��Ϣ"
    
    On Error GoTo errhandler
    If cchkType(0).Value = 1 Or cchkType(1).Value = 1 Then
        lstr״̬���� = "(����='������֤'"
        If cchkType(0).Value = 1 And cchkType(1).Value = 0 Then
            lstr״̬���� = lstr״̬���� & " and ״̬='δ��ӡ'"
        ElseIf cchkType(0).Value = 0 And cchkType(1).Value = 1 Then
            lstr״̬���� = lstr״̬���� & " and ״̬='�Ѵ�ӡ'"
        End If
        lstr״̬���� = lstr״̬���� & ")"
    End If
    If cchkType(2).Value = 1 Then
        lstr״̬���� = lstr״̬���� & IIf(lstr״̬���� = "", "", " or ") & "����='����'"
    End If
    If lstr״̬���� = "" Then lstr״̬���� = "1=0"
    
    If mstr������ڴ� = "" Then mstr������ڴ� = DateAdd("d", -30, Date)
    
    Set mobjRec = pobj������.func��������ѯ(mstrϵͳ���, mstr����, mstr��λ, mstr������ڴ�, mstr������ڵ�, mstr����, lstr״̬����, mstr��֤��λ)
    
    cgrdMain.FormatString = ""
    Set cgrdMain.DataSource = mobjRec
    For i = 1 To cgrdMain.Rows - 1
        cgrdMain.Cell(flexcpChecked, i, 1) = flexUnchecked
        '��ʾ��ɫ��
        If mobjRec!���� = "������֤" Then
            If mobjRec!״̬ = "δ��ӡ" Then
                cgrdMain.Cell(flexcpBackColor, i, 0, i, cgrdMain.Cols - 1) = cchkType(0).BackColor
                
            Else
                cgrdMain.Cell(flexcpBackColor, i, 0, i, cgrdMain.Cols - 1) = cchkType(1).BackColor
            End If
        Else
            cgrdMain.Cell(flexcpBackColor, i, 0, i, cgrdMain.Cols - 1) = cchkType(2).BackColor
        End If
        mobjRec.MoveNext
    Next
    Set mcolIndex = New Collection
    For i = 0 To cgrdMain.Cols - 1
        mcolIndex.Add i, cgrdMain.TextMatrix(0, i)
    Next
    
    cgrdMain.ColWidth(1) = 1000
    
    '����ϵͳ��š�
    cgrdMain.ColHidden(0) = True
    
    clblInfo.Caption = "������" & cgrdMain.Rows - 1

    Exit Sub
errhandler:
    sfsub������ "����֤������", "frm����֤����", "subRefresh", Err.Number, Err.Description, True
End Sub



Private Sub Form_Resize()
    On Error Resume Next
    cgrdMain.Width = Me.ScaleWidth - cgrdMain.Left - 60
    cgrdMain.Height = Me.ScaleHeight - cgrdMain.Top - 60
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set mobjGUI = Nothing
    Set mobjRec = Nothing
    pblnInUse = False
End Sub

Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    On Error GoTo errhandler
    
    Select Case Operate
    Case "ˢ��"
        cmnuItemView_Click 2
    Case "���"
        cmnuItemInput_Click 1
    Case "�޸�"
        cmnuItemInput_Click 2
    Case "ɾ��"
        cmnuItemInput_Click 3
    Case "��ӡ֤"
        cmnuItemPrint_Click 1
    Case "����֪ͨ"
        cmnuItemPrint_Click 2
    End Select
    
    Exit Sub
errhandler:
    sfsub������ "����֤������", "frm����֤����", "mobjGUI_BeforeOperate", Err.Number, Err.Description, False
    
End Sub
Function funcCheckCardno(paraCardno As String) As Boolean
    Dim i As Long
    
    i = CLng(Left(paraCardno, 10))
    If Format(((i Mod 99) * 3) Mod 75, "00") = Right(paraCardno, 2) Then
        funcCheckCardno = True
    Else
        funcCheckCardno = False
    End If
End Function
