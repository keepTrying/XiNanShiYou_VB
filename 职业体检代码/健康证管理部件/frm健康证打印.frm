VERSION 5.00
Object = "{8099FCC2-0A81-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frm����֤���� 
   Caption         =   "����֤����"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frm����֤��ӡ.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  '����ȱʡ
   WindowState     =   2  'Maximized
   Begin VB.TextBox ctxtNum 
      Height          =   270
      Left            =   1080
      TabIndex        =   8
      Text            =   "10"
      Top             =   8280
      Width           =   855
   End
   Begin VB.CheckBox cchkType 
      BackColor       =   &H00C0C0FF&
      Caption         =   "����"
      Height          =   255
      Index           =   2
      Left            =   4680
      TabIndex        =   5
      Top             =   840
      Width           =   1575
   End
   Begin VB.CheckBox cchkType 
      BackColor       =   &H00C0FFFF&
      Caption         =   "�Ѵ�ӡ"
      Height          =   255
      Index           =   1
      Left            =   3120
      TabIndex        =   4
      Top             =   840
      Width           =   1575
   End
   Begin VB.CheckBox cchkType 
      BackColor       =   &H00C0FFC0&
      Caption         =   "δ��ӡ"
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   3
      Top             =   840
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin MSComctlLib.Toolbar ctbMain 
      Align           =   1  'Align Top
      Height          =   705
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1244
      ButtonWidth     =   2037
      ButtonHeight    =   1085
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "��ѯ(&Q)"
            Key             =   "query"
            ImageKey        =   "query"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "����(&N)"
            Key             =   "new"
            ImageKey        =   "new"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "�޸�(&U)"
            Key             =   "update"
            ImageKey        =   "update"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "ɾ��(&D)"
            Key             =   "delete"
            ImageKey        =   "delete"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "��ӡ(&P)"
            Key             =   "print"
            ImageKey        =   "print"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "����֪ͨ(&D)"
            Key             =   "dl"
            ImageKey        =   "dl"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "�˳�(&E)"
            Key             =   "exit"
            ImageKey        =   "exit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   21
      ImageHeight     =   21
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm����֤��ӡ.frx":0E42
            Key             =   "query"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm����֤��ӡ.frx":115C
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm����֤��ӡ.frx":1476
            Key             =   "dl"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm����֤��ӡ.frx":18C8
            Key             =   "new"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm����֤��ӡ.frx":1D1A
            Key             =   "update"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm����֤��ӡ.frx":216C
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm����֤��ӡ.frx":2486
            Key             =   "print"
         EndProperty
      EndProperty
   End
   Begin VSFlex6Ctl.vsFlexGrid cgrdMain 
      Height          =   6975
      Left            =   -360
      TabIndex        =   1
      Top             =   1200
      Width           =   11655
      _cx             =   23810126
      _cy             =   23801871
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
      ExplorerBar     =   0
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "��"
      Height          =   180
      Left            =   2160
      TabIndex        =   9
      Top             =   8280
      Width           =   180
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "ѡ��ǰ�棺"
      Height          =   180
      Left            =   120
      TabIndex        =   7
      Top             =   8280
      Width           =   900
   End
   Begin VB.Label clblInfo 
      AutoSize        =   -1  'True
      Caption         =   "������"
      Height          =   180
      Left            =   6840
      TabIndex        =   6
      Top             =   840
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����Ϣ"
      Height          =   180
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   720
   End
End
Attribute VB_Name = "frm����֤����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'��ѯ������
Private mstrϵͳ��� As String
Private mstr���� As String
Private mstr��λ As String
Private mstr������ڴ� As String
Private mstr������ڵ� As String
Private mstr���� As String

Private mobjRec As Object

Private Sub cchkType_Click(Index As Integer)
    subRefresh
End Sub

Private Sub cgrdMain_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 1 Then Cancel = True
End Sub

Private Sub ctbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim lobj��� As cls���
    Dim i As Long
    Dim j As Long
    
    Select Case Button.Key
    Case "query"
        frm��ѯ.Show 1, Me
        
        If frm��ѯ.pblnOk Then
            mstr���� = frm��ѯ.pstrName
            mstrϵͳ��� = frm��ѯ.pstrNo
            mstr������ڴ� = frm��ѯ.pstrStartDate
            mstr������ڵ� = frm��ѯ.pstrEndDate
            mstr���� = frm��ѯ.pstrType
            
            subRefresh
        End If
        
    Case "new"
        frm���¼��.pstrϵͳ��� = ""
        frm���¼��.Show 1, Me
        'ˢ�½��档
        subRefresh
        
    Case "update"
        If cgrdMain.Row = 0 Then
            MsgBox "��ѡ��Ҫ�޸ĵ������Ա��", vbOKOnly + vbExclamation, "ϵͳ��ʾ"
            Exit Sub
        End If
        
        frm���¼��.pstrϵͳ��� = cgrdMain.TextMatrix(cgrdMain.Row, 0)
        frm���¼��.Show 1, Me
        
        'ˢ�½��档
        subRefresh
        
    Case "print" '��ӡ����֤
        Dim lcolInfo As Collection
        Dim lstrCN As String
        
        '����ҵ�����ã��ж��Ƿ���Ҫ�Զ����ɽ���֤�š�
        If pobj������.ҵ������("����֤������") = "��" Then
        
            '�û����뽡��֤�ŵ���ʼ�š�
            lstrCN = InputBox("�����뽡��֤����ʼ��", "����")
            If lstrCN = "" Then
                Exit Sub
            End If
            
            '�ж����뽡��֤���Ƿ�Ϊ���֡�
            Do While Not (IsNumeric(lstrCN))
                If MsgBox("������Ľ���֤�Ÿ�ʽ���ԡ��Ƿ��������룿", vbYesNo, "ϵͳ��ʾ") = vbYes Then
                    lstrCN = InputBox("�����뽡��֤����ʼ��", "����")
                Else
                    Exit Sub
                End If
            Loop
            
            '�޸ģ�2002-05-6������жϿ��Ƿ�Ϸ���
            Dim lobjEncrypt As Object
            Set lobjEncrypt = CreateObject("fycarddes.clsDataEncrypt")
            If Not lobjEncrypt.funcCheckJkzCardno(lstrCN) Then
                Err.Raise 6666, , "ϵͳ�޷�ʶ�����ſ�����ȷ��������ָ���ĸ�ʽ���Ƿ����𻵣�"
            End If
            '������У��λ��
            lstrCN = lobjEncrypt.����
            Set lobjEncrypt = Nothing
        Else
            lstrCN = pobj������.func���ɽ���֤��(False)
        End If
        
        
        '��ȡѡ�е�ϵͳ��ţ�����������
        Set lcolInfo = New Collection
        For i = 1 To cgrdMain.Rows - 1
            If cgrdMain.Cell(flexcpChecked, i, 1) = flexChecked Then
                Set lobj��� = New cls���
                lobj���.ϵͳ��� = cgrdMain.TextMatrix(i, 0)
                lobj���.����֤�� = lstrCN
                lcolInfo.Add lobj���
                
                '����֤���Զ�������
                lstrCN = Format(Val(lstrCN) + 1, String(Len(lstrCN), "0"))
            End If
        Next
        
        If lcolInfo.Count = 0 Then
            MsgBox "��ѡ��Ҫ��ӡ�������Ա�������ϴ򹴣���", vbOKOnly + vbExclamation, "ϵͳ��ʾ"
            Exit Sub
        End If
        
        pobj������.sub��ӡ����֤ lcolInfo
        
        'ˢ�½��档
        subRefresh
        
    Case "delete"
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
        
    Case "dl" '����
        frm�������.Show 1, Me
    Case "exit"
        End
    End Select
End Sub

Private Sub ctxtNum_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    Dim llngNum As Long
    On Error GoTo errHandler
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
errHandler:
End Sub

Private Sub Form_Load()
    
    '��ȡ������ܵ�δ��ӡ����¼��
    mstr������ڴ� = Format(DateAdd("d", 1 - DatePart("w", Now, vbMonday), Now) - 7, "yyyy-mm-dd")
    mstr������ڵ� = Format(Now, "yyyy-mm-dd")
    
    subRefresh
    
End Sub

'���ܣ����ݲ�ѯ������ʾ��ѯ�����
Private Sub subRefresh()
    
    Dim lstr״̬���� As String
    Dim i As Long
    lstr״̬���� = ""
    If cchkType(0).Value = 1 Or cchkType(0).Value = 1 Then
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
        
    Set mobjRec = pobj������.func��������ѯ(mstrϵͳ���, mstr����, mstr��λ, mstr������ڴ�, mstr������ڵ�, mstr����, lstr״̬����)
    
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
        mobjRec.movenext
    Next
    cgrdMain.ColWidth(1) = 1000
    
    '����ϵͳ��š�
    cgrdMain.ColHidden(0) = True
    
    clblInfo.Caption = "������" & cgrdMain.Rows - 1

End Sub

Private Sub Text1_Change()

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    
End Sub
