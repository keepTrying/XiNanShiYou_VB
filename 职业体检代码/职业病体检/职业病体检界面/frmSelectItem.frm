VERSION 5.00
Object = "{8099FCC2-0A81-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelectItem 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ѡ�񸴲���Ŀ���շ���Ŀ"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   4110
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   120
      TabIndex        =   11
      Top             =   4680
      Width           =   3975
   End
   Begin VSFlex6Ctl.vsFlexGrid cgrd�շ���Ŀ 
      Height          =   3615
      Left            =   4080
      TabIndex        =   5
      Top             =   360
      Visible         =   0   'False
      Width           =   3855
      _cx             =   59120272
      _cy             =   59119848
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
      AllowBigSelection=   0   'False
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
      FormatString    =   "��Ŀ���     |��Ŀ����            |����"
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
      Begin VB.Label clblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "��ҵ�����á��������˲��շѣ����Բ���Ҫѡ���շ���Ŀ��"
         ForeColor       =   &H00800000&
         Height          =   2340
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Visible         =   0   'False
         Width           =   3420
      End
   End
   Begin VB.CommandButton ccmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton ccmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   375
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4920
      Width           =   1215
   End
   Begin MSComctlLib.TreeView ctrwItem 
      Height          =   3645
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   6429
      _Version        =   393217
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Checkboxes      =   -1  'True
      Appearance      =   1
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
   Begin VB.Label clblMsg 
      BackStyle       =   0  'Transparent
      Caption         =   "˵���������г����ǡ����ϵĸ�����Ŀ�����㷢����Ŀ����������롰�������á�������������Ҫ����Ŀ���������"
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   4080
      Width           =   3855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(Ԫ)"
      Height          =   180
      Index           =   3
      Left            =   6720
      TabIndex        =   9
      Top             =   4080
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label clblTotal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   5400
      TabIndex        =   8
      Top             =   4080
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�շ��ܶ"
      Height          =   180
      Index           =   2
      Left            =   4440
      TabIndex        =   7
      Top             =   4080
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ѡ���շ���Ŀ��"
      Height          =   180
      Index           =   1
      Left            =   4080
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ѡ�񸴲���Ŀ��"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1260
   End
End
Attribute VB_Name = "frmSelectItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'���ƣ������Ŀѡ��
'���ܣ�ѡ�������Ŀ
'������
'���ߣ�������
'ʱ�䣺2012-03
Option Explicit

'2012-08-22 �ڵ�� ��
'��ͣX���롣���������ѡ��ʱ����ֹ����ٶȹ��죬���¿ؼ���Ӧ��������
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'2012-08-22 �ڵ�� ��

Public pstr�������� As String
Public pcol������Ŀ As Collection  '���ؽ��[���룬����],key:���롣
Public pcol�շ���Ŀ As Collection  '���ؽ��[�շ���Ŀ���,����]key:��š�
Public pblnOk As Boolean           '�Ƿ�ȷ�����ء�

Private Sub ccmdCancel_Click()
    On Error Resume Next
    pblnOk = False
    Set pcol�շ���Ŀ = New Collection
    Unload Me
End Sub

Public Sub ccmdOk_Click()
    Dim lobjNode As Node
    Dim lcolInfo As Collection
    Dim lcolItem As Collection
    Dim lstrItem As String
    
    Dim i As Long
    
    On Error GoTo errHandler
    '��ȡѡ�����Ŀ��
    Set lcolInfo = New Collection
    For Each lobjNode In ctrwItem.Nodes
        If lobjNode.Checked And Not lobjNode.Parent Is Nothing Then
            lstrItem = Right(lobjNode.Key, Len(lobjNode.Key) - 1)
            Set lcolItem = New Collection
            lcolItem.Add lstrItem, "����"
            lcolItem.Add Right(lobjNode.Text, Len(lobjNode.Text) - InStr(lobjNode.Text, " ")), "����"
            lcolInfo.Add lcolItem, lstrItem
        End If
    Next
    Set lobjNode = Nothing
'    If lcolInfo.Count = 0 Then
'        sffuncMsg "����ѡ�������Ŀ��", sf����
'        Set lcolInfo = Nothing
'        Exit Sub
'    End If
    
'    '�����    2016-6-15 by Ĳ��
'    Dim par����ϵͳ��� As String
'    Dim obj As Object
'    par����ϵͳ��� = frmFinalConclusion.cgrdInfo.TextMatrix(frmFinalConclusion.cgrdInfo.Row, frmFinalConclusion.cgrdInfo.ColIndex("ϵͳ���")) & "F"
'    Set obj = dafuncGetData("select * from ְҵ�����_������Ŀ�� where ϵͳ���='" & par����ϵͳ��� & "'")
'        If obj.RecordCount > 0 Then
'            dafuncGetData ("delete from ְҵ�����_������Ŀ�� where ϵͳ���='" & par����ϵͳ��� & "'")
'        End If
'    For Each lobjNode In ctrwItem.Nodes
'        If lobjNode.Checked And Not lobjNode.Parent Is Nothing Then
'            lstrItem = Right(lobjNode.Key, Len(lobjNode.Key) - 1)
'            dafuncGetData ("insert into ְҵ�����_������Ŀ�� values ('" & par����ϵͳ��� & "','" & lstrItem & "')")
'        End If
'    Next
    
    
    '��ȡ�շ���Ŀ��
    'Set pcol�շ���Ŀ = New Collection
    'If pobjҵ�����.ҵ������("�Ƿ��շ�") = "��" Then
    '    For i = 1 To cgrd�շ���Ŀ.Rows - 1
    '        If cgrd�շ���Ŀ.Cell(flexcpChecked, i, 0, i, 0) = flexChecked Then
    '            Set lcolItem = New Collection
    '            lcolItem.Add cgrd�շ���Ŀ.TextMatrix(i, 0), "�շ���Ŀ���"
    '            lcolItem.Add cgrd�շ���Ŀ.ValueMatrix(i, 2), "����"
    '            pcol�շ���Ŀ.Add lcolItem, lcolItem("�շ���Ŀ���")
    '        End If
    '    Next
    '    If pcol�շ���Ŀ.Count = 0 Then
    '        If MsgBox("ҵ��������Ҫ�շѣ���ȷ�ϸ���Ա���շ���" & Chr(13) & Chr(10) & "���޷�ѡ���շ���Ŀ������롰ҵ�����á����������ò��շѣ��򰴽����ϵ���ɫ��ʾ��Ϣ���д���", vbYesNo + vbQuestion + vbDefaultButton2, "ϵͳѯ��") = vbNo Then
    '            Exit Sub
    '        End If
    '    End If
    'End If
    
    Set pcol������Ŀ = lcolInfo
    pblnOk = True
    Unload Me

    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����沿��", "frmSelectItem", "ccmdOk_Click", 6666, lstrError, False

End Sub

Private Sub cgrd�շ���Ŀ_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    Dim dblTotal As Double
    
    On Error Resume Next
    '�����ܽ��
    For i = 1 To cgrd�շ���Ŀ.rows - 1
        If cgrd�շ���Ŀ.Cell(flexcpChecked, i, 0) = flexChecked Then
            dblTotal = Format(dblTotal + cgrd�շ���Ŀ.ValueMatrix(i, 2), "0.00")
        End If
    Next
    
    clblTotal.Caption = dblTotal
End Sub

Private Sub cgrd�շ���Ŀ_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    On Error Resume Next
    If Row > 0 And Col = 0 Then
    Else
        Cancel = True
    End If
End Sub
'���ܣ����Ը��ڵ���в������Զ����ӽڵ���в�������֮��
Private Sub ctrwItem_NodeCheck(ByVal Node As MSComctlLib.Node)
    On Error Resume Next
    Dim lobjNode As Node
    Dim i As Long
    
    If Node.Parent Is Nothing Then
        '��ǰ�������Ǹ��ڵ㡣
        '�Զ�ѡ��(��ѡ��)�����ӽڵ㡣
        If Node.Children > 0 Then
            For i = 1 To Node.Children
                Set lobjNode = ctrwItem.Nodes(Node.Index + i)
                lobjNode.Checked = Node.Checked
            Next
        End If
    Else
        '��ǰ���������ӽڵ㡣
        If Node.Checked Then
            Node.Parent.Checked = True
        Else
            For i = 1 To Node.Parent.Children
                If ctrwItem.Nodes(Node.Parent.Index + i).Checked Then
                    Exit For
                End If
            Next
            If i > Node.Parent.Children Then
                Node.Parent.Checked = False
            End If
        End If
    End If
    
    '2012-08-22 ���� ��
    '���ƿؼ���Ӧ�ٶȣ���ֹ�ٶȹ��죬�����Ŀ������󡣱�����350msһ�������û�����⣬����ǳ���Ļ�����Ȼ����ѡ��Ŀ��ܡ�
    ctrwItem.Visible = False
    Sleep (500)
    ctrwItem.Visible = True
    '2012-08-22 ���� ��
End Sub

Private Sub Form_Load()
    Dim lobj���ģ��  As Object
    Dim lobjDict As Object
    Dim lobj�����Ŀ�� As Object
    Dim lobjRec As Object
    Dim lobjItem As Object
    Dim lobjNode As Node
    Dim lcolInfo As Collection
    Dim ldblTotal As Double
    Dim i As Long
    Dim lblnȫ��ѡ�� As Boolean
    
    On Error GoTo errHandler
    'ȷ����������  2016-6-13
    pstr�������� = frmFinalConclusion.cgrdInfo.TextMatrix(frmFinalConclusion.cgrdInfo.Row, frmFinalConclusion.cgrdInfo.ColIndex("������"))
    clblMsg.Caption = "˵���������г����ǡ�" & pstr�������� & "���ϵ������Ŀ�����㷢����Ŀ����������롰�������á�������������Ҫ����Ŀ���������"
    
    '��������ģ�����
    Set lobj���ģ�� = CreateObject("ְҵ������.clsMedicalExamTemplate")
    lobj���ģ��.������ = pstr��������
    
    '��ȡ��������������Ŀ��
    Set lcolInfo = lobj���ģ��.�����Ŀ��
    
    '�����ֵ����
    Set lobjDict = CreateObject("�ֵ����.clsDictionary")
        
    lblnȫ��ѡ�� = True
    
    '���������Ŀ������
    Set lobj�����Ŀ�� = CreateObject("ְҵ������.clsTestItemSet")
    
    '��ȡ���������ࡢ�����Ŀ��
    Set lobjRec = lobjDict.Fetch("ְҵ���������ֵ�")
    '��ʾ��������ctrvItem�У����нڵ��key=������id����
    ctrwItem.Nodes.Clear
    Do While Not lobjRec.EOF
        'ͨ��"lobj�����Ŀ��"���λ�ȡ������������Ŀ��
        lobj�����Ŀ��.������ = lobjRec("InnerID")
        Set lobjItem = lobj�����Ŀ��.�����Ŀ
        If Not lobjItem.EOF Then
            ctrwItem.Nodes.Add , , "I" & lobjRec("InnerID"), lobjRec("���") & " " & Trim(lobjRec("����"))
        End If
        
        '��ʾ�����Ŀ��ctrvItem�У����нڵ��key=����,parent=�����ࣩ��
        Do While Not lobjItem.EOF
            If sffunc�жϼ��ϼ�ֵ�Ƿ����(lcolInfo, lobjItem("����")) Then
                Set lobjNode = ctrwItem.Nodes.Add("I" & lobjRec("InnerID"), tvwChild, "I" & lobjItem("����"), lobjItem("����") & " " & Trim(lobjItem("����")))
                
                'ѡ�����еĸ�����Ŀ��
                If Not pcol������Ŀ Is Nothing Then
                    If sffunc�жϼ��ϼ�ֵ�Ƿ����(pcol������Ŀ, lobjItem("����")) Then
                        lobjNode.Checked = True
                        lobjNode.Parent.Checked = True
                    Else
                        lblnȫ��ѡ�� = False
                    End If
                    
                End If
            End If
            
            lobjItem.MoveNext
        Loop

        If lobjItem.RecordCount > 0 Then
            If ctrwItem.Nodes("I" & lobjRec("InnerID")).Children = 0 Then
                ctrwItem.Nodes.Remove "I" & lobjRec("InnerID")
            End If
        End If

        lobjRec.MoveNext
    Loop
    
    '�ж��Ƿ��շѣ���Ҫ�����������շ���Ŀ��
    clblInfo.Visible = False
    '�޸ģ�2002-10-28�����Ϊ�˼ζ��������󣬲������ڲ��շ���Ϣʱ���Կ���ѡ���շ���Ŀ��
    'If pobjҵ�����.ҵ������("�Ƿ��շ�") = "��" Then
    '    If lobj���ģ��.�շѱ�׼ = "" Then
    '        clblInfo.Caption = "û������������շѱ�׼���޷�ѡ���շ���Ŀ�����Ƚ��롰�������á������������շѱ�׼���ã�"
    '        clblInfo.Visible = True
    '    Else
    '        '��ȡ���շѱ�׼���շ���Ŀ��
    '        Set lobjRec = pobjҵ�����.�շѱ�׼����Ŀ(lobj���ģ��.�շѱ�׼)
    '        cgrd�շ���Ŀ.Rows = lobjRec.RecordCount + 1
    '        i = 1
    '        ldblTotal = 0
    '        Do While Not lobjRec.EOF
    '            cgrd�շ���Ŀ.TextMatrix(i, 0) = lobjRec("�շ���Ŀ���")
    '            cgrd�շ���Ŀ.TextMatrix(i, 1) = IIf(IsNull(lobjRec("�շ���Ŀ����")), "", lobjRec("�շ���Ŀ����"))
    '            cgrd�շ���Ŀ.TextMatrix(i, 2) = IIf(IsNull(lobjRec("����")), 0, lobjRec("����"))
    '
    '            If sffunc�жϼ��ϼ�ֵ�Ƿ����(pcol�շ���Ŀ, lobjRec("�շ���Ŀ���")) Or lblnȫ��ѡ�� Then
    '                cgrd�շ���Ŀ.Cell(flexcpChecked, i, 0, i, 0) = flexChecked
    '                ldblTotal = Format(ldblTotal + IIf(IsNull(lobjRec("����")), 0, lobjRec("����")), "0.00")
    '            Else
    '                cgrd�շ���Ŀ.Cell(flexcpChecked, i, 0, i, 0) = flexUnchecked
    '            End If
    '
    '            lobjRec.MoveNext
    '            i = i + 1
    '        Loop
    '        If cgrd�շ���Ŀ.Rows > 1 Then
    '            cgrd�շ���Ŀ.Editable = True
    '            clblTotal.Caption = ldblTotal
    '        Else
    '            clblInfo.Caption = "��ǰ������շѱ�׼û���κ��շ���Ŀ���޷�ѡ���շ���Ŀ�����շѿ����Ƚ����շѱ�׼���á�"
    '            clblInfo.Visible = True
    '        End If
    '    End If
    'Else
    '    clblInfo.Caption = "��ҵ�����á��������˲��շѣ����Բ���Ҫѡ���շ���Ŀ��"
    '    clblInfo.Visible = True
    'End If
    
    Set lobjDict = Nothing

    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����沿��", "frmSelectItem", "Form_Load", 6666, lstrError, False
    Exit Sub
    Resume
End Sub

