VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm���� 
   Caption         =   "�շ�Ա�����ձ���"
   ClientHeight    =   9945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12750
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9945
   ScaleWidth      =   12750
   StartUpPosition =   2  '��Ļ����
   Begin VB.ComboBox clstName 
      Height          =   300
      Left            =   5040
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer91 
      Height          =   9135
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   12375
      lastProp        =   600
      _cx             =   21828
      _cy             =   16113
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   0   'False
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   0   'False
      EnableRefreshButton=   0   'False
      EnableDrillDown =   0   'False
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   0   'False
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
   End
   Begin VB.CommandButton ccmdOk 
      Caption         =   "ȷ  ��"
      Height          =   375
      Left            =   6840
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin MSComCtl2.DTPicker cdtpBeginDate 
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      Format          =   51970049
      CurrentDate     =   39654
   End
   Begin VB.CommandButton Ccmd�˳� 
      Caption         =   "����"
      Height          =   360
      Left            =   8040
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin MSComCtl2.DTPicker cdtpEndDate 
      Height          =   255
      Left            =   3000
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      Format          =   51970049
      CurrentDate     =   39654
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "��"
      Height          =   180
      Left            =   2760
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label clblName 
      AutoSize        =   -1  'True
      Caption         =   "�շ�Ա��"
      Height          =   180
      Left            =   4320
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�շ����ڣ�"
      Height          =   180
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   900
   End
End
Attribute VB_Name = "frm����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub ccmdOK_Click()
    Dim i As Integer, j As Integer
    Dim lstr���ϵ���Ʊ�� As String, lstr����Ʊ�ݺ� As String
    Dim lobjRec As Object
    Dim lcur�ֽ�ϼ� As Currency, lcur���˺ϼ� As Currency, lcur�˷Ѻϼ� As Currency
    
    Dim capp As New CRAXDRT.Application
    Dim cr As CRAXDRT.Report
    Dim Item As CRAXDRT.FormulaFieldDefinition
    
    Set cr = capp.OpenReport(App.Path & "\�շ�Ա�����ձ���.rpt")
    
    On Error GoTo errhandle
    
    '��ʼ��Ʊ�ݸ�ʽ�ļ�
    With cr
        If Not clstName.Visible Then
            Set lobjRec = dafuncGetData("�շѹ���_�շ�Ա�����ձ��� '" & cdtpBeginDate.Value & "','" & um�û���� & "'")
        Else
            Set lobjRec = dafuncGetData("�շѹ���_�շ�Ա�����ձ��� '" & cdtpBeginDate.Value & "','" & Left(clstName.Text, InStr(clstName.Text, " ") - 1) & "'")
        End If
        i = 1
        j = 1
        If lobjRec.RecordCount Then
            Do While lobjRec(0) <> "�Ŷ�" And lobjRec(0) <> "����"
                Set Item = cr.FormulaFields.GetItemByName("��Ŀ" & i)
                Item.Text = "'" & lobjRec(0) & "'"
                Set Item = cr.FormulaFields.GetItemByName("�ֽ�" & i)
                Item.Text = "'" & IIf(IsNull(lobjRec(1)), "0.00", lobjRec(1)) & "'"
                Set Item = cr.FormulaFields.GetItemByName("ת��" & i)
                Item.Text = "'" & IIf(IsNull(lobjRec(2)), "0.00", lobjRec(2)) & "'"
                Set Item = cr.FormulaFields.GetItemByName("�˷�" & i)
                Item.Text = "'" & IIf(IsNull(lobjRec(3)), "0.00", lobjRec(3)) & "'"
                If Not IsNull(lobjRec(1)) Then lcur�ֽ�ϼ� = lcur�ֽ�ϼ� + CCur(lobjRec(1))
                If Not IsNull(lobjRec(2)) Then lcur���˺ϼ� = lcur���˺ϼ� + CCur(lobjRec(2))
                If Not IsNull(lobjRec(3)) Then lcur�˷Ѻϼ� = lcur�˷Ѻϼ� + CCur(lobjRec(3))
                i = i + 1
                lobjRec.MoveNext
            Loop
            Set Item = cr.FormulaFields.GetItemByName("�ֽ�ϼ�")
            Item.Text = "'" & Format(lcur�ֽ�ϼ�, "#####0.00") & "'"
            Set Item = cr.FormulaFields.GetItemByName("ת�˺ϼ�")
            Item.Text = "'" & Format(lcur���˺ϼ�, "#####0.00") & "'"
            Set Item = cr.FormulaFields.GetItemByName("�˷Ѻϼ�")
            Item.Text = "'" & Format(lcur�˷Ѻϼ�, "#####0.00") & "'"
            Set Item = cr.FormulaFields.GetItemByName("�շ�Ա")
            If Not clstName.Visible Then
                Item.Text = "'" & um�û��� & "'"
            Else
                Item.Text = "'" & clstName.Text & "'"
            End If
            Set Item = cr.FormulaFields.GetItemByName("��ӡ����")
            Item.Text = "'" & cdtpBeginDate.Value & "��" & cdtpEndDate.Value & "'"
            i = 1
            Do While lobjRec(0) = "�Ŷ�"
                Set Item = cr.FormulaFields.GetItemByName("�Ŷ�����" & i)
                Item.Text = "'" & lobjRec(1) & "'"
                Set Item = cr.FormulaFields.GetItemByName("�Ŷ�Ʊ��" & i)
                Item.Text = "'" & lobjRec(2) & "'"
                i = i + 1
                lobjRec.MoveNext
            Loop
            If lobjRec(0) = "����" Then
                Set Item = cr.FormulaFields.GetItemByName("����")
                Item.Text = "'" & lobjRec(1) & "'"
                lobjRec.MoveNext
            End If
            i = 0
            Do While Not lobjRec.EOF      '������Ʊ��
                Do While lobjRec(0) = "����Ʊ��"
                    'lstr���ϵ���Ʊ�� = lstr���ϵ���Ʊ�� & lobjRec(1) & "��"
                    lstr����Ʊ�ݺ� = lstr����Ʊ�ݺ� & lobjRec(1) & "��"
                    i = i + 1
                    lobjRec.MoveNext
                    If lobjRec.EOF Then Exit Do
                Loop
                Exit Do
            Loop
            'If lstr���ϵ���Ʊ�� <> "" Then lstr���ϵ���Ʊ�� = Left(lstr���ϵ���Ʊ��, Len(lstr���ϵ���Ʊ��) - 1)
            If lstr����Ʊ�ݺ� <> "" Then lstr����Ʊ�ݺ� = Left(lstr����Ʊ�ݺ�, Len(lstr����Ʊ�ݺ�) - 1)
            Set Item = cr.FormulaFields.GetItemByName("��������")
            Item.Text = "'" & IIf(i = 0, "", i) & "'"
            Set Item = cr.FormulaFields.GetItemByName("����Ʊ�ݺ�")
            Item.Text = "'" & lstr����Ʊ�ݺ� & "'"
            
            lstr����Ʊ�ݺ� = ""
            i = 0
            Do While Not lobjRec.EOF      '���˷�Ʊ��
                'lstr���ϵ���Ʊ�� = lstr���ϵ���Ʊ�� & lobjRec(1) & "��"
                lstr����Ʊ�ݺ� = lstr����Ʊ�ݺ� & lobjRec(1) & "��"
                i = i + 1
                lobjRec.MoveNext
            Loop
            'If lstr���ϵ���Ʊ�� <> "" Then lstr���ϵ���Ʊ�� = Left(lstr���ϵ���Ʊ��, Len(lstr���ϵ���Ʊ��) - 1)
            If lstr����Ʊ�ݺ� <> "" Then lstr����Ʊ�ݺ� = Left(lstr����Ʊ�ݺ�, Len(lstr����Ʊ�ݺ�) - 1)
            Set Item = cr.FormulaFields.GetItemByName("�˷�����")
            Item.Text = "'" & IIf(i = 0, "", i) & "'"
            Set Item = cr.FormulaFields.GetItemByName("�˷�Ʊ�ݺ�")
            Item.Text = "'" & lstr����Ʊ�ݺ� & "'"
        End If
    End With
    CRViewer91.ReportSource = cr
    CRViewer91.ViewReport
    Exit Sub
errhandle:
    MsgBox "ͳ�Ʊ���ʱ���ִ���" & Error, vbInformation, "ϵͳ��ʾ"
End Sub
Private Sub Ccmd�˳�_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    cdtpBeginDate.Value = Format(Date, "yyyy-mm-dd")
    cdtpEndDate.Value = Format(Date, "yyyy-mm-dd")
    
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
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
'    CrystalReport1.WindowLeft = 0
'    CrystalReport1.WindowTop = 0
'    CrystalReport1.WindowWidth = Me.ScaleWidth
'    CrystalReport1.WindowHeight = Me.ScaleHeight
End Sub
