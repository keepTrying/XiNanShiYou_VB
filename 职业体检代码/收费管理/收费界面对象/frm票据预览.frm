VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmƱ��Ԥ�� 
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
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer91 
      Height          =   9135
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   12375
      lastProp        =   500
      _cx             =   21828
      _cy             =   16113
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
   Begin VB.CommandButton ccmdOk 
      Caption         =   "ȷ  ��"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin MSComCtl2.DTPicker cdtpDate 
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      Format          =   25624577
      CurrentDate     =   39654
   End
   Begin VB.CommandButton Ccmd�˳� 
      Caption         =   "����"
      Height          =   360
      Left            =   4440
      TabIndex        =   0
      Top             =   120
      Width           =   855
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
Attribute VB_Name = "frmƱ��Ԥ��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Explicit

Private Sub ccmdOk_Click()
    Dim i As Integer, j As Integer
    Dim lstr���ϵ���Ʊ�� As String, lstr����Ʊ�ݺ� As String
    Dim lobjRec As Object
    Dim lcur�ֽ�ϼ� As Currency, lcur���˺ϼ� As Currency
    
    Dim capp As New CRAXDRT.Application
    Dim cr As CRAXDRT.Report
    Dim Item As CRAXDRT.FormulaFieldDefinition
    
    Set cr = capp.OpenReport(App.Path & "\�շ�Ա�����ձ���.rpt")
    
    '��ʼ��Ʊ�ݸ�ʽ�ļ�
    With cr
        Set lobjRec = dafuncGetData("�շѹ���_�շ�Ա�����ձ��� '" & cdtpDate.Value & "','" & um�û���� & "'")
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
                If Not IsNull(lobjRec(1)) Then lcur�ֽ�ϼ� = lcur�ֽ�ϼ� + CCur(lobjRec(1))
                If Not IsNull(lobjRec(2)) Then lcur���˺ϼ� = lcur���˺ϼ� + CCur(lobjRec(2))
                i = i + 1
                lobjRec.MoveNext
            Loop
            Set Item = cr.FormulaFields.GetItemByName("�ֽ�ϼ�")
            Item.Text = "'" & Format(lcur�ֽ�ϼ�, "#####0.00") & "'"
            Set Item = cr.FormulaFields.GetItemByName("ת�˺ϼ�")
            Item.Text = "'" & Format(lcur���˺ϼ�, "#####0.00") & "'"
            Set Item = cr.FormulaFields.GetItemByName("�շ�Ա")
            Item.Text = "'" & um�û��� & "'"
            Set Item = cr.FormulaFields.GetItemByName("��ӡ����")
            Item.Text = "'" & cdtpDate.Value & "'"
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
            Do While Not lobjRec.EOF      '������Ʊ��
                lstr���ϵ���Ʊ�� = lstr���ϵ���Ʊ�� & lobjRec(1) & "��"
                lstr����Ʊ�ݺ� = lstr����Ʊ�ݺ� & lobjRec(2) & "��"
                lobjRec.MoveNext
            Loop
            If lstr���ϵ���Ʊ�� <> "" Then lstr���ϵ���Ʊ�� = Left(lstr���ϵ���Ʊ��, Len(lstr���ϵ���Ʊ��) - 1)
            If lstr����Ʊ�ݺ� <> "" Then lstr����Ʊ�ݺ� = Left(lstr����Ʊ�ݺ�, Len(lstr����Ʊ�ݺ�) - 1)
            Set Item = cr.FormulaFields.GetItemByName("���ϵ���Ʊ��")
            Item.Text = "'" & lstr���ϵ���Ʊ�� & "'"
            Set Item = cr.FormulaFields.GetItemByName("����Ʊ�ݺ�")
            Item.Text = "'" & lstr����Ʊ�ݺ� & "'"
        End If
    End With
    CRViewer91.ReportSource = cr
    CRViewer91.ViewReport
End Sub
Private Sub Ccmd�˳�_Click()
    Unload Me
End Sub

Private Sub Form_Resize()
    On Error Resume Next
'    CrystalReport1.WindowLeft = 0
'    CrystalReport1.WindowTop = 0
'    CrystalReport1.WindowWidth = Me.ScaleWidth
'    CrystalReport1.WindowHeight = Me.ScaleHeight
End Sub
