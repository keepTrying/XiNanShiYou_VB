VERSION 5.00
Begin VB.Form frmConclusion 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "�������������"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   8880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame2 
      Caption         =   "ʹ��˵����"
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Left            =   120
      TabIndex        =   9
      Top             =   4560
      Width           =   5895
      Begin VB.Label Label1 
         Caption         =   $"frmConclusion.frx":0000
         Height          =   840
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   5520
      End
   End
   Begin VB.CommandButton ccmdAdd 
      Caption         =   "���"
      Height          =   375
      Left            =   6360
      TabIndex        =   8
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton ccmdDel 
      Caption         =   "ɾ��"
      Height          =   375
      Left            =   6360
      TabIndex        =   3
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton ccmdSure 
      Caption         =   "ȷ��"
      Height          =   375
      Left            =   7680
      TabIndex        =   2
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton ccmdCancel 
      Caption         =   "�˳�"
      Height          =   375
      Left            =   7680
      TabIndex        =   1
      Top             =   5160
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "�޸ġ�����������"
      ForeColor       =   &H000080FF&
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8655
      Begin VB.Frame Frame4 
         Caption         =   "���������"
         Height          =   1695
         Left            =   120
         TabIndex        =   6
         Top             =   2640
         Width           =   8415
         Begin VB.TextBox ctxtConclusion��ѡ���� 
            Height          =   1455
            Left            =   0
            ScrollBars      =   2  'Vertical
            TabIndex        =   7
            Top             =   240
            Width           =   8295
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "����ģ��"
         Height          =   2295
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   8415
         Begin VB.ListBox clstConclusion��ѡ���� 
            Height          =   2040
            ItemData        =   "frmConclusion.frx":00C8
            Left            =   120
            List            =   "frmConclusion.frx":00CA
            TabIndex        =   5
            Top             =   240
            Width           =   8175
         End
      End
   End
End
Attribute VB_Name = "frmConclusion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'���壺ְҵ��������ģ��¼��
'���ܣ���ģ���ÿ�����ҵĽ���¼��������
'���ߣ���¶
'ʱ�䣺2012-05-03
'��ע������

Option Explicit
Public lobj���ÿ��� As String
Public lobj���� As String   '�ж��ĸ����ҵ���
Public lobj���� As String   '�ж��ǵ��õ��ĸ�����
Public lobj���ұ�� As String
Public lobjҽ����� As String
Public lobjʱ�� As String
Attribute lobjʱ��.VB_VarHelpID = -1

'������ݿ���û�е�������ģ��
Private Sub ccmdAdd_Click()
    Dim lobj��ӽ��� As Object
    Dim lobjRec��� As Object
    If Trim(ctxtConclusion��ѡ����.Text & "") <> "" Then
        If LCase(clstConclusion��ѡ����) <> LCase(ctxtConclusion��ѡ����) Then
            clstConclusion��ѡ����.AddItem (Trim(ctxtConclusion��ѡ����.Text))
            lobj���� = ctxtConclusion��ѡ����.Text
            ctxtConclusion��ѡ����.Text = ""
            Set lobj��ӽ��� = CreateObject("ְҵ������.clsMedicalExaminer")
            Set lobjRec��� = lobj��ӽ���.func����ض����ҵĽ���ģ��(lobj���ұ��, lobj����, lobj����, lobjҽ�����, lobjʱ��)
        End If
    End If
End Sub

'�˳�������ģ�崰��
Private Sub ccmdCancel_Click()
    Unload frmConclusion
    Set frmConclusion = Nothing
End Sub

'��ȡ���ݿ������е�������ģ��
Private Sub Form_Load()
    '�����ۿ��еĽ��ۿ���Ϊģ�����ѡ��Ҳ�ɵ���֮�����·����ۿ��н����޸ĺ�ɾ����Ҳ��������û�еĽ���ģ�壬�ڳɹ�ѡ�������ۺ��ȵ��ȷ����ť���ٵ���˳���ť���Ϳ����˳��ô��壬�ɲ����������塣
    Label1.Caption = "�����ۿ��еĽ��ۿ���Ϊģ�����ѡ��Ҳ�ɵ���֮�����·����ۿ��н����޸ĺ�ɾ����Ҳ��������û�еĽ���ģ�壬�ڳɹ�ѡ�������ۺ��ȵ��ȷ����ť���ٵ���˳���ť���Ϳ����˳��ô��壬�ɲ����������塣"
    Dim lobjģ�� As Object
    Dim lobjRec As Object

    Set lobjģ�� = CreateObject("ְҵ������.clsMedicalExaminer") '��ȡ������ģ��
    Set lobjRec = lobjģ��.func��ȡ�ض����ҵĽ���ģ��(lobj����)
    While Not lobjRec.EOF
        clstConclusion��ѡ����.AddItem lobjRec("����ģ��")
        lobjRec.MoveNext
    Wend
End Sub

'�������۽��е���ɾ��
Private Sub ccmdDel_Click()
    Dim i As Integer
    Dim lobjɾ������ As Object
    Dim lobjRecɾ�� As Object
    If clstConclusion��ѡ����.SelCount > 0 Then
        For i = clstConclusion��ѡ����.ListCount - 1 To 0 Step -1
            If clstConclusion��ѡ����.Selected(i) Then
            
                 lobj���� = clstConclusion��ѡ����.List(i)
                 clstConclusion��ѡ����.RemoveItem (i)
                 '2015.7.30 modify by lanchao
'                lobj���� = ctxtConclusion��ѡ����.Text
                ctxtConclusion��ѡ����.Text = ""
                Set lobjɾ������ = CreateObject("ְҵ������.clsMedicalExaminer")
                Set lobjRecɾ�� = lobjɾ������.funcɾ���ض����ҵĽ���ģ��(lobj����, lobj����)
            End If
        Next
    End If
End Sub

'�������۽���ȷ�������ݵ��������ҵĽ���¼�����
Private Sub ccmdSure_Click()
    If lobj���ÿ��� = "frmResultInput_Assay" Then
        If frmResultInput_Assay.ctxtResult = "" Then
        '2015.7.30 modify by lanchao
          frmResultInput_Assay.ctxtResult = ctxtConclusion��ѡ����.Text
        Else
          frmResultInput_Assay.ctxtResult = frmResultInput_Assay.ctxtResult + "," + ctxtConclusion��ѡ����.Text
        End If
        
    Else
       If frmResultInput_Routine.ctxtConclun = "" Then
        frmResultInput_Routine.ctxtConclun = ctxtConclusion��ѡ����.Text
        Else
        frmResultInput_Routine.ctxtConclun = frmResultInput_Routine.ctxtConclun + "," + ctxtConclusion��ѡ����.Text
        End If
        
    End If
    Unload Me
'    If frmBloodRoutine_ResultInput.mblnInUse = True Then
'        frmBloodRoutine_ResultInput.ctxtConclun = ctxtConclusion��ѡ����.Text
'    ElseIf frmHEENTnew_ResultInput.mblnInUse = True Then
'        frmHEENTnew_ResultInput.ctxtConclun = ctxtConclusion��ѡ����.Text
'    ElseIf frmBUS_ResultInput.mblnInUse = True Then
'        frmBUS_ResultInput.ctxtResult = ctxtConclusion��ѡ����.Text
'    ElseIf frmChromosome_ResultInput.mblnInUse = True Then
'        frmChromosome_ResultInput.ctxtConclun = ctxtConclusion��ѡ����.Text
'    ElseIf frmECG_ResultInput.mblnInUse = True Then
'        frmECG_ResultInput.ctxtConclun = ctxtConclusion��ѡ����.Text
'    ElseIf frmElectroaudiometer_ResultInput.mblnInUse = True Then
'        frmElectroaudiometer_ResultInput.ctxtConclun = ctxtConclusion��ѡ����.Text
'    ElseIf FrmInMedi_ResultInput.mblnInUse = True Then
'        FrmInMedi_ResultInput.ctxtConclun = ctxtConclusion��ѡ����.Text
'    ElseIf frmLiverFunc_ResultInput.mblnInUse = True Then
'        frmLiverFunc_ResultInput.ctxtConclun = ctxtConclusion��ѡ����.Text
'    ElseIf frmPFT_ResultInput.mblnInUse = True Then
'        frmPFT_ResultInput.ctxtConclun = ctxtConclusion��ѡ����.Text
'    ElseIf frmSurgery_ResultInput.mblnInUse = True Then
'        frmSurgery_ResultInput.ctxtConclun = ctxtConclusion��ѡ����.Text
'    ElseIf frmURT_ResultInput.mblnInUse = True Then
'        frmURT_ResultInput.ctxtConclun = ctxtConclusion��ѡ����.Text
'    ElseIf frmXRay_ResultInput.mblnInUse = True Then
'        frmXRay_ResultInput.ctxtResult = ctxtConclusion��ѡ����.Text
'    End If
End Sub

'�������۵�������ѡ�� ���޸����еĽ���ģ��
Private Sub clstConclusion��ѡ����_Click()
     '2015.7.31 modify by lanchao
     If ctxtConclusion��ѡ����.Text = "" Then
        ctxtConclusion��ѡ����.Text = clstConclusion��ѡ����.Text
        Else
        ctxtConclusion��ѡ����.Text = ctxtConclusion��ѡ����.Text + "," + clstConclusion��ѡ����.Text
        End If
     
End Sub

