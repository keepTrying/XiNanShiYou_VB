VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmҵ������ 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "����֤ҵ������"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   8760
   Icon            =   "frmҵ������.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   8760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox ctxtFlowNo 
      Height          =   375
      Left            =   2160
      TabIndex        =   12
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox ctxt�������� 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   2160
      TabIndex        =   9
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton ccmdPrintSetting 
      Caption         =   "��ӡ��ʽ����"
      Height          =   495
      Left            =   480
      TabIndex        =   8
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CheckBox cchk���� 
      Caption         =   "�Ǽ�ʱҪ����"
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Caption         =   "��������"
      Height          =   5295
      Left            =   3960
      TabIndex        =   3
      Top             =   960
      Width           =   4455
      Begin VB.ListBox clstDisease 
         Height          =   4050
         ItemData        =   "frmҵ������.frx":0442
         Left            =   240
         List            =   "frmҵ������.frx":0449
         Style           =   1  'Checkbox
         TabIndex        =   4
         Top             =   840
         Width           =   3495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "��Ҫ����Ĳ��֣�"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "����֤��ʽ����"
      Height          =   1935
      Left            =   360
      TabIndex        =   1
      Top             =   3960
      Visible         =   0   'False
      Width           =   3135
      Begin VB.CheckBox cchk�ֹ� 
         Caption         =   "�ֹ����뽡��֤��"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   2055
      End
      Begin VB.CheckBox cchk���� 
         Caption         =   "����֤������"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Value           =   1  'Checked
         Width           =   1455
      End
   End
   Begin MSComctlLib.Toolbar C������ 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8760
      _ExtentX        =   15452
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
   Begin VB.Label Label3 
      Caption         =   "���ý���֤��ˮ�ţ�"
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "���õ������룺"
      Height          =   375
      Left            =   480
      TabIndex        =   10
      Top             =   2520
      Width           =   1455
   End
End
Attribute VB_Name = "frmҵ������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public pblnInUse As Boolean

Private WithEvents mobjGUI As cls����ͨ�ö��� '���������õĽ���ͨ�ö�
Attribute mobjGUI.VB_VarHelpID = -1

Private Sub cchk����_Click()
    On Error Resume Next
    If cchk����.Value = 1 Then
        cchk�ֹ�.Value = 0
        cchk�ֹ�.Enabled = False
    Else
        cchk�ֹ�.Enabled = True
    End If
End Sub

Private Sub ccmdPrintSetting_Click()
    frm��ӡ��ʽ����.Show 1
End Sub

Private Sub Form_Load()
    Dim lcolInfo As Collection
    Dim lstrTemp As String
    Dim i As Long
    Dim lobjRec As Object
    On Error GoTo errhandler
    
    If pblnInUse Then Exit Sub
    pblnInUse = True
    
    '��ʼ����������
    Dim lcol��������ť As New Collection '�������ϵİ�ť��ʼ�����ϡ�
    Set mobjGUI = New cls����ͨ�ö���
    Set mobjGUI.Form = Me
    Set mobjGUI.C������ = C������
    lcol��������ť.Add "����"
    lcol��������ť.Add "|"
    lcol��������ť.Add "�˳�"
    mobjGUI.subInitialize lcol��������ť, ""
    
    
    '���벡�֡�
    lstrTemp = pobj������.ҵ������("��������")
    '��ȡ������֡�
    Set lcolInfo = pobj����.������ֵ("�������", True)
    clstDisease.Clear
    For i = 1 To lcolInfo.Count
        clstDisease.AddItem lcolInfo(i)
        If InStr("," + lstrTemp + ",", "," & lcolInfo(i) & ",") > 0 Then
            clstDisease.Selected(clstDisease.ListCount - 1) = True
        End If
    Next
        
'    lstrTemp = pobj������.ҵ������("����֤������")
'    If lstrTemp = "��" Then
'        cchk����.Value = 1
'    Else
'        cchk����.Value = 0
'        lstrTemp = pobj������.ҵ������("�ֹ����뽡��֤��")
'        If lstrTemp = "��" Then
'            cchk�ֹ�.Value = 1
'        Else
'            cchk�ֹ�.Value = 0
'        End If
'
'
'    End If
    
    If pobj������.ҵ������("�Ƿ�����") = "��" Then
        cchk����.Value = 1
    Else
        cchk����.Value = 0
    End If
    Set lobjRec = dafuncGetData("select top 1 �غ���ַ from ����֤_ҵ�����ñ�")
    ctxt��������.Text = lobjRec(0)
            '��佡��֤��ˮ��
    Set lobjRec = dafuncGetData("select ��ǰֵ from ϵͳ����_ϵͳ������ɼ�¼�� where ҵ������='����֤����' and �������='����֤���'")
    If lobjRec.RecordCount > 0 Then
        ctxtFlowNo.Text = lobjRec(0)
    End If

    
    Exit Sub
errhandler:
    sfsub������ "����֤������", "frmҵ������", "Form_Load", Err.Number, Err.Description, False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set mobjGUI = Nothing
    pblnInUse = False
End Sub

Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    Dim i As Long
    Dim lstrTemp As String
    Dim lobjRec As Object
    On Error GoTo errhandler
    Select Case Operate
    Case "����"
        pobj������.ҵ������("����֤������") = IIf(cchk����.Value = 1, "��", "��")
        pobj������.ҵ������("�ֹ����뽡��֤��") = IIf(cchk�ֹ�.Value = 1, "��", "��")
        
        lstrTemp = ""
        For i = 0 To clstDisease.ListCount - 1
            If clstDisease.Selected(i) Then
                lstrTemp = lstrTemp & clstDisease.List(i) & ","
            End If
        Next
        If lstrTemp <> "" Then lstrTemp = Left(lstrTemp, Len(lstrTemp) - 1)
        '��������
        pobj������.ҵ������("��������") = lstrTemp
        pobj������.ҵ������("�Ƿ�����") = IIf(cchk����.Value = 1, "��", "��")
        dafuncGetData ("update ����֤_ҵ�����ñ� set �غ���ַ='" & ctxt��������.Text & "'")
        '����֤���
        Set lobjRec = dafuncGetData("select * from ϵͳ����_ϵͳ������ɼ�¼�� where ҵ������='����֤����' and �������='����֤���'")
        If lobjRec.RecordCount > 0 Then
            dafuncGetData ("update ϵͳ����_ϵͳ������ɼ�¼�� set ��ǰֵ='" & IIf(Trim(ctxtFlowNo.Text) = "", 0, Trim(ctxtFlowNo.Text)) & "' where ҵ������='����֤����' and �������='����֤���'")
        Else
            dafuncGetData ("Insert Into ϵͳ����_ϵͳ������ɼ�¼��(ҵ������,�������,��ǰֵ,��������,����,���ֵ,�Ƿ����ر�,��ǰ���) values " _
                            & "('����֤����','����֤���','" & IIf(Trim(ctxtFlowNo.Text) = "", 0, Trim(ctxtFlowNo.Text)) & "','C','6','999999','��','" & Year(Date) & "')")
        End If
        Cancel = True
        
    End Select
    
    Exit Sub
errhandler:
    sfsub������ "����֤������", "frmҵ������", "mobjGUI_BeforeOperate", Err.Number, Err.Description, False
    
End Sub
