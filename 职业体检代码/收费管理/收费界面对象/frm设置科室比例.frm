VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm���ÿ��ұ��� 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "�շ���Ŀ���ұ�������"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   9780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton ccmd���� 
      Caption         =   ">>"
      Height          =   495
      Left            =   4320
      TabIndex        =   2
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton ccmd���� 
      Caption         =   "<<"
      Height          =   495
      Left            =   4320
      TabIndex        =   1
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmd���� 
      Caption         =   "��������"
      Height          =   495
      Left            =   4320
      TabIndex        =   0
      Top             =   3360
      Width           =   1335
   End
   Begin MSComctlLib.TreeView ctr��Ŀ�� 
      Height          =   5175
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   9128
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   6
      Appearance      =   1
   End
   Begin MSComctlLib.TreeView ctr������ 
      Height          =   5295
      Left            =   6000
      TabIndex        =   4
      Top             =   240
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   9340
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   6
      Appearance      =   1
   End
End
Attribute VB_Name = "frm���ÿ��ұ���"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim lnodParent As Node
    Dim lnodTemp As Node
    Dim lint���� As Integer
    Dim lobj��Ŀ As Object
    Dim lstrtvwme As String
    Dim lobj���� As Object
    Dim lint��Ŀ���� As Integer
    Dim lstr�շ���Ŀ��� As String
    
    On Error GoTo errHandler
    ccmd����.Enabled = False
    ccmd����.Enabled = False
    cmd����.Enabled = False
    Dim lrsd������Ŀ As Recordset
'    Dim lstr�շ���Ŀ��� As String
    Dim lstr���ұ�� As String
    '����Ŀ����ֵ
    

    Set lnodParent = ctr��Ŀ��.Nodes.Add(, , "s", "�շ���Ŀ")
    lint��Ŀ���� = Val(pobj�շѹ���.ҵ������("��Ŀ����"))
    If lint��Ŀ���� = 0 Then lint��Ŀ���� = 2

    For lint���� = 1 To lint��Ŀ����
        Set lobj��Ŀ = dafuncGetData("select �շ���Ŀ���,�շ���Ŀ���� from �շѹ���_�շ���Ŀ�ֵ�� where len(�շ���Ŀ���)=" & (lint���� * 3) & "  order by �շ���Ŀ��� ")
        
        If (Not lobj��Ŀ.BOF) And (Not lobj��Ŀ.EOF) Then
            lobj��Ŀ.MoveFirst
            Do While (Not lobj��Ŀ.EOF)
                lstrtvwme = "s" & lobj��Ŀ("�շ���Ŀ���").Value
                Set lnodTemp = ctr��Ŀ��.Nodes.Add("s" & Mid(lstrtvwme, 2, ((lint���� - 1) * 3)), tvwChild, lstrtvwme, lobj��Ŀ("�շ���Ŀ����").Value)

                lstr�շ���Ŀ��� = lobj��Ŀ("�շ���Ŀ���").Value
                Set lrsd������Ŀ = func��ѯ������Ŀ("'" & lstr�շ���Ŀ��� & "'")
                If lrsd������Ŀ.RecordCount > 0 Then
                   If (Not lrsd������Ŀ.BOF) And (Not lrsd������Ŀ.EOF) Then
                      lrsd������Ŀ.MoveFirst
                      Do While (Not lrsd������Ŀ.EOF)
                         lstr���ұ�� = lrsd������Ŀ("���ұ��").Value
                         Set lobj���� = dafuncGetData("select * from ϵͳ����_�����ֵ�� where ���= '" & lstr���ұ�� & "'")
                         Set lnodTemp = ctr��Ŀ��.Nodes.Add(lstrtvwme, tvwChild, "k" & lobj����("����") & lstrtvwme, lobj����("����"))
                         lrsd������Ŀ.MoveNext
                         Set lnodTemp = Nothing
                      Loop
                   End If
                End If
                lobj��Ŀ.MoveNext
                Set lnodTemp = Nothing
                ctr��Ŀ��.Refresh
            Loop
        End If
    Next
    
    '����������ֵ
    Set lnodParent = ctr������.Nodes.Add(, , "s", "��������")
    Set lobj���� = dafuncGetData("select * from ϵͳ����_�����ֵ�� order by ���")
    If (Not lobj����.BOF) And (Not lobj����.EOF) Then
        lobj����.MoveFirst
        Do While (Not lobj����.EOF)
            lstrtvwme = "s" & lobj����("���").Value
            Set lnodTemp = ctr������.Nodes.Add("s", tvwChild, lstrtvwme, lobj����("����").Value)
            lnodTemp.EnsureVisible 'չ�����нڵ�
            lobj����.MoveNext
            Set lnodTemp = Nothing
            ctr������.Refresh
        Loop
    End If
    Exit Sub
errHandler:
    sfsub������ "�շѽ��沿��", "frm���ÿ��ұ���", "Form_Load", Err.Number, Err.Description, False
End Sub
Public Function func��ѯ������Ŀ(ByVal para��ѯ���� As String) As Object
    On Error GoTo errHandler
    If para��ѯ���� = "ALL" Then
        Set func��ѯ������Ŀ = dafuncGetData("select * from  �շѹ���_���ұ��ʱ�")
    Else
        Set func��ѯ������Ŀ = dafuncGetData("select * from �շѹ���_���ұ��ʱ� where �շ���Ŀ���=" + para��ѯ����)
    End If
    
    Exit Function
errHandler:
    sfsub������ "�շѽ���", "frm����", "func��ѯ������Ŀ", Err.Number, Err.Description, True
End Function

Private Sub ccmd����_Click()
On Error GoTo errHandler
    Dim mstr�������� As Recordset
    If (ctr������.SelectedItem.Children = 0) And (Len(ctr��Ŀ��.SelectedItem.Key) = 7) Then
        ctr��Ŀ��.Nodes.Add ctr��Ŀ��.SelectedItem.Key, tvwChild, "k" & ctr������.SelectedItem.Text & ctr��Ŀ��.SelectedItem.Key, ctr������.SelectedItem.Text ' ctr��Ŀ��.SelectedItem.Key & ctr������.SelectedItem.Key
        Set mstr�������� = dafuncGetData("select * from ϵͳ����_�����ֵ�� where ����= '" & Trim(ctr������.SelectedItem.Text) & "'")
        sub���� Right(ctr��Ŀ��.SelectedItem.Key, 6), mstr��������("���") 'ctr������.SelectedItem.Text
        ctr��Ŀ��.Refresh
 
    End If
Exit Sub
errHandler:
    If Err.Number = 35602 Then
         Err.Number = 0
         'Err.Raise 6666
         
         Exit Sub
    End If
    sfsub������ "�շѽ��沿��", "frm���ÿ��ұ���", "ccmd����_Click", Err.Number, Err.Description, False
End Sub

Private Sub ccmd����_Click()
On Error GoTo errHandler
Dim mstr�������� As Recordset

If ctr��Ŀ��.SelectedItem.Children = 0 And Left(ctr��Ŀ��.SelectedItem.Key, 1) = "k" Then
   If MsgBox("ȷ��Ҫ��" & ctr��Ŀ��.SelectedItem.Parent.Text & "��Ŀ��ɾ���ÿ�����", vbOKCancel, "ϵͳ��ʾ") = vbOK Then
      Set mstr�������� = dafuncGetData("select * from ϵͳ����_�����ֵ�� where ����= '" & Trim(ctr��Ŀ��.SelectedItem.Text) & "'")
      subɾ�� Right(ctr��Ŀ��.SelectedItem.Key, 6), mstr��������("���") 'ctr��Ŀ��.SelectedItem.Text
    
    'ɾ��ѡ���Ľڵ�
    ctr��Ŀ��.Nodes.Remove ctr��Ŀ��.SelectedItem.Index
    ctr��Ŀ��.Refresh
   End If
End If

Exit Sub
errHandler:
    sfsub������ "�շѽ��沿��", "frm���ÿ��ұ���", "ccmd����_Click", Err.Number, Err.Description, False
End Sub

Private Sub cmd����_Click()
    If (ctr��Ŀ��.SelectedItem.Children > 0) And (Len(ctr��Ŀ��.SelectedItem.Key) = 7) Then
      
      frm����.pstr�շ���Ŀ��� = Right(ctr��Ŀ��.SelectedItem.Key, 6)
      frm����.Show 1
    End If
End Sub

Private Sub ctr������_NodeClick(ByVal Node As MSComctlLib.Node)
     If Len(ctr��Ŀ��.SelectedItem.Key) = 7 Then
        ccmd����.Enabled = True
     Else
        ccmd����.Enabled = False
     End If

End Sub

Private Sub ctr��Ŀ��_NodeClick(ByVal Node As MSComctlLib.Node)
    If Left(ctr��Ŀ��.SelectedItem.Key, 1) = "k" Then
      ccmd����.Enabled = True
    Else
      ccmd����.Enabled = False
    End If
    If Len(ctr��Ŀ��.SelectedItem.Key) = 7 And ctr��Ŀ��.SelectedItem.Children > 0 Then
      cmd����.Enabled = True
    Else
      cmd����.Enabled = False
    End If
End Sub


Public Sub sub����(para�շ���Ŀ���, para���ұ�� As String)
    On Error GoTo errHandler
    Dim lobjRec As Object '�����
    Set lobjRec = dafuncGetData("select * from �շѹ���_���ұ��ʱ� where �շ���Ŀ��� = '" + para�շ���Ŀ��� + "'and ���ұ�� ='" + para���ұ�� + "' ")
    
    If lobjRec.RecordCount > 0 Then
       Exit Sub
    End If
    
    dasubBeginTran
      dafuncGetData ("insert into �շѹ���_���ұ��ʱ�(�շ���Ŀ���,���ұ��)  values('" + para�շ���Ŀ��� + "','" + para���ұ�� + "')")
    dasubCommitTran
    
    Exit Sub
errHandler:
    sfsub������ "�շѽ���", "frm����", "sub����", Err.Number, Err.Description, True
End Sub

'���ܣ������ݿ�������շ���Ŀ��Ϣ�Ŀ�������
'�޸ģ��켽�� 2006/06/05
Public Sub subɾ��(para�շ���Ŀ���, para���ұ�� As String)
    On Error GoTo errHandler
    dafuncGetData ("delete from �շѹ���_���ұ��ʱ�  where �շ���Ŀ���='" + para�շ���Ŀ��� + "' and ���ұ��= '" + para���ұ�� + "'")
    Exit Sub
errHandler:
    sfsub������ "�շѽ���", "frm����", "subɾ��", Err.Number, Err.Description, True
End Sub
