VERSION 5.00
Begin VB.Form frmҵ������ 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ҵ������"
   ClientHeight    =   6420
   ClientLeft      =   150
   ClientTop       =   675
   ClientWidth     =   8880
   Icon            =   "frmҵ������.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   8880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Menu cmnuItemOther 
      Caption         =   "�շ���Ŀ(&I)"
      Index           =   1
   End
   Begin VB.Menu cmnuItemOther 
      Caption         =   "�շѱ�׼(&T)"
      Index           =   2
   End
   Begin VB.Menu cmnuItemOther 
      Caption         =   "Ʊ�ݸ�ʽ(&F)"
      Index           =   3
   End
   Begin VB.Menu cmnuItemOther 
      Caption         =   "����(&D)"
      Index           =   4
   End
   Begin VB.Menu cmnuItemOther 
      Caption         =   "��������(&B)"
      Index           =   5
   End
   Begin VB.Menu cmnuBase 
      Caption         =   "�˳�ϵͳ"
   End
End
Attribute VB_Name = "frmҵ������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public pblnInUse As Boolean

Private WithEvents mobjGUI As cls����ͨ�ö���
Attribute mobjGUI.VB_VarHelpID = -1
Dim mlngID As Long      '��ǰ�޸ĵĺŶε�ID

Private Sub ccmdSet_Click(Index As Integer)
    On Error GoTo errhandler
    Select Case Index
    Case 0, 1, 2
        cmnuItemOther_Click Index + 1
    Case 3
        cmnuItemOther_Click 5
    Case 4
        frm���ÿ��ұ���.Show 1, Me
        
    End Select
    Exit Sub
errhandler:
    sfsub������ "�շѽ��沿��", "frmҵ������", "ccmdSet_Click", Err.Number, Err.Description, False
    Exit Sub
    Resume
End Sub

Private Sub cmnuBase_Click()
    Unload Me
End Sub

Private Sub cmnuItemOther_Click(Index As Integer)
    On Error GoTo errhandler
    Select Case Index
    Case 1
        frm�����շ���Ŀ.Move Me.Left, Me.Top
        frm�����շ���Ŀ.Show 1
    Case 2 '�շѱ�׼
        frm�����շѱ�׼.Move Me.Left, Me.Top
        frm�����շѱ�׼.Show 1, Me
    Case 3 'Ʊ�ݸ�ʽ
        frm����Ʊ�ݸ�ʽ.Move Me.Left, Me.Top
        frm����Ʊ�ݸ�ʽ.Show 1, Me
    Case 4 '����
        frm���ô���.Move Me.Left, Me.Top
        frm���ô���.Show 1, Me
    Case 5
        frm����������.Move Me.Left, Me.Top
        frm����������.Show 1, Me
        
    End Select
    
    Exit Sub
errhandler:
    sfsub������ "�շѽ��沿��", "frmҵ������", "Form_Load", Err.Number, Err.Description, False
    Exit Sub
    Resume
    
End Sub

Private Sub Form_Load()
        
    If pblnInUse Then Exit Sub
    pblnInUse = True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    pblnInUse = False
End Sub
