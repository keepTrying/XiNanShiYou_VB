VERSION 5.00
Begin VB.Form frm�����б� 
   BackColor       =   &H00B5D0D7&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6960
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   2160
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frm�����б�.frx":0000
   ScaleHeight     =   6960
   ScaleWidth      =   2160
   Begin VB.Label clbl���� 
      BackStyle       =   0  'Transparent
      Caption         =   "��¼���Ǽ���Ϣ"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1815
      WordWrap        =   -1  'True
   End
   Begin VB.Label clblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "��λ����"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   2175
   End
End
Attribute VB_Name = "frm�����б�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public pfrmParent As frmMain

Public Sub subClear()
    Dim i As Long
    
    '�Ƴ�ԭ�еĲ�����
    If clbl����.Count > 1 Then
        For i = 1 To clbl����.Count - 1
            Unload clbl����(i)
        Next i
    End If

End Sub

Public Sub subAddOperation(ByVal paraCation As String, ByVal paraKey As String)
    Dim i As Long
    i = clbl����.Count
    Load frm�����б�.clbl����(i)
    frm�����б�.clbl����(i).Caption = paraCation
    frm�����б�.clbl����(i).Tag = paraKey
    frm�����б�.clbl����(i).Top = clbl����(i - 1).Top + clbl����(i - 1).Height + 100
    frm�����б�.clbl����(i).Visible = True

End Sub

Private Sub clbl����_Click(Index As Integer)
    On Error Resume Next
    Call pfrmParent.sub��������(clbl����(Index).Tag)
    Unload Me
End Sub

Private Sub Form_LostFocus()
    Unload Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X > 4000 Then
        Unload frm�����б�
    End If

End Sub
