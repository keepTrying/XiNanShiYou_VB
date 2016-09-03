VERSION 5.00
Begin VB.Form frm操作列表 
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
   Picture         =   "frm操作列表.frx":0000
   ScaleHeight     =   6960
   ScaleWidth      =   2160
   Begin VB.Label clbl操作 
      BackStyle       =   0  'Transparent
      Caption         =   "补录体检登记信息"
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
      Caption         =   "单位档案"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   2175
   End
End
Attribute VB_Name = "frm操作列表"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public pfrmParent As frmMain

Public Sub subClear()
    Dim i As Long
    
    '移除原有的操作组
    If clbl操作.Count > 1 Then
        For i = 1 To clbl操作.Count - 1
            Unload clbl操作(i)
        Next i
    End If

End Sub

Public Sub subAddOperation(ByVal paraCation As String, ByVal paraKey As String)
    Dim i As Long
    i = clbl操作.Count
    Load frm操作列表.clbl操作(i)
    frm操作列表.clbl操作(i).Caption = paraCation
    frm操作列表.clbl操作(i).Tag = paraKey
    frm操作列表.clbl操作(i).Top = clbl操作(i - 1).Top + clbl操作(i - 1).Height + 100
    frm操作列表.clbl操作(i).Visible = True

End Sub

Private Sub clbl操作_Click(Index As Integer)
    On Error Resume Next
    Call pfrmParent.sub创建窗体(clbl操作(Index).Tag)
    Unload Me
End Sub

Private Sub Form_LostFocus()
    Unload Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X > 4000 Then
        Unload frm操作列表
    End If

End Sub
