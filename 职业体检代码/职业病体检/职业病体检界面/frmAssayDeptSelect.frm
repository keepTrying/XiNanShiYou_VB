VERSION 5.00
Begin VB.Form frmAssayDeptSelect 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "��ӡ�Թܱ�ǩ"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame1 
      Caption         =   "ѡ���ǩ����"
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.CheckBox cchkType 
         Caption         =   "�ι�2��������GLU��Ѫ֬��ACP"
         Height          =   375
         Index           =   5
         Left            =   2160
         TabIndex        =   8
         Top             =   1200
         Width           =   1935
      End
      Begin VB.CommandButton ccmdCancel 
         Caption         =   "ȡ��"
         Height          =   375
         Left            =   2160
         TabIndex        =   7
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CommandButton ccmdConfirm 
         Caption         =   "ȷ��"
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CheckBox cchkType 
         Caption         =   "Ⱦɫ��"
         Height          =   375
         Index           =   4
         Left            =   360
         TabIndex        =   5
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CheckBox cchkType 
         Caption         =   "�ι�1�����������԰룬GLU��Ѫ֬��ACP"
         Height          =   375
         Index           =   3
         Left            =   2160
         TabIndex        =   4
         Top             =   720
         Width           =   1935
      End
      Begin VB.CheckBox cchkType 
         Caption         =   "����.Ѫ��"
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   3
         Top             =   720
         Width           =   1455
      End
      Begin VB.CheckBox cchkType 
         Caption         =   "�򳣹�"
         Height          =   375
         Index           =   1
         Left            =   2160
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
      Begin VB.CheckBox cchkType 
         Caption         =   "Ѫ����.����Ѫ"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmAssayDeptSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'2012-08-19 �ڵ��
'ѡ���ӡ��ǩ���ͣ������ȷ�����󣬼��ɴ�ӡ��ǩ

Option Explicit
Public pblnOk As Boolean
Public selectedDeptName As Collection

Private Sub ccmdCancel_Click()
    pblnOk = False
    subExit
End Sub

Private Sub ccmdConfirm_Click()
    Dim i As Integer
    
    Set selectedDeptName = New Collection
    selectedDeptName.Add ""
    For i = 0 To cchkType.Count - 1
        If cchkType(i).Value = 1 Then
            selectedDeptName.Add cchkType(i).Caption
        End If
    Next i
    
    pblnOk = True
    subExit
End Sub

Sub subExit()
    'End
    Unload frmAssayDeptSelect
End Sub
