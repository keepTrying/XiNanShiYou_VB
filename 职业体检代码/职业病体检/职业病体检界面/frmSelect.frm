VERSION 5.00
Begin VB.Form frmSelect 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8385
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   8385
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��"
      Height          =   615
      Left            =   3240
      TabIndex        =   1
      Top             =   1560
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frmSelect.frx":0000
      Left            =   360
      List            =   "frmSelect.frx":000A
      TabIndex        =   0
      Top             =   480
      Width           =   7335
   End
   Begin VB.Label Label1 
      Caption         =   "��ѡ��Ҫ��ӡ��ָ������"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If Combo1.Text <> "" Then
        frm��ӡ����嵥�򻯰�.Label1.Caption = Combo1.Text
        Unload Me
    Else
        MsgBox "��ѡ��ָ����"
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Combo1.ListIndex = 1
End Sub

