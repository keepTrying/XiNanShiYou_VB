VERSION 5.00
Begin VB.Form frmˢ���� 
   Caption         =   "Form1"
   ClientHeight    =   2325
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   ScaleHeight     =   2325
   ScaleWidth      =   6105
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton cmdquit 
      Caption         =   "ȡ  ��"
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "ȷ  ��"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox Text���� 
      Height          =   270
      Left            =   1440
      TabIndex        =   1
      Top             =   840
      Width           =   3375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��ˢ���룺"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   1200
   End
End
Attribute VB_Name = "frmˢ����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdok_Click()
    FrmRegister.Show 1, Me
End Sub

Private Sub cmdquit_Click()
    Unload Me
End Sub

