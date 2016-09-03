VERSION 5.00
Begin VB.Form frm刷条码 
   Caption         =   "Form1"
   ClientHeight    =   2325
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   ScaleHeight     =   2325
   ScaleWidth      =   6105
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdquit 
      Caption         =   "取  消"
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "确  定"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox Text条码 
      Height          =   270
      Left            =   1440
      TabIndex        =   1
      Top             =   840
      Width           =   3375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "请刷条码："
      BeginProperty Font 
         Name            =   "宋体"
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
Attribute VB_Name = "frm刷条码"
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

