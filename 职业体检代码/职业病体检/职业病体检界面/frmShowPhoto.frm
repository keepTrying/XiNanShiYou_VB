VERSION 5.00
Begin VB.Form frmShowPhoto 
   Caption         =   "Form1"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12165
   LinkTopic       =   "Form1"
   ScaleHeight     =   9180
   ScaleWidth      =   12165
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox Picture1 
      Height          =   6615
      Left            =   360
      ScaleHeight     =   6555
      ScaleWidth      =   11595
      TabIndex        =   5
      Top             =   1680
      Width           =   11655
      Begin VB.Image Image1 
         Height          =   6015
         Left            =   120
         Top             =   240
         Width           =   11295
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�鿴"
      Height          =   495
      Left            =   5640
      TabIndex        =   4
      Top             =   480
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   1440
      TabIndex        =   3
      Text            =   "�ڼ����ϴε���Ƭ"
      Top             =   840
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "�鿴�ڼ���"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "ϵͳ���"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmShowPhoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
