VERSION 5.00
Begin VB.Form frmConclunInput_Write 
   Caption         =   "结论录入窗口"
   ClientHeight    =   4815
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   8775
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox ctxtConclusion 
      Height          =   2895
      Left            =   720
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1080
      Width           =   7335
   End
   Begin VB.Label llabDoctor 
      BackColor       =   &H00C0E0FF&
      Caption         =   "医师："
      Height          =   255
      Left            =   3960
      TabIndex        =   1
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label llabDept 
      BackColor       =   &H00C0E0FF&
      Caption         =   "结论科室："
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
End
Attribute VB_Name = "frmConclunInput_Write"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
