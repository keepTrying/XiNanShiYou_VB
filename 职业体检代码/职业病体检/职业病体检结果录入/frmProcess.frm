VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmProcess 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   780
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   -70
      Width           =   4575
      Begin MSComctlLib.ProgressBar proPercent 
         Height          =   210
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   370
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "正在保存，请等待......"
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   1290
         TabIndex        =   2
         Top             =   240
         Width           =   1995
      End
   End
End
Attribute VB_Name = "frmProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
