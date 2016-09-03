VERSION 5.00
Begin VB.Form frmhuatu 
   Caption         =   "Form1"
   ClientHeight    =   10005
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13035
   LinkTopic       =   "Form1"
   ScaleHeight     =   10005
   ScaleWidth      =   13035
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.PictureBox Picture1 
      Height          =   7575
      Left            =   120
      ScaleHeight     =   7515
      ScaleWidth      =   11475
      TabIndex        =   0
      Top             =   120
      Width           =   11535
      Begin VB.Image Image1 
         Height          =   6495
         Left            =   0
         Top             =   120
         Width           =   10335
      End
   End
End
Attribute VB_Name = "frmhuatu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Image1.Stretch = True
Picture1.Width = frmhuatu.Width
Picture1.Height = frmhuatu.Height
Image1.Width = Picture1.Width - 300
Image1.Height = Picture1.Height - 500
End Sub
