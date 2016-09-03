VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4185
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9300
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   9300
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "save jpg object"
      Height          =   495
      Index           =   1
      Left            =   5400
      TabIndex        =   6
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "save jpg object"
      Height          =   495
      Index           =   0
      Left            =   4080
      TabIndex        =   4
      Top             =   3360
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4680
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "save jpg file"
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton ccmdLoad 
      Caption         =   "Load bmp"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   3360
      Width           =   975
   End
   Begin VB.Image Image2 
      Height          =   2775
      Index           =   1
      Left            =   5400
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "转换后的jpg(70):"
      Height          =   180
      Index           =   2
      Left            =   5400
      TabIndex        =   5
      Top             =   120
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "转换后的jpg(75):"
      Height          =   180
      Index           =   1
      Left            =   3000
      TabIndex        =   3
      Top             =   120
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "转换前的bmp:"
      Height          =   180
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   1080
   End
   Begin VB.Image Image2 
      Height          =   2775
      Index           =   0
      Left            =   3000
      Top             =   480
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   2775
      Left            =   360
      Top             =   480
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ccmdLoad_Click()
    
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName <> "" Then
        Image1.Picture = LoadPicture(CommonDialog1.FileName)
    End If
End Sub

Private Sub Command1_Click()
    Dim lobjTransfer As Object
    Set lobjTransfer = CreateObject("BmpToJGP.clsBmpToJPG")
    lobjTransfer.subSetBMP Image1.Picture
    lobjTransfer.subSaveToJPGFile App.Path & "\a.jpg"
    Image2(0).Picture = LoadPicture(App.Path & "\a.jpg")
End Sub

Private Sub Command2_Click(index As Integer)
    Dim lobjTransfer As Object
    Set lobjTransfer = CreateObject("BmpToJGP.clsBmpToJPG")
    lobjTransfer.subSetBMP Image1.Picture
    
'    Set Image2(index).Picture = lobjTransfer.funcSaveToJPG
    Set Image2(index).Picture = lobjTransfer.funcSaveToJPG(IIf(index = 0, 75, 70))
End Sub

