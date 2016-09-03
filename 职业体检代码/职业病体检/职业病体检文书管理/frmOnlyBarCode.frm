VERSION 5.00
Object = "{D9347025-9612-11D1-9D75-00C04FCC8CDC}#1.0#0"; "MSBCODE9.OCX"
Begin VB.Form frmOnlyBarCode 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   14850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   ScaleHeight     =   14850
   ScaleWidth      =   12000
   StartUpPosition =   3  '窗口缺省
   Begin BARCODELibCtl.BarCodeCtrl cbccBarCode 
      Height          =   1335
      Index           =   0
      Left            =   480
      TabIndex        =   14
      Top             =   360
      Width           =   2895
      Style           =   7
      SubStyle        =   0
      Validation      =   0
      LineWeight      =   3
      Direction       =   0
      ShowData        =   1
      Value           =   "123456 Code-128"
      ForeColor       =   0
      BackColor       =   16777215
   End
   Begin BARCODELibCtl.BarCodeCtrl cbccBarCode 
      Height          =   1335
      Index           =   3
      Left            =   480
      TabIndex        =   13
      Top             =   3360
      Visible         =   0   'False
      Width           =   2895
      Style           =   7
      SubStyle        =   -1
      Validation      =   0
      LineWeight      =   3
      Direction       =   0
      ShowData        =   1
      Value           =   "123456 Code-128"
      ForeColor       =   0
      BackColor       =   16777215
   End
   Begin BARCODELibCtl.BarCodeCtrl cbccBarCode 
      Height          =   1335
      Index           =   12
      Left            =   480
      TabIndex        =   12
      Top             =   13560
      Visible         =   0   'False
      Width           =   2895
      Style           =   7
      SubStyle        =   0
      Validation      =   0
      LineWeight      =   3
      Direction       =   0
      ShowData        =   1
      Value           =   "123456 Code-128"
      ForeColor       =   0
      BackColor       =   16777215
   End
   Begin BARCODELibCtl.BarCodeCtrl cbccBarCode 
      Height          =   1335
      Index           =   9
      Left            =   480
      TabIndex        =   11
      Top             =   10200
      Visible         =   0   'False
      Width           =   2895
      Style           =   7
      SubStyle        =   0
      Validation      =   0
      LineWeight      =   3
      Direction       =   0
      ShowData        =   1
      Value           =   "123456 Code-128"
      ForeColor       =   0
      BackColor       =   16777215
   End
   Begin BARCODELibCtl.BarCodeCtrl cbccBarCode 
      Height          =   1335
      Index           =   6
      Left            =   480
      TabIndex        =   10
      Top             =   6720
      Visible         =   0   'False
      Width           =   2895
      Style           =   7
      SubStyle        =   -1
      Validation      =   0
      LineWeight      =   3
      Direction       =   0
      ShowData        =   1
      Value           =   "123456 Code-128"
      ForeColor       =   0
      BackColor       =   16777215
   End
   Begin BARCODELibCtl.BarCodeCtrl cbccBarCode 
      Height          =   1335
      Index           =   1
      Left            =   4680
      TabIndex        =   9
      Top             =   360
      Visible         =   0   'False
      Width           =   2895
      Style           =   7
      SubStyle        =   -1
      Validation      =   0
      LineWeight      =   3
      Direction       =   0
      ShowData        =   1
      Value           =   "123456 Code-128"
      ForeColor       =   0
      BackColor       =   16777215
   End
   Begin BARCODELibCtl.BarCodeCtrl cbccBarCode 
      Height          =   1335
      Index           =   2
      Left            =   8880
      TabIndex        =   8
      Top             =   360
      Visible         =   0   'False
      Width           =   2895
      Style           =   7
      SubStyle        =   -1
      Validation      =   0
      LineWeight      =   3
      Direction       =   0
      ShowData        =   1
      Value           =   "123456 Code-128"
      ForeColor       =   0
      BackColor       =   16777215
   End
   Begin BARCODELibCtl.BarCodeCtrl cbccBarCode 
      Height          =   1335
      Index           =   14
      Left            =   8880
      TabIndex        =   7
      Top             =   13560
      Visible         =   0   'False
      Width           =   2895
      Style           =   7
      SubStyle        =   0
      Validation      =   0
      LineWeight      =   3
      Direction       =   0
      ShowData        =   1
      Value           =   "123456 Code-128"
      ForeColor       =   0
      BackColor       =   16777215
   End
   Begin BARCODELibCtl.BarCodeCtrl cbccBarCode 
      Height          =   1335
      Index           =   11
      Left            =   8880
      TabIndex        =   6
      Top             =   10200
      Visible         =   0   'False
      Width           =   2895
      Style           =   7
      SubStyle        =   0
      Validation      =   0
      LineWeight      =   3
      Direction       =   0
      ShowData        =   1
      Value           =   "123456 Code-128"
      ForeColor       =   0
      BackColor       =   16777215
   End
   Begin BARCODELibCtl.BarCodeCtrl cbccBarCode 
      Height          =   1335
      Index           =   8
      Left            =   8880
      TabIndex        =   5
      Top             =   6720
      Visible         =   0   'False
      Width           =   2895
      Style           =   7
      SubStyle        =   0
      Validation      =   0
      LineWeight      =   3
      Direction       =   0
      ShowData        =   1
      Value           =   "123456 Code-128"
      ForeColor       =   0
      BackColor       =   16777215
   End
   Begin BARCODELibCtl.BarCodeCtrl cbccBarCode 
      Height          =   1335
      Index           =   5
      Left            =   8880
      TabIndex        =   4
      Top             =   3360
      Visible         =   0   'False
      Width           =   2895
      Style           =   7
      SubStyle        =   -1
      Validation      =   0
      LineWeight      =   3
      Direction       =   0
      ShowData        =   1
      Value           =   "123456 Code-128"
      ForeColor       =   0
      BackColor       =   16777215
   End
   Begin BARCODELibCtl.BarCodeCtrl cbccBarCode 
      Height          =   1335
      Index           =   13
      Left            =   4680
      TabIndex        =   3
      Top             =   13560
      Visible         =   0   'False
      Width           =   2895
      Style           =   7
      SubStyle        =   0
      Validation      =   0
      LineWeight      =   3
      Direction       =   0
      ShowData        =   1
      Value           =   "123456 Code-128"
      ForeColor       =   0
      BackColor       =   16777215
   End
   Begin BARCODELibCtl.BarCodeCtrl cbccBarCode 
      Height          =   1335
      Index           =   10
      Left            =   4680
      TabIndex        =   2
      Top             =   10200
      Visible         =   0   'False
      Width           =   2895
      Style           =   7
      SubStyle        =   0
      Validation      =   0
      LineWeight      =   3
      Direction       =   0
      ShowData        =   1
      Value           =   "123456 Code-128"
      ForeColor       =   0
      BackColor       =   16777215
   End
   Begin BARCODELibCtl.BarCodeCtrl cbccBarCode 
      Height          =   1335
      Index           =   7
      Left            =   4680
      TabIndex        =   1
      Top             =   6720
      Visible         =   0   'False
      Width           =   2895
      Style           =   7
      SubStyle        =   -1
      Validation      =   0
      LineWeight      =   3
      Direction       =   0
      ShowData        =   1
      Value           =   "123456 Code-128"
      ForeColor       =   0
      BackColor       =   16777215
   End
   Begin BARCODELibCtl.BarCodeCtrl cbccBarCode 
      Height          =   1335
      Index           =   4
      Left            =   4680
      TabIndex        =   0
      Top             =   3360
      Visible         =   0   'False
      Width           =   2895
      Style           =   7
      SubStyle        =   -1
      Validation      =   0
      LineWeight      =   3
      Direction       =   0
      ShowData        =   1
      Value           =   "123456 Code-128"
      ForeColor       =   0
      BackColor       =   16777215
   End
End
Attribute VB_Name = "frmOnlyBarCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Public 体检条码号 As String

Private Sub Form_Load()
    'cbccBarCode.ShowData = 1
    'cbccBarCode.Style = 7
End Sub
