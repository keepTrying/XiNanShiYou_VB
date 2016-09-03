VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmProgress 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "正在处理，请稍候..."
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5445
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin ComctlLib.ProgressBar cpbMain 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
      MousePointer    =   2
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   0
      Picture         =   "frmProgress.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   945
   End
   Begin VB.Label clblInfo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "正在处理..."
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   405
      Left            =   1080
      TabIndex        =   1
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

