VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���� ���߹���ϵͳ"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5385
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton ccmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   2520
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   360
      Picture         =   "frmAbout.frx":000C
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   4
      Top             =   240
      Width           =   480
   End
   Begin VB.Frame Frame1 
      Height          =   75
      Left            =   120
      TabIndex        =   0
      Top             =   2220
      Width           =   5175
   End
   Begin VB.Label clblVersion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Version 3.0"
      Height          =   255
      Left            =   1800
      TabIndex        =   1
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��Ȩ����(C)2003 �ɶ���Դ������޹�˾"
      Height          =   180
      Left            =   1800
      TabIndex        =   3
      Top             =   1920
      Width           =   3240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����ۺϹ�����Ϣϵͳ"
      Height          =   180
      Left            =   1815
      TabIndex        =   2
      Top             =   240
      Width           =   1800
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ccmdOk_Click()
     Unload Me
End Sub

Private Sub Form_Load()
'    Dim lstrVersion As String
'    Dim lstrSp As String
'    On Error Resume Next
'    '��ȡ�汾�š�
'    lstrVersion = sffuncGetVersion(pstrSysName)
'    If lstrVersion = "" Then
'        lstrVersion = "2.6"
'    ElseIf Len(lstrVersion) > 3 Then
'        lstrSp = Trim(Str(Val(Right(lstrVersion, 3))))
'        lstrVersion = Left(lstrVersion, 3)
'    End If
'    If lstrSp <> "" Then
'        lstrVersion = lstrVersion & "(SP" & lstrSp & ")"
'    End If
'    clblVersion.Caption = "Version " & lstrVersion
    Me.Caption = "���� " & pstrSysName
    Label1.Caption = pstrSysName ' & "�����ð棩"
End Sub

