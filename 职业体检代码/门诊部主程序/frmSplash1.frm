VERSION 5.00
Begin VB.Form frmSplash1 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7230
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   11145
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   11145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "�����̣��ɶ�����ʽ�Ƽ����޹�˾"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   5490
      TabIndex        =   1
      Top             =   6645
      Width           =   4830
   End
   Begin VB.Label clblSys 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�������������ල��������Ϣϵͳ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   435
      Left            =   1995
      TabIndex        =   0
      Top             =   5700
      Width           =   6750
   End
   Begin VB.Image Image1 
      Height          =   5310
      Left            =   -75
      Picture         =   "frmSplash1.frx":000C
      Top             =   -30
      Width           =   11295
   End
End
Attribute VB_Name = "frmSplash1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
    On Error GoTo errHandle
    'ȡ������
    'lbl����վ��.Caption = lbl����վ��.Caption & um����վ��
    
    '��ȡ�汾�š�
'    Dim lstrVersion  As String
'    lstrVersion = sffuncGetVersion(pstrSysName)
'    If lstrVersion = "" Then lstrVersion = "3.0"
'    Label1.Caption = "V " & lstrVersion
    
    'clblSys.Caption = IIf(pstrSysName = "", "�������߹�����Ϣϵͳ", pstrSysName)
    clblSys.Caption = um����վ�� & "������Ϣϵͳ"
    
    Exit Sub
errHandle:
    Call sfsub������("������", "frmSplash", "form_load", Err.Number, Err.Description, False)
End Sub


