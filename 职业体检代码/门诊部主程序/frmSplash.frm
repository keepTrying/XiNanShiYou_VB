VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5850
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   9285
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   9285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ṩ�̣��ɶ�����ʽ�Ƽ����޹�˾"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00664802&
      Height          =   285
      Left            =   3945
      TabIndex        =   1
      Top             =   5010
      Width           =   4725
   End
   Begin VB.Label clblSys 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "������Ϣϵͳ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00664802&
      Height          =   840
      Left            =   30
      TabIndex        =   0
      Top             =   1005
      Width           =   7365
   End
   Begin VB.Image Image1 
      Height          =   6030
      Left            =   0
      Picture         =   "frmSplash.frx":000C
      Top             =   0
      Width           =   9360
   End
End
Attribute VB_Name = "frmSplash"
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
    clblSys.Caption = "��ӭ����" & um����վ�� & "������Ϣϵͳ"
    
    Exit Sub
errHandle:
    Call sfsub������("������", "frmSplash", "form_load", Err.Number, Err.Description, False)
End Sub


