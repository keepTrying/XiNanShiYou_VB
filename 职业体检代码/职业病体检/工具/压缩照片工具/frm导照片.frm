VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frm����Ƭ 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ѹ����콡��֤��Ƭ"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton ccmdExit 
      Caption         =   "�˳�(&X)"
      Height          =   495
      Left            =   4920
      TabIndex        =   1
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton ccmdExport 
      Caption         =   "��ʼѹ��(&S)"
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   1080
      Width           =   1455
   End
   Begin ComctlLib.StatusBar Cstau״̬�� 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   2925
      Width           =   7830
      _ExtentX        =   13811
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   7409
            MinWidth        =   7409
            Text            =   ""
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm����Ƭ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ccmdExit_Click()
    Unload Me
End Sub

Private Sub ccmdExport_Click()
    Dim lobj���ݷ��� As Object
    On Error GoTo errHandler
    MousePointer = 11
    ccmdExit.Enabled = False
    ccmdExport.Enabled = False
    Cstau״̬��.Panels.Item(1).Text = "����ѹ����콡��֤ͼƬ..."
    
    Set lobj���ݷ��� = New Cls���ݷ���
    lobj���ݷ���.sub�������

    Cstau״̬��.Panels.Item(1).Text = "��Ƭѹ����ϣ�"
    MousePointer = 0
    ccmdExit.Enabled = True
    ccmdExport.Enabled = True
    Exit Sub
errHandler:
    MsgBox Error, vbOKOnly + vbExclamation, "ϵͳ��ʾ"
End Sub

