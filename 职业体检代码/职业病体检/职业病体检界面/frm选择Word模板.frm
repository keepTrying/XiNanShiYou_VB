VERSION 5.00
Begin VB.Form frmѡ��Wordģ�� 
   Caption         =   "��ѡ���������ɱ����Wordģ��"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   ScaleHeight     =   5445
   ScaleWidth      =   5160
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton ccmdCancel 
      Caption         =   "ȡ  ��"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   4920
      Width           =   855
   End
   Begin VB.CommandButton ccmdOk 
      Caption         =   "ȷ  ��"
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   4920
      Width           =   855
   End
   Begin VB.ListBox clstFile 
      Height          =   4020
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   4695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��ѡ��ģ���ļ���"
      Height          =   180
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1440
   End
End
Attribute VB_Name = "frmѡ��Wordģ��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public pstrFilename As String
Public pstrWordname As String

Private Sub ccmdCancel_Click()
    pstrFilename = ""
    Unload Me
End Sub

Private Sub ccmdOk_Click()
    If clstFile.ListIndex >= 0 Then
        pstrFilename = clstFile.Text
    Else
        pstrFilename = ""
    End If
    Unload Me
End Sub

Private Sub clstFile_DblClick()
    ccmdOk_Click
End Sub

Private Sub Form_Load()
    Dim lstrFile As String
    Dim lobjRec As Object
    
    lstrFile = Dir(App.Path & "\ͨ��_*.dot")
    Do While lstrFile <> ""
       clstFile.AddItem lstrFile
       lstrFile = Dir
    Loop
    'Ѱ�ҵ�ǰ�û�������������Ӧ��ר��ģ����ǰ׺
    Set lobjRec = dafuncGetData("select ���� from ϵͳ����_�����ֵ�� where ���='" & um�û��������ұ�� & "'")
'    If lobjRec(0) <> "" Then
'        lstrFile = Dir(App.Path & "\" & lobjRec(0) & "_�Ĵ�ʡ" & Left(pstrWordname, 2) & "*.dot")
        lstrFile = Dir(App.Path & "\ְҵ�����_�Ĵ�ʡ" & Left(pstrWordname, 2) & "*.dot")
'    lstrFile = Dir(App.Path & "\ְҵ�����_�Ĵ�ʡ*.dot")
'        MsgBox lstrFile
'        Do While lstrFile <> ""
           clstFile.AddItem lstrFile
'           lstrFile = Dir
'        Loop
    
'    End If
    If clstFile.ListCount = 0 Then
        MsgBox "û���ҵ�Wordģ���ļ���", vbInformation, "ϵͳ��ʾ"
        pstrFilename = ""
        'Unload Me
'    ElseIf clstFile.ListCount = 1 Then
'        pstrFilename = lstrFile
'        Unload Me
    Else
    
        pstrFilename = lstrFile
'        clstFile.ListIndex = 0
    End If
End Sub
