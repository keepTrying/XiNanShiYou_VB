VERSION 5.00
Begin VB.Form frm�޸��շ� 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "�޸�"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.ComboBox cmb���ѷ�ʽ 
      Height          =   300
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   360
      Width           =   2445
   End
   Begin VB.CommandButton ccmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton ccmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���ѷ�ʽ"
      Height          =   180
      Index           =   24
      Left            =   720
      TabIndex        =   3
      Top             =   390
      Width           =   720
   End
End
Attribute VB_Name = "frm�޸��շ�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public pstr�շ����� As String
Public pblnOk As Boolean
Private Sub ccmdCancel_Click()
    pblnOk = False
    Unload Me
End Sub

Private Sub ccmdOK_Click()
    On Error GoTo errhandler
    If cmb���ѷ�ʽ.Text = "" Then
        Err.Raise 6666, , "���ѷ�ʽ������գ�"
    End If
    dafuncGetData "update �շѹ���_������Ϣ�� set ���ѷ�ʽ=" & Right(cmb���ѷ�ʽ.ItemData(cmb���ѷ�ʽ.ListIndex), Len(Trim(Str(cmb���ѷ�ʽ.ItemData(cmb���ѷ�ʽ.ListIndex)))) - 1) & " where �շѱ��='" & pstr�շ����� & "'"
    
    Unload Me
    pblnOk = True
    
    Exit Sub
errhandler:
    sfsub������ "�շѽ��沿��", "frm�޸��շ�", "ccmdOK_Click()", Err.Number, Err.Description, False
    
End Sub

Private Sub Form_Load()
    Dim lobjRec As Object
    Dim lobj���ѷ�ʽ As Object
    On Error GoTo errhandler
    Set lobjRec = dafuncGetData("select ���ѷ�ʽ from �շѹ���_������Ϣ�� where �շѱ��='" & pstr�շ����� & "'")
    
    Set lobj���ѷ�ʽ = dafuncGetData("select * from �շѹ���_���ѷ�ʽ�ֵ��")
    '��ʼ�����ѷ�ʽ�б�
    If Not (lobj���ѷ�ʽ Is Nothing) Then
        Do While Not lobj���ѷ�ʽ.EOF
            cmb���ѷ�ʽ.AddItem lobj���ѷ�ʽ("����").Value
            cmb���ѷ�ʽ.ItemData(cmb���ѷ�ʽ.ListCount - 1) = "1" & lobj���ѷ�ʽ("���")
            
            If Val(lobj���ѷ�ʽ("���")) = lobjRec!���ѷ�ʽ Then
                cmb���ѷ�ʽ.ListIndex = cmb���ѷ�ʽ.ListCount - 1
            End If
            
            lobj���ѷ�ʽ.MoveNext
        Loop
        'cmb���ѷ�ʽ.ListIndex = 0
    End If
    Exit Sub
errhandler:
    sfsub������ "�շѽ��沿��", "frm�޸��շ�", "Form_Load", Err.Number, Err.Description, False
    
End Sub
