VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmƽ̨���� 
   BackColor       =   &H8000000B&
   BorderStyle     =   0  'None
   Caption         =   "ƽ̨����"
   ClientHeight    =   7455
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   10470
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmƽ̨����.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7455
   ScaleMode       =   0  'User
   ScaleWidth      =   10337.84
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin MSComctlLib.TreeView ctrwȨ�� 
      Height          =   5175
      Left            =   6000
      TabIndex        =   14
      Top             =   1680
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   9128
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   804
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      Appearance      =   1
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000B&
      Height          =   465
      Left            =   60
      TabIndex        =   4
      Top             =   840
      Width           =   10500
      Begin VB.OptionButton coptѡ�� 
         BackColor       =   &H8000000B&
         Caption         =   "������Ϣ"
         Height          =   270
         Index           =   3
         Left            =   2040
         TabIndex        =   8
         Top             =   120
         Width           =   1215
      End
      Begin VB.OptionButton coptѡ�� 
         BackColor       =   &H8000000B&
         Caption         =   "����"
         Height          =   270
         Index           =   0
         Left            =   255
         TabIndex        =   7
         Top             =   120
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton coptѡ�� 
         BackColor       =   &H8000000B&
         Caption         =   "����"
         Height          =   270
         Index           =   1
         Left            =   4200
         TabIndex        =   6
         Top             =   120
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.OptionButton coptѡ�� 
         BackColor       =   &H8000000B&
         Caption         =   "��ѯ"
         Height          =   270
         Index           =   2
         Left            =   6480
         TabIndex        =   5
         Top             =   120
         Visible         =   0   'False
         Width           =   1695
      End
   End
   Begin VB.Frame fram���� 
      BackColor       =   &H8000000B&
      Height          =   5130
      Left            =   4800
      TabIndex        =   1
      Top             =   1800
      Width           =   1155
      Begin VB.CommandButton ccmd�ƶ� 
         BackColor       =   &H8000000B&
         Caption         =   "<"
         Enabled         =   0   'False
         Height          =   360
         Index           =   0
         Left            =   120
         MaskColor       =   &H00FFF1EC&
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1800
         Width           =   960
      End
      Begin VB.CommandButton ccmd�ƶ� 
         BackColor       =   &H8000000B&
         Caption         =   ">"
         Enabled         =   0   'False
         Height          =   360
         Index           =   1
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   4080
         Width           =   960
      End
      Begin VB.Label clblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "�����ѡ�����ƽ̨����������ļ�ʱ���ſ���ȥ��������"
         ForeColor       =   &H8000000D&
         Height          =   1095
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label clblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "�����ѡ�����ƽ̨����������ļ�ʱ����ѡ���ұ�Ȩ��ʱ�ſ�����Ӳ�����"
         ForeColor       =   &H8000000D&
         Height          =   1335
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CheckBox ccheckȱʡ 
      BackColor       =   &H8000000B&
      Caption         =   "ʹ��ȱʡ����"
      Height          =   345
      Left            =   4440
      TabIndex        =   0
      Top             =   1320
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5010
      Top             =   3855
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmƽ̨����.frx":030A
            Key             =   "Second"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmƽ̨����.frx":039E
            Key             =   "First"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList cimg��ťͼ�� 
      Left            =   8220
      Top             =   870
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   21
      ImageHeight     =   21
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.TreeView ctvƽ̨�� 
      Height          =   5235
      Left            =   90
      TabIndex        =   9
      Top             =   1680
      Width           =   4590
      _ExtentX        =   8096
      _ExtentY        =   9234
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin MSComctlLib.Toolbar c������ 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   10470
      _ExtentX        =   18468
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "cimg��ťͼ��"
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar c״̬�� 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   7080
      Width           =   10470
      _ExtentX        =   18468
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Label clbl��ǩ 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ƽ̨��(˫������������ȥ������)"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   1440
      Width           =   2700
   End
   Begin VB.Label clbl��ǩ 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ǰ�û����õĲ���Ȩ��(˫���������������)"
      Height          =   180
      Index           =   1
      Left            =   6120
      TabIndex        =   11
      Top             =   1440
      Width           =   3780
   End
   Begin VB.Menu ƽ̨ 
      Caption         =   "ƽ̨"
      Visible         =   0   'False
      Begin VB.Menu add 
         Caption         =   "���(&A)"
         Index           =   1
      End
      Begin VB.Menu delete 
         Caption         =   "ɾ��(&D)"
         Index           =   2
      End
      Begin VB.Menu modify 
         Caption         =   "�޸�(&U)"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmƽ̨����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mobj�� As Collection               '��ǰƽ̨�Ĳ��������
Private mobjSmartInfos As Collection        '��ǰ�û���������Ϣ
Private mobj���� As Collection              '��ǰƽ̨�����в���
Private mobj���ò��� As Object
Private mobj��ѯ As Collection              '��ǰƽ̨�����в�ѯ
Private mobj���� As Collection               '��ǰƽ̨�����б���
Private mobjȨ�� As Object                  '��ǰ�û����ò���Ȩ��
Private mint��ǰѡ�� As Integer            ' ��ǰ�û���ѡ��Ĳ���
Private WithEvents mobjGUI As cls����ͨ�ö��� '���������õĽ���ͨ�ö���
Attribute mobjGUI.VB_VarHelpID = -1
Private mblnSave  As Boolean                 '�ж��Ƿ��Ѿ�����

Private Enum �ƶ���ʽ
    ���� = 0
    ���� = 1
End Enum

Private Enum ��ǰѡ��
    ���� = 0
    ���� = 1
    ��ѯ = 2
    ������Ϣ = 3
End Enum

Public pblnInUse As Boolean



'���ܣ���Ӧ�û�ѡ���ƽ̨����
'���룺�ƶ���ʽ
'�������
'���أ���
'ע�����
'���ߣ�������
'����ʱ�䣺2001-3-12
Private Sub ccmd�ƶ�_Click(Index As Integer)
    On Error GoTo errHandle
    Dim lnodeTemp As Node
    Dim lobj���� As Collection
    Dim lstrChild As String
    Dim lstr���� As String
    Dim lintCount As Integer
    Dim lstr������ As String
    Dim i As Integer
    Dim ii As Integer
    Dim llngChildren As Long '�ӽڵ�����
    Dim llngIndex As Long
    
    mblnSave = False
    
    Select Case Index
        Case ����                         '����
            If ctvƽ̨��.SelectedItem Is Nothing Then Exit Sub
            
            '�ж��ܷ����ơ�
            If func�ڵ�λ��(ctvƽ̨��.SelectedItem.FullPath) = 3 Or func�ڵ�λ��(ctvƽ̨��.SelectedItem.FullPath) = 4 Then
                If ctrwȨ��.SelectedItem.Parent Is Nothing Then
                    'һ���������ӡ�
                    llngIndex = ctrwȨ��.SelectedItem.Child.Index
                    llngChildren = ctrwȨ��.SelectedItem.Children
                    For i = llngIndex To llngIndex + llngChildren - 1
                        lstr������ = ctrwȨ��.Nodes(i).Key
                        If Not funcӵ��(lstr������) Then
                            lstr���� = lstr������                         'ʹ��ȱʡ����
                            '�޸ģ�2001-11-2��ȥ��ҵ����ǰ׺����
                            If InStr(lstr����, "_") > 0 Then
                                lstr���� = Right(lstr����, Len(lstr����) - InStr(lstr����, "_"))
                            End If
                            Call sub����(lstr������, lstr����)                           '�����ƶ�����
                        End If
                    Next
                Else
                    '����������ӡ�
                    lstr������ = ctrwȨ��.SelectedItem.Key
                    '�жϸò����û��Ƿ���ӵ�С�
                    If funcӵ��(lstr������) Then
                        Call sffuncMsg("����ӵ��" & lstr������ & "������", sf����)
                    Else
                        If ccheckȱʡ = 0 Then                            '��ʹ��ȱʡ����
                            lstr���� = InputBox("������ò����ı���", "ϵͳ��ʾ", lstr������)
                            lstr���� = Trim(Replace(lstr����, "'", ""))
                            If lstr���� = "" Then Exit Sub
                            If Len(lstr����) > 15 Then Call sffuncMsg("�����������ܳ���ʮ����ַ�!", sf����): Exit Sub
                        Else
                            lstr���� = lstr������                         'ʹ��ȱʡ����
                            '�޸ģ�2001-11-2��ȥ��ҵ����ǰ׺����
                            If InStr(lstr����, "_") > 0 Then
                                lstr���� = Right(lstr����, Len(lstr����) - InStr(lstr����, "_"))
                            End If
                        End If
                        Call sub����(lstr������, lstr����)                           '�����ƶ�����
                    End If
                End If
            End If
        Case ����
            If func�ڵ�λ��(ctvƽ̨��.SelectedItem.FullPath) = 3 Then
                'ȫ���ơ�
                lintCount = ctvƽ̨��.SelectedItem.Children               '���������еĲ�����
                For ii = 1 To lintCount
                   lstrChild = ctvƽ̨��.SelectedItem.Child.Key
                    If func�ڵ�λ��(ctvƽ̨��.SelectedItem.Child.FullPath) = 4 Then
                        Call subɾ��(lstrChild, "ȫ")
                    End If
                Next ii
            ElseIf func�ڵ�λ��(ctvƽ̨��.SelectedItem.FullPath) = 4 Then    '�ж��ܷ�����
                '�������ơ�
                lstr������ = ctvƽ̨��.SelectedItem.Key   '��ǰѡ�е�ƽ̨���еĲ�������
                Call subɾ��(lstr������, "")
            End If
    End Select
    ctvƽ̨��.Refresh
    
    Exit Sub
errHandle:
    Call sfsub������("ƽ̨����", "frmƽ̨����", "ccmd�ƶ�_Click", Err.Number, Err.Description, False)
End Sub



'���ܣ������û���ѡ�������û��Ĳ�����������
'���룺��
'�������
'���أ���
'ע�����
'���ߣ�������
'����ʱ�䣺2001-3-15
Private Sub subAdd()
    On Error GoTo errHandle
    Dim lstr���� As String      '�����ӵ�����������
    Dim lstrTemp As String      'ȷ�����ӵ������
    Dim lobj�� As New Collection '���ӵļ���
    Dim lnodeTemp As Node
    Select Case func�ڵ�λ��(ctvƽ̨��.SelectedItem.FullPath)  '���ݽڵ��λ��ȷ���������໹��������
        Case 1
            lstrTemp = "�������"
        Case 2
            lstrTemp = "�������"
        Case Else
            sffuncMsg "����ѡ����Ҫ���ӵ�����", sf����
            Exit Sub
    End Select
    lstr���� = InputBox("��������Ҫ����" & lstrTemp, "ϵͳ��ʾ")
    lstr���� = Trim(Replace(lstr����, "'", ""))
    If lstr���� = "" Then Exit Sub '�û�ȡ��
    If Len(lstr����) > 6 Then Call sffuncMsg("���������Ʋ��ܳ��������ַ�!", sf����): Exit Sub
    If IsNumeric(lstr����) Then Call sffuncMsg("���Ʋ���ȫ��������!", sf����): Exit Sub
    If IsDate(lstr����) Then Call sffuncMsg("���Ʋ�����������ʽ!", sf����): Exit Sub
    If funcInOperation(lstr����) Then sffuncMsg "���������Ʋ�����ϵͳ���еĲ�������! ", sf����: Exit Sub
    Set lnodeTemp = ctvƽ̨��.Nodes.Add(ctvƽ̨��.SelectedItem.Key, 4, lstr����, lstr����, "First")   'ctvƽ̨�����ӽڵ�
    If lstrTemp = "�������" Then          '���Ӳ�����
        lobj��.Add ctvƽ̨��.SelectedItem.Text, "��������"
        lobj��.Add lstr����, "������"
    Else                                 '���Ӳ�����
        lobj��.Add lstr����, "��������"
        lobj��.Add CStr(Now), "������"
    End If
    mobj��.Add lobj��
    mblnSave = False
    Exit Sub
errHandle:
    If Err.Number = 35602 Then
        Err.Number = 6666
        Err.Description = "�������������Ѿ����ڣ��뻻�����ƣ�"
    End If
    Call sfsub������("������", "frmƽ̨����", "form_Load", Err.Number, Err.Description, False)
End Sub



'���ܣ������û���ѡ��ɾ���û��Ĳ�����������
'���룺��
'�������
'���أ���
'ע�����
'���ߣ�������
'����ʱ�䣺2001-3-15
Private Sub subDelete()
    On Error GoTo errHandle
    Dim lstrTemp As String
    Dim i As Integer
    Select Case func�ڵ�λ��(ctvƽ̨��.SelectedItem.FullPath)  '�жϸýڵ��ܷ�ɾ��
        Case 1
            Call sffuncMsg("����㲻����ɾ����", sf����)         '����㲻����ɾ��
        Case 2                      'ɾ�� ������
            If ctvƽ̨��.SelectedItem.Children > 0 Then          '�ý�������ӽ�㲻����ɾ��
                Call sffuncMsg("����ɾ���ý����ӽ�㣡", sf����)
            Else
                For i = 1 To mobj��.Count
                    If mobj��(i)("��������") = ctvƽ̨��.SelectedItem.Text Then
                        mobj��.Remove (i)
                        Exit For
                    End If
                Next i
                ctvƽ̨��.Nodes.Remove (ctvƽ̨��.SelectedItem.Key)
            End If
               mblnSave = False
        Case 3        'ɾ�� ������
            lstrTemp = func�ܷ�ɾ��(ctvƽ̨��.SelectedItem.Key)
            If Not lstrTemp = "" Then
                Call sffuncMsg("����ɾ���ý���" & lstrTemp & "�е��ӽ�㣡", sf����)
            Else
                For i = 1 To mobj��.Count
                    If mobj��(i)("������") = ctvƽ̨��.SelectedItem.Text Then
                        mobj��.Remove (i)
                        ctvƽ̨��.Nodes.Remove (ctvƽ̨��.SelectedItem.Key)
                        Exit For
                    End If
                Next i
            End If
             mblnSave = False
        Case 4
            mblnSave = False
            Call subɾ��(ctvƽ̨��.SelectedItem.Key, "")         'ɾ��Ȩ��
    End Select
    Exit Sub
errHandle:
    Call sfsub������("������", "frmƽ̨����", "subDelete", Err.Number, Err.Description, False)
End Sub

Private Sub coptѡ��_Click(Index As Integer)
    On Error Resume Next
    mint��ǰѡ�� = Index
    subLoadƽ̨�� (Mid(coptѡ��.Item(Index).Caption, 1, 2))
    subLoadȨ���� (Mid(coptѡ��.Item(Index).Caption, 1, 2))
    Set ctvƽ̨��.SelectedItem = ctvƽ̨��.Nodes(1)
End Sub

Private Sub ctrwȨ��_Click()
    ctvƽ̨��_Click
End Sub

'��Ӧ�û���˫������
Private Sub ctrwȨ��_DblClick()
    On Error Resume Next
    If ccmd�ƶ�(����).Enabled = True And Not ctrwȨ��.SelectedItem Is Nothing Then
        'ֻ�е������ͨ��˫�����ơ�
        If Not ctrwȨ��.SelectedItem.Parent Is Nothing Then
            Call ccmd�ƶ�_Click(����)
        End If
    End If
End Sub

'ȷ���ƶ��İ�Ŧ��ʱ����
Private Sub ctvƽ̨��_Click()
    On Error Resume Next
    ccmd�ƶ�(����).Enabled = False
    ccmd�ƶ�(����).Enabled = False
    Select Case func�ڵ�λ��(ctvƽ̨��.SelectedItem.FullPath)
        Case 3, 4
            ccmd�ƶ�(����).Enabled = True
            If Not ctrwȨ��.SelectedItem Is Nothing Then
                ccmd�ƶ�(����).Enabled = True
            End If
    End Select
'    clblInfo(0).Visible = Not ccmd�ƶ�(����).Enabled
'    clblInfo(1).Visible = Not ccmd�ƶ�(����).Enabled
End Sub

Private Sub ctvƽ̨��_DblClick()
    '˫������ȥ��Ȩ�ޡ�
    On Error Resume Next
    If ccmd�ƶ�(����).Enabled = True And Not ctvƽ̨��.SelectedItem Is Nothing Then
        'ֻ�е������ͨ��˫�����ơ�
        If func�ڵ�λ��(ctvƽ̨��.SelectedItem.FullPath) = 4 Then
            Call ccmd�ƶ�_Click(����)
        End If
    End If
    
End Sub

Private Sub ctvƽ̨��_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo errHandle
    If Button = 2 Then
        PopupMenu ƽ̨, vbPopupMenuCenterAlign
    End If
    Exit Sub
errHandle:
    Call sfsub������("ƽ̨����", "frmƽ̨����", "ctvƽ̨��_MouseUp", Err.Number, Err.Description, False)
End Sub


Private Sub Form_Load()
    On Error GoTo errHandle
    'dasubInitialize ("driver=sql server;server=wangxiaohua;database=����26ϵͳ�������ݿ�;uid=user26;pwd=welcome")
    If pblnInUse Then Exit Sub
    pblnInUse = True
    Dim llng As Long
    llng = GetWindowLong(Me.hWnd, GWL_STYLE)
    
    '���ر�������
'    If (llng And WS_BORDER) = WS_BORDER Then
'        llng = llng - WS_BORDER
'    End If
    SetWindowLong Me.hWnd, GWL_STYLE, llng
    SetWindowPos Me.hWnd, Me.hWnd, 0, 0, 0, 0, SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME
    
    Dim lcol��������ť As New Collection '�������ϵİ�ť��ʼ�����ϡ�
    Set mobjGUI = New cls����ͨ�ö���
    Set mobjGUI.Form = Me
    Set mobjGUI.c������ = c������
    Set mobjGUI.c״̬�� = c״̬��
    
    lcol��������ť.Add "���"
    lcol��������ť.Add "ɾ��"
    lcol��������ť.Add "�޸�"
    lcol��������ť.Add "|"
    lcol��������ť.Add "����"
    lcol��������ť.Add "|"
    lcol��������ť.Add "�˳�"
    
    mobjGUI.subInitialize lcol��������ť, ""
    pobjƽ̨�ṹ.ƽ̨���� = um�û����
    
    Set mobj�� = funcConvertColl(pobjƽ̨�ṹ.���������)
    Set mobj���� = funcConvertColl(pobjƽ̨�ṹ.Operations)
    Set mobj��ѯ = funcConvertColl(pobjƽ̨�ṹ.Queries)
    Set mobj���� = funcConvertColl(pobjƽ̨�ṹ.Reports)
    Set mobj���ò��� = pobjƽ̨�ṹ.Operation
    
    If um�û���� = "0000" And um����Ȩ��.RecordCount = 0 Then
        On Error Resume Next
        '��ϵͳ����Ա��Ȩ��Ϊ�գ��Զ����Ȩ�ޡ�
        dafuncGetData "insert into ϵͳ����_�û�����Ȩ�ޱ� values('0000','ϵͳ����_�û���Ȩ������')"
        On Error GoTo errHandle
    End If
    
    If um����Ȩ��.RecordCount = 0 Then sffuncMsg "��֪ͨϵͳ����Ա�Ƚ��롰�û�Ȩ�����á��������������Ӳ���Ȩ�ޣ�", sf����
    
    Set mobjȨ�� = funcConvertColl(um����Ȩ��)
    Set mobjSmartInfos = funcConvertColl(pobjƽ̨�ṹ.SmartInfos)
    subLoadƽ̨�� ("")       '��ʼ�������ϵ�ƽ̨��Ϣ
    subLoadȨ���� ("����")   '��ʼ�������ϵĿ��ò�����Ϣ
    mint��ǰѡ�� = 0
    Call coptѡ��_Click(mint��ǰѡ��)
    
    '�޸ģ�2001-8-24���������ϵͳ����Ա����������������Ϣ����
    If um�û���� = "0000" Then
        coptѡ��(3).Visible = False
    Else
        coptѡ��(3).Visible = True
    End If
    
    mblnSave = True
    c״̬��.Panels(1).Text = "��ע������������������ʱ��Ҫ�����в���ͬ�����Ƽ��������������ֺ����ֽ�β��"
    Exit Sub
errHandle:
    Call sfsub������("ƽ̨����", "frmƽ̨����", "form_Load", Err.Number, Err.Description, False)
End Sub


'���ܣ������û�ѡ������ͣ���ʾ��ͬ���û�ƽ̨��Ϣ
'���룺��������
'�������
'���أ���
'ע�����
'���ߣ�������
'����ʱ�䣺2001-3-12
Private Sub subLoadƽ̨��(ByVal para���� As String)
    On Error GoTo errHandle
    Dim lobj�� As Object
    Dim lnodeTemp As Node
    Dim lstrTemp As String
    Dim i As Integer
    
    ctvƽ̨��.Nodes.Clear        '���ԭ�е�ƽ̨��Ϣ
    ctvƽ̨��.Nodes.Add , , "ƽ̨�ṹ", "��ǰ�û���ƽ̨�ṹ", "First" '�����ʼֵ
    If mobj��.Count > 0 Then      '�����������
        For i = 1 To mobj��.Count
            Set lnodeTemp = ctvƽ̨��.Nodes.Add("ƽ̨�ṹ", 4, mobj��(i)("��������"), mobj��(i)("��������"), "First")
        Next i
        '�����������
        For i = 1 To mobj��.Count
            lstrTemp = mobj��(i)("��������")
            If Not IsDate(mobj��(i)("������")) Then
                Set lnodeTemp = ctvƽ̨��.Nodes.Add(lstrTemp, 4, mobj��(i)("������"), mobj��(i)("������"), "First")
            End If
        Next i
    
        Select Case para����   '���ݲ�ͬ��������ʾ��ͬ��ƽ̨��Ϣ
            Case "����"
                Set lobj�� = mobj����       '������Ϣ
            Case "��ѯ"
                Set lobj�� = mobj��ѯ      '��ѯ��Ϣ
            Case "����"
                Set lobj�� = mobjSmartInfos '������Ϣ
            Case Else
                Set lobj�� = mobj����       '������Ϣ
        End Select
        If lobj��.Count > 0 Then    '����Ϣ����TreeView
            For i = 1 To lobj��.Count
                lstrTemp = lobj��(i)("��������")
                '�޸ģ�2003-7-9������жϵ�ǰ��������ҵ�����Ƿ��ڼ��ܹ���ɷ�Χ�ڡ�
                mobj���ò���.Filter = ""
                mobj���ò���.Filter = "��������" & "='" & lobj��(i)("��������") & "'"
                If mobj���ò���.RecordCount > 0 Then
                    If pstr��ϵͳ��� = "" Or InStr(pstr��ϵͳ���, mobj���ò���.Fields("ҵ����") & ",") > 0 Then
                        Set lnodeTemp = ctvƽ̨��.Nodes.Add(lstrTemp, 4, lobj��(i)("��������"), lobj��(i)("��������") & "(������" & lobj��(i)("��������") & ")", "Second")
                    End If
                End If
            Next i
        End If
    End If
    Exit Sub
errHandle:
    If Err.Number = 35602 Then
        Resume Next
    Else
        Call sfsub������("������", "frmƽ̨����", "subloadƽ̨��", Err.Number, Err.Description, False)
    End If
End Sub


'���ܣ������û�ѡ������ͣ���ʾ��ͬ���û��ɼ�����Ϣ
'���룺��������
'�������
'���أ���
'ע�����
'���ߣ�������
'����ʱ�䣺2001-3-12
'�޸ģ�2001-11-2�������Ȩ������Ϊ��״��
Private Sub subLoadȨ����(ByVal para���� As String)
    On Error GoTo errHandle
    Dim lobjȨ�� As Object
    Dim lobjһ������ As Object
    Dim i As Integer
    Dim lobjList As ListItem
    Dim lstrҵ���� As String
    
    ctrwȨ��.Nodes.Clear
    Select Case para����
        Case "����"    '�ɼ�����
            
        Case "��ѯ"
        Case "����"   '�ɼ���������Ϣ
            Set lobjȨ�� = dafuncGetData("select ������,ҵ���� from ϵͳ����_ҵ��������Ϣ��")
            If lobjȨ��.RecordCount > 0 Then
                lobjȨ��.MoveFirst
                For i = 1 To lobjȨ��.RecordCount
                    On Error Resume Next
                    '����ҵ�������ڵ㡣
                    ctrwȨ��.Nodes.Add , , lobjȨ��("ҵ����"), lobjȨ��("ҵ����")
                    '��������ӽڵ㡣
                    ctrwȨ��.Nodes.Add lobjȨ��("ҵ����").Value, tvwChild, lobjȨ��("������"), lobjȨ��("������")
                    On Error GoTo errHandle
                    lobjȨ��.MoveNext
                Next i
            End If
        Case Else   '���õĲ���Ȩ��
            Set lobjһ������ = pobjƽ̨�ṹ.һ������
            If mobjȨ��.Count > 0 And lobjһ������.RecordCount > 0 Then
                For i = 1 To mobjȨ��.Count
                    '�ж��Ƿ���һ������Ȩ�ޡ�
                    lobjһ������.Filter = "������" & "= '" & mobjȨ��(i)("Ȩ����") & "'"
                    If lobjһ������.RecordCount > 0 Then
                        '�޸ģ�2003-7-9������жϵ�ǰ��������ҵ�����Ƿ��ڼ��ܹ���ɷ�Χ�ڡ�
                        If pstr��ϵͳ��� = "" Or InStr(pstr��ϵͳ���, lobjһ������("ҵ����") & ",") > 0 Then
                            On Error Resume Next
                            '����ҵ�������ڵ㡣
                            ctrwȨ��.Nodes.Add , , lobjһ������("ҵ����"), lobjһ������("ҵ����")
                            '��������ӽڵ㡣
                            ctrwȨ��.Nodes.Add lobjһ������("ҵ����").Value, tvwChild, mobjȨ��(i)("Ȩ����"), mobjȨ��(i)("Ȩ����")
                            On Error GoTo errHandle
                        End If
                   End If
                   lobjһ������.Filter = ""
                Next i
           Else
           End If
    End Select
Exit Sub
errHandle:
    Call sfsub������("������", "frmƽ̨����", "subloadȨ����", Err.Number, Err.Description, False)
End Sub


'���ܣ������û�ѡ��Ĳ������ж��Ƿ�ӵ��
'���룺����
'�������
'���أ�True ��ӵ�У�False��û��
'ע�����
'���ߣ�������
'����ʱ�䣺2001-3-12
Private Function funcӵ��(ByVal para���� As String) As Boolean
    On Error GoTo errHandle
    Dim i As Integer
    funcӵ�� = False
    For i = 1 To ctvƽ̨��.Nodes.Count    '����ƽ̨��
       ' MsgBox ctvƽ̨��.Nodes.Item(i).Key
    If para���� = ctvƽ̨��.Nodes.Item(i).Key Then
            funcӵ�� = True              'ӵ��
            Exit For
        End If
    Next i
    Exit Function
errHandle:
    Call sfsub������("������", "frmƽ̨����", "funcӵ��", Err.Number, Err.Description, False)
End Function


'���ܣ������û�ѡ��Ĳ������жϸò�������ƽ̨���ĵڼ���
'���룺���Ƶ�����·��
'�������
'���أ��������㡡������㡡������� ����������
'ע�����
'���ߣ�������
'����ʱ�䣺2001-3-12
Private Function func�ڵ�λ��(ByVal paraPath As String) As Integer
    On Error GoTo errHandle
    Dim lint As Integer
    func�ڵ�λ�� = 1
    lint = InStr(1, paraPath, "\", vbTextCompare) '�ж�·����"\"������
    Do While lint > 0
        lint = InStr(lint + 1, paraPath, "\", vbTextCompare)
        func�ڵ�λ�� = func�ڵ�λ�� + 1
    Loop
Exit Function
errHandle:
    Call sfsub������("������", "frmƽ̨����", "func�ڵ�λ��", Err.Number, Err.Description, False)
End Function


'���ܣ������û�ѡ��Ĳ��������ò����Ƶ�ƽ̨����
'���룺��������
'�������
'���أ���
'ע�����
'���ߣ�������
'����ʱ�䣺2001-3-12
'�޸ģ�2001-11-2�������Ӳ�����para����������
Private Sub sub����(ByVal para������ As String, ByVal para���� As String)
    On Error GoTo errHandle
    Dim i As Integer
    Dim lobj���� As New Collection
    Dim lstrTemp As String
    Dim lnodeTemp As Node
    
    If func�ڵ�λ��(ctvƽ̨��.SelectedItem.FullPath) = 3 Then
        Set lnodeTemp = ctvƽ̨��.SelectedItem
    Else
        Set lnodeTemp = ctvƽ̨��.SelectedItem.Parent
    End If
    
    lstrTemp = lnodeTemp.Key
    lobj����.Add lstrTemp, "��������"
    lobj����.Add para������, "��������"
    lobj����.Add para����, "��������"
    Select Case mint��ǰѡ��
        Case ����
            mobj����.Add lobj����
        Case ��ѯ
            mobj��ѯ.Add lobj����
        Case ����
            mobj����.Add lobj����
        Case ������Ϣ
            mobjSmartInfos.Add lobj����
    End Select
    Set lnodeTemp = ctvƽ̨��.Nodes.Add(lstrTemp, 4, para������, para������ & "(������" & para���� & ")", "Second")
    
    Exit Sub
errHandle:
    Call sfsub������("������", "frmƽ̨����", "sub����", Err.Number, Err.Description, False)
End Sub

'���ܣ������û�ѡ��Ĳ��������ò�����ƽ̨����ɾ��
'���룺��������
'�������
'���أ���
'ע�����
'���ߣ�������
'����ʱ�䣺2001-3-12
'�޸ģ�2001-11-2�����ȥ������Ȩ�����е�ӵ���ֶΡ�
Private Sub subɾ��(ByVal paraName As String, ByVal para�Ƴ���ʽ As String)
    On Error GoTo errHandle
    Dim i As Integer
    Dim j As Integer
    mblnSave = False
    If para�Ƴ���ʽ = "ȫ" Then
        ctvƽ̨��.Nodes.Remove (ctvƽ̨��.SelectedItem.Child.Index)   '�Ƴ�
    Else
        ctvƽ̨��.Nodes.Remove (ctvƽ̨��.SelectedItem.Index)   '�Ƴ�
    End If
    Select Case mint��ǰѡ��
    Case ����
        For j = 1 To mobj����.Count
            If mobj����(j)("��������") = paraName Then
                mobj����.Remove (j)
                Exit For
            End If
        Next j
    Case ��ѯ
        For j = 1 To mobj��ѯ.Count
            If mobj��ѯ(j)("��������") = paraName Then
                mobj��ѯ.Remove (j)
                Exit For
            End If
        Next j
    Case ����
        For j = 1 To mobj����.Count
            If mobj����(j)("��������") = paraName Then
                mobj����.Remove (j)
                Exit For
            End If
        Next j
    Case ������Ϣ
        For j = 1 To mobjSmartInfos.Count
            If mobjSmartInfos(j)("��������") = paraName Then
                mobjSmartInfos.Remove (j)
                Exit For
            End If
        Next j
End Select
    ctvƽ̨��.Refresh                                      'ˢ��ƽ̨��
    Exit Sub
errHandle:
    Call sfsub������("������", "frmƽ̨����", "subɾ��", Err.Number, Err.Description, False)
End Sub

'���ܣ�����¼��ת���ɼ���
'���룺��¼��
'�������
'���أ�����
'ע�����
'���ߣ�������
'����ʱ�䣺2001-3-12
Private Function funcConvertColl(ByVal paraRes As Object) As Object
    On Error GoTo errHandle
    Dim i As Integer  'ѭ������
    Dim lintTemp As Integer
    Dim lobjTemp As Collection  '����
    Dim lstr��Ŀ As String      '��Ŀ����
    Dim lstr�ֶ��� As String    '�ֶ�����
    Set funcConvertColl = New Collection
    If paraRes.RecordCount > 0 Then      '��¼���ļ�¼�����������
        paraRes.MoveFirst
        For i = 1 To paraRes.RecordCount
            Set lobjTemp = New Collection
            lintTemp = paraRes.Fields.Count
            Do While lintTemp > 0
                lstr��Ŀ = paraRes.Fields(lintTemp - 1)
                lstr�ֶ��� = paraRes.Fields(lintTemp - 1).Name
                lobjTemp.Add lstr��Ŀ, lstr�ֶ���
                'lobjTemp.Add paraRes.Fields("������"), "������"
                lintTemp = lintTemp - 1
            Loop
            paraRes.MoveNext
            funcConvertColl.Add lobjTemp, CStr(i)
        Next i
    End If
    Exit Function
errHandle:
    Call sfsub������("������", "frmƽ̨����", "funcConvertColl", Err.Number, Err.Description, False)
End Function


'���ܣ������û���ƽ̨���޸�
'���룺��
'�������
'���أ���
'ע�����
'���ߣ�������
'����ʱ�䣺2001-3-12
Private Sub subSave()
    On Error GoTo errHandle
    '��ƽ̨�����鱣��
    pobjƽ̨�ṹ.��������� = mobj��   'ͬʱɾ��ƽ̨��ǰ�Ĳ��������á�
    pobjƽ̨�ṹ.Operations = mobj���� '�������鸳ֵ
    pobjƽ̨�ṹ.Queries = mobj��ѯ    '����ѯ��ֵ
    pobjƽ̨�ṹ.Reports = mobj����    '������ֵ
    pobjƽ̨�ṹ.SmartInfos = mobjSmartInfos '��������Ϣ��ֵ
    pobjƽ̨�ṹ.funcSaveSetupOP       '�����������
    pobjƽ̨�ṹ.funcSaveSetupRP       '���汨������
    pobjƽ̨�ṹ.funcSaveSetupQE       '�����ѯ����
    pobjƽ̨�ṹ.funcSaveSetupSI       '����������Ϣ����
    sffuncMsg "Ҫʹ��ǰ������Ч����ע�����µ�¼ϵͳ��", sf����
    mblnSave = True
    Exit Sub
errHandle:
    Call sfsub������("������", "frmƽ̨����", "subSave", Err.Number, Err.Description, False)
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    Set mobjGUI.Form = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set mobjȨ�� = Nothing
    Set mobjGUI = Nothing
    pblnInUse = False
End Sub

'����������
Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    On Error GoTo errHandle
    Select Case Operate
        Case "����"
            subSave

            Cancel = True
        Case "ɾ��"
            subDelete
            Cancel = True
        Case "���"
            subAdd
            Cancel = True
        Case "�޸�"
            subModify
            Cancel = True
        Case "�˳�"
            If mblnSave = False Then
                If sffuncMsg("��ǰ�����ĸĶ��Ƿ񱣴棿", sfѯ��) Then
                    subSave
                End If
            End If
    End Select
    Exit Sub
errHandle:
    Call sfsub������("������", "frmƽ̨����", "mobjGUI_BeforeOperate", Err.Number, Err.Description, False)
End Sub


'���ܣ��жϸ����ܷ�ɾ��
'���룺����
'�������
'���أ��������ɾ���򷵻ؿմ������򷵻���Ӧ����ʾ��Ϣ
'ע�����
'���ߣ�������
'����ʱ�䣺2001-3-12
Private Function func�ܷ�ɾ��(ByVal para���� As String) As String
    On Error GoTo errHandle
    Dim i As Integer
    For i = 1 To mobj����.Count        '�жϱ������Ƿ��и���
        If mobj����(i)("��������") = para���� Then
            func�ܷ�ɾ�� = "������Ϣ"
            Exit Function
        End If
    Next i
    For i = 1 To mobj��ѯ.Count       '�жϲ�ѯ���Ƿ��и���
        If mobj��ѯ(i)("��������") = para���� Then
            func�ܷ�ɾ�� = "��ѯ��Ϣ"
            Exit Function
        End If
    Next i
    For i = 1 To mobjSmartInfos.Count     '�ж�������Ϣ���Ƿ��и���
        If mobjSmartInfos(i)("��������") = para���� Then
            func�ܷ�ɾ�� = "������Ϣ"
            Exit Function
        End If
    Next i
    For i = 1 To mobj����.Count         ' '�жϲ������Ƿ��и���
        If mobj����(i)("��������") = para���� Then
            func�ܷ�ɾ�� = "������Ϣ"
            Exit Function
        End If
    Next i
    func�ܷ�ɾ�� = ""
    Exit Function
errHandle:
Call sfsub������("������", "frmƽ̨����", "func�ܷ�ɾ��", Err.Number, Err.Description, True)
End Function



'���ܣ��ж��������������Ƿ�����в���ͬ��
'���룺����������
'�������
'���أ�True������ͬ�� False��������ͬ��
'ע�����
'���ߣ�������
'����ʱ�䣺2001-3-12
'�޸ģ�2001-11-2�������
Private Function funcInOperation(ByVal para���� As String) As Boolean
    Dim lobjNode As Node
    
    On Error GoTo errHandle
    
    funcInOperation = True
    For Each lobjNode In ctrwȨ��.Nodes
        If Not lobjNode.Parent Is Nothing Then
            If para���� = lobjNode.Key Then
                funcInOperation = True
                Exit Function
            End If
        End If
    Next
    funcInOperation = False
    Exit Function
errHandle:
Call sfsub������("������", "frmƽ̨����", "funcinoperation", Err.Number, Err.Description, True)
End Function



'���ܣ��޸��������������������
'���룺��
'�������
'���أ���
'ע�����
'���ߣ�������
'����ʱ�䣺2001-4-16
Private Sub subModify()
    On Error GoTo errHandle
    Dim lstr������ As String      '�޸ĵ�����������
    Dim lstr���� As String      '�޸ĵ�����������
    Dim lstrTemp As String      'ȷ���޸ĵ������
    Dim lobj�� As New Collection '�޸ĵļ���
    Dim i As Integer
    Dim lnodeTemp As Node
    Select Case func�ڵ�λ��(ctvƽ̨��.SelectedItem.FullPath)  '���ݽڵ��λ��ȷ���������໹��������
        Case 2
            lstrTemp = "����"
            lstr������ = ctvƽ̨��.SelectedItem.Text
            
        Case 3
            lstrTemp = "����"
            lstr������ = ctvƽ̨��.SelectedItem.Text
            
        Case 4
            lstrTemp = "����"
            lstr������ = ctvƽ̨��.SelectedItem.Text
            lstr������ = Mid(lstr������, Len(ctvƽ̨��.SelectedItem.Key) + 5)
            lstr������ = Left(lstr������, Len(lstr������) - 1)     '����������ȡ��
        Case Else
            sffuncMsg "����ѡ����Ҫ�޸ĵ��ࡢ����������", sf����
            Exit Sub
    End Select
    lstr���� = InputBox("������" & lstr������ & "�µ�����", "ϵͳ��ʾ", lstr������)
    lstr���� = Trim(Replace(lstr����, "'", ""))
    If IsDate(lstr����) Then Call sffuncMsg("���Ʋ�����������ʽ!", sf����): Exit Sub
    If lstr���� = "" Or lstr������ = lstr���� Then Exit Sub '�û�ȡ��
    If IsNumeric(lstr����) Then Call sffuncMsg("���Ʋ���ȫ��������!", sf����): Exit Sub
    If lstrTemp <> "����" Then
        If Len(lstr����) > 6 Then Call sffuncMsg("���������Ʋ��ܳ��������ַ�!", sf����): Exit Sub
        If funcInOperation(lstr����) Then sffuncMsg "���������Ʋ�����ϵͳ���еĲ�������!", sf����: Exit Sub
        For i = 1 To mobj��.Count
            If lstr���� = mobj��.Item(i)("��������") Then  '���ƴ���
                Err.Raise 6666, , "��ĵ��������Ѿ����ڣ��뻻�����ƣ�"
                Exit For
            End If
        Next i
        For i = 1 To mobj��.Count
            If lstr���� = mobj��.Item(i)("������") Then
                Err.Raise 6666, , "��ĵ��������Ѿ����ڣ��뻻�����ƣ�"
                Exit For
            End If
        Next i
    Else
        If Len(lstr����) > 15 Then Call sffuncMsg("�����������ܳ���ʮ����ַ�!", sf����): Exit Sub
    End If
    If func�޸�����(lstr������, lstr����, lstrTemp) Then
        If lstrTemp <> "����" Then
            ctvƽ̨��.SelectedItem.Text = lstr����  '���滻��
            ctvƽ̨��.SelectedItem.Key = lstr����
        Else
            ctvƽ̨��.SelectedItem.Text = ctvƽ̨��.SelectedItem.Key & "(������" & lstr���� & ")"
        End If
    End If
    mblnSave = False
    Exit Sub
errHandle:
    Call sfsub������("������", "frmƽ̨����", "subModify", Err.Number, Err.Description, False)
End Sub

'���ܣ��޸���������������
'���룺�����ƣ������ƣ��޸�����
'�������
'���أ�True���޸ĳɹ���False���޸�ʧ��
'ע�����
'���ߣ�������
'����ʱ�䣺2001-4-16
Private Function func�޸�����(ByVal para������ As String, ByVal para������ As String, ByVal para���� As String) As Boolean
    On Error GoTo errHandle
    Dim i As Integer
    Dim lstr������ As String
    Dim lstr���� As String
    Dim lstr������ As String
    Dim lstr���� As String
    Dim lstr�� As String
    Dim lstr���� As String
    Dim lstrTemp As String
    Dim lobjTemp As Collection
    Dim lobj�� As Collection
    Select Case para����    '�жϸ��ĵ�������������������
        Case "����"     '����
           For i = 1 To mobj��.Count
                If mobj��.Item(i)("��������") = para������ Then
                    lstr������ = mobj��.Item(i)("������")
                    Set lobjTemp = New Collection
                    mobj��.Remove (i)         'ɾ��ԭ�е�
                    lobjTemp.Add para������, "��������"
                    lobjTemp.Add lstr������, "������"
                    mobj��.Add lobjTemp
                    i = 0
                End If
            Next i
            func�޸����� = True
           Exit Function
        Case "����"     '������������ɾ������
            lstrTemp = ctvƽ̨��.SelectedItem.Key
            For i = 1 To mobj��.Count
                If mobj��.Item(i)("������") = para������ Then
                    lstr�� = mobj��.Item(i)("��������")
                    Set lobjTemp = New Collection
                    mobj��.Remove (i)
                    lobjTemp.Add lstr��, "��������"
                    lobjTemp.Add para������, "������"
                    mobj��.Add lobjTemp
                    Exit For
                End If
            Next i
'            Set lobj�� = mobj����
        Case "����"
            lstrTemp = ctvƽ̨��.SelectedItem.Key
    End Select
    For i = 1 To 4
        func��Ծ���������޸� para������, para������, para����, i
    Next i
    func�޸����� = True
    Exit Function
errHandle:
    func�޸����� = False
    Call sfsub������("������", "frmƽ̨����", "func�޸�����", Err.Number, Err.Description, True)
End Function

'���ܣ����ݴ��������������޸Ĳ����е��飬�����е��飬������Ϣ�е���
'���룺�����ƣ������ƣ��޸ĵ����ͣ�����
'�������
'���أ���
'ע�����
'���ߣ�������
'����ʱ�䣺2001-4-16
Private Sub func��Ծ���������޸�(ByVal para������ As String, ByVal para������ As String, ByVal para���� As String, ByVal paraInt�� As Integer)
    On Error GoTo errHandle
    Dim lobj�� As Object
    Dim lobjTemp As Object
    Dim lstr������ As String
    Dim lstr���� As String
    Dim lstr���� As String
    Dim lstrTemp As String
    Dim i As Integer
    Select Case paraInt��
        Case 1
        Set lobj�� = mobj����
        Case 2
        Set lobj�� = mobj����
        Case 3
        Set lobj�� = mobj��ѯ
        Case 4
        Set lobj�� = mobjSmartInfos
    End Select
    If para���� = "����" Then '��������
        For i = 1 To lobj��.Count
            If lobj��.Item(i)("��������") = para������ Then
                lstr���� = lobj��.Item(i)("��������")
                lstr���� = lobj��.Item(i)("��������")
                Set lobjTemp = New Collection
                lobj��.Remove (i)
                lobjTemp.Add para������, "��������"
                lobjTemp.Add lstr����, "��������"
                lobjTemp.Add lstr����, "��������"
                lobj��.Add lobjTemp
                i = 0
            End If
        Next i
    Else         '���ı���
        lstrTemp = ctvƽ̨��.SelectedItem.Key
       For i = 1 To lobj��.Count
          If lobj��.Item(i)("��������") = lstrTemp Then
             lstr������ = lobj��.Item(i)("��������")
            Set lobjTemp = New Collection
            lobj��.Remove (i)
            lobjTemp.Add lstr������, "��������"
            lobjTemp.Add lstrTemp, "��������"
            lobjTemp.Add para������, "��������"
            lobj��.Add lobjTemp
            Exit For
        End If
        Next i
    End If
    Select Case paraInt��
        Case 1
        Set mobj���� = lobj��
        Case 2
        Set mobj���� = lobj��
        Case 3
        Set mobj��ѯ = lobj��
        Case 4
        Set mobjSmartInfos = lobj��
    End Select
    Exit Sub
errHandle:
    Call sfsub������("������", "frmƽ̨����", "func�޸�����", Err.Number, Err.Description, True)
End Sub

Private Sub add_Click(Index As Integer)
    subAdd
End Sub

Private Sub delete_Click(Index As Integer)
    subDelete
End Sub

Private Sub modify_Click(Index As Integer)
    subModify
End Sub
