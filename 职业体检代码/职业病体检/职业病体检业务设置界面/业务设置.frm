VERSION 5.00
Begin VB.Form frmConfigureMedicalExam 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "������ҵ������"
   ClientHeight    =   2850
   ClientLeft      =   1170
   ClientTop       =   870
   ClientWidth     =   7155
   ClipControls    =   0   'False
   Icon            =   "ҵ������.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   7155
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton ccmdAction 
      Caption         =   "�����С����"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   9
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   960
      Width           =   2500
   End
   Begin VB.CommandButton ccmdAction 
      Caption         =   "���¸����޸�Ȩ��"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   8
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "��ӡ��޸ġ�ɾ�������Ա������Ŀ���Ա���������ʱʹ�á�"
      Top             =   2280
      Width           =   2500
   End
   Begin VB.Frame Frame1 
      Caption         =   "����"
      Height          =   735
      Left            =   6000
      TabIndex        =   22
      Top             =   4680
      Visible         =   0   'False
      Width           =   2655
      Begin VB.TextBox ctxtdatenumber 
         Height          =   375
         Left            =   1080
         TabIndex        =   23
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "�������ڣ�"
         Height          =   180
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   180
         Left            =   2280
         TabIndex        =   24
         Top             =   360
         Width           =   180
      End
   End
   Begin VB.CommandButton ccmdAction 
      Caption         =   "���ս���ģ������"
      Height          =   450
      Index           =   7
      Left            =   4080
      TabIndex        =   21
      Top             =   2160
      Visible         =   0   'False
      Width           =   2500
   End
   Begin VB.CommandButton ccmdAction 
      Caption         =   "Ȩ������"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   6
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   360
      Width           =   2500
   End
   Begin VB.CommandButton ccmdAction 
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   5
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1560
      Width           =   2500
   End
   Begin VB.CommandButton ccmdAction 
      Caption         =   "�����Ա������Ŀ����(&O)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   3
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "��ӡ��޸ġ�ɾ�������Ա������Ŀ���Ա���������ʱʹ�á�"
      Top             =   5520
      Visible         =   0   'False
      Width           =   2500
   End
   Begin VB.CommandButton ccmdAction 
      Caption         =   "�����Ŀ����(&I)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   2
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "��ӡ��޸ġ�ɾ�������Ŀ��"
      Top             =   360
      Width           =   2500
   End
   Begin VB.CommandButton ccmdAction 
      Caption         =   "�������ж���������(&F)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   1
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "���ø������۵��Զ��ж�������"
      Top             =   1560
      Visible         =   0   'False
      Width           =   2500
   End
   Begin VB.CommandButton ccmdAction 
      Caption         =   "���ҽʦ����(&D)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   0
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "���ø����ҽʦ���Բ����������Ŀ��"
      Top             =   960
      Width           =   2500
   End
   Begin VB.CommandButton ccmdAction 
      Caption         =   "�� �� (&X)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   4
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1560
      Width           =   2500
   End
   Begin VB.Frame cfraSet 
      Caption         =   "���Ǽ�ʱ�Ƿ��ӡ"
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   3
      Left            =   5040
      TabIndex        =   11
      Top             =   3600
      Visible         =   0   'False
      Width           =   4695
      Begin VB.OptionButton coptPrint 
         Caption         =   "��ӡ����"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   17
         Top             =   420
         Width           =   1215
      End
      Begin VB.OptionButton coptPrint 
         Caption         =   "����ӡ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   3600
         TabIndex        =   7
         Top             =   420
         Width           =   975
      End
      Begin VB.OptionButton coptPrint 
         Caption         =   "��ӡ��쵥"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   6
         Top             =   420
         Width           =   1395
      End
   End
   Begin VB.Frame cfraSet 
      Caption         =   "�Ƿ���ٵǼ�"
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   240
      TabIndex        =   10
      Top             =   4320
      Visible         =   0   'False
      Width           =   4695
      Begin VB.CheckBox chk���ٵǼǼƷ� 
         Caption         =   "���ٵǼǼƷ�"
         Height          =   255
         Left            =   2760
         TabIndex        =   18
         Top             =   240
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.OptionButton coptQuick 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   4
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton coptQuick 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   5
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame cfraSet 
      Caption         =   "�Ƿ�����"
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   5280
      Visible         =   0   'False
      Width           =   4695
      Begin VB.OptionButton coptPhoto 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   2
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton coptPhoto 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   3
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame cfraSet 
      Caption         =   "�Ƿ��շ�"
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   360
      TabIndex        =   8
      Top             =   2880
      Visible         =   0   'False
      Width           =   4695
      Begin VB.OptionButton coptCharge 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
      Begin VB.OptionButton coptCharge 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   0
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmConfigureMedicalExam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'���ܣ�ְҵ������������

Option Explicit

Private mblnInUse As Boolean '������ǰ�����Ƿ��Ѽ��ء�

Private mblnSys As Boolean

Public Property Get pblnInUse() As Boolean
    pblnInUse = mblnInUse
End Property

Private Sub chk���ٵǼǼƷ�_Click()
    On Error GoTo errHandler
    Dim lstrTemp As String 'ҵ������ֵ��
    
    If mblnSys Then Exit Sub
    
    If chk���ٵǼǼƷ�.Value = 1 Then
        lstrTemp = "��"
    Else
        lstrTemp = "��"
    End If
    pobjҵ�����.Sub�޸�ҵ������ "���ٵǼ��Ƿ�Ʒ�", lstrTemp
    
    Exit Sub
errHandler:
    sfsub������ "ְҵ�����ý���", "frmConfigureMedicalExam", "chk���ٵǼǼƷ�_Click", Err.Number, Err.Description, False
    
End Sub



Private Sub Form_Load()
    On Error GoTo errHandler
    
    '���ô����Ѽ��ر�־��
    mblnInUse = True
    
    If um�û���� = "8882" Then
        ccmdAction(9).Top = ccmdAction(6).Top
        ccmdAction(0).Visible = False
        ccmdAction(6).Visible = False
        ccmdAction(8).Visible = False
    End If
    
    '��ʾҵ�����á�
    mblnSys = True
    If pobjҵ�����.ҵ������("�Ƿ��շ�") = "��" Then
        coptCharge(0).Value = True
    Else
        coptCharge(1).Value = True
    End If
    
    If pobjҵ�����.ҵ������("�Ƿ�����") = "��" Then
        coptPhoto(0).Value = True
    Else
        coptPhoto(1).Value = True
    End If

    
    If pobjҵ�����.ҵ������("���ٵǼ��Ƿ�Ʒ�") = "��" Then
        chk���ٵǼǼƷ�.Value = 1
    Else
        chk���ٵǼǼƷ�.Value = 0
    End If
    'chk���ٵǼǼƷ�.Enabled = False
    
    If pobjҵ�����.ҵ������("�Ƿ���ٵǼ�") = "��" Then
        coptQuick(0).Value = True
        If coptCharge(0).Value Then
            'ֻ��Ҫ�շѣ����ҿ��ٵǼ�ʱ��ѡ����Ŀ�ſɲ���
            chk���ٵǼǼƷ�.Enabled = True
        Else
            chk���ٵǼǼƷ�.Value = 0
        End If
    Else
        coptQuick(1).Value = True
        chk���ٵǼǼƷ�.Value = 0
    End If
    
    '����ҵ�����á��Ƿ�ʹ����쵥����
    If pobjҵ�����.ҵ������("�Ƿ�ʹ����쵥") = "��" Then
        coptPrint(0).Visible = False
        coptPrint(2).Left = coptPrint(1).Left
        coptPrint(1).Left = coptPrint(0).Left
    End If
    If pobjҵ�����.ҵ������("�Ƿ��ӡ��쵥") = "��" And pobjҵ�����.ҵ������("�Ƿ�ʹ����쵥") <> "��" Then
        coptPrint(0).Value = True
    ElseIf pobjҵ�����.ҵ������("�Ƿ��ӡ����") = "��" Then
        coptPrint(1).Value = True
    Else
        coptPrint(2).Value = True
    End If
    
    ctxtdatenumber.Text = pobjҵ�����.ҵ������("��������")
    
    mblnSys = False
    
    Exit Sub
errHandler:
    sfsub������ "ְҵ�����ý���", "frmConfigureMedicalExam", "Form_Load", Err.Number, Err.Description, False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '���ô���δ���ر�־��
    mblnInUse = False
End Sub

Private Sub ccmdAction_Click(Index As Integer)
    On Error GoTo errHandler
    Dim lobjFrm As Form
    Select Case Index
    Case 0
        Set lobjFrm = frmSetDoctor
    Case 1
        Set lobjFrm = frmSetConclusionFilter
    Case 2
        Set lobjFrm = frmSetTestItem
    Case 3
        Set lobjFrm = frmSetBaseItem
    Case 5
        Set lobjFrm = frmSetMedicalExamTemplate
    Case 6
        Set lobjFrm = frmSetDoctorPermission
    Case 4
        Unload Me
    Case 7
        Set lobjFrm = frmSetConclusion
    Case 8
        Set lobjFrm = frm���¸����޸�Ȩ��
    Case 9
        Set lobjFrm = frm�����С����
    End Select
    If Not lobjFrm Is Nothing Then
        lobjFrm.Move Me.Left, Me.Top
        lobjFrm.Show 1
    End If
    Set lobjFrm = Nothing
    Exit Sub
errHandler:
    sfsub������ "ְҵ�����ý���", "frmConfigureMedicalExam", "ccmdAction_Click", Err.Number, Err.Description, False
End Sub

Private Sub coptCharge_Click(Index As Integer)
    On Error GoTo errHandler
    Dim lstrTemp As String 'ҵ������ֵ��
    
    If mblnSys Then Exit Sub
    
    If coptCharge(0).Value Then
        lstrTemp = "��"
     '   If coptQuick(0).Value Then
      '      chk���ٵǼǼƷ�.Enabled = True
       ' End If
    Else
        lstrTemp = "��"
        
     '   chk���ٵǼǼƷ�.Enabled = False
     '   chk���ٵǼǼƷ�.Value = 0
        
    End If
    pobjҵ�����.Sub�޸�ҵ������ "�Ƿ��շ�", lstrTemp
    'chk���ٵǼǼƷ�_Click
    Exit Sub
errHandler:
    sfsub������ "ְҵ�����ý���", "frmConfigureMedicalExam", "coptCharge_Click", Err.Number, Err.Description, False
End Sub

Private Sub coptPhoto_Click(Index As Integer)
    Dim lstrTemp As String 'ҵ������ֵ��
    
    On Error GoTo errHandler
    If mblnSys Then Exit Sub
    
    If coptPhoto(0).Value Then
        lstrTemp = "��"
    Else
        lstrTemp = "��"
    End If
    pobjҵ�����.Sub�޸�ҵ������ "�Ƿ�����", lstrTemp

    Exit Sub
errHandler:
    sfsub������ "ְҵ�����ý���", "frmConfigureMedicalExam", "coptPhoto_Click", Err.Number, Err.Description, False
End Sub

Private Sub coptPrint_Click(Index As Integer)
    Dim lstr��쵥  As String 'ҵ������ֵ��
    Dim lstr����  As String 'ҵ������ֵ��
    
    On Error GoTo errHandler
    
    If mblnSys Then Exit Sub
    
    lstr��쵥 = "��"
    lstr���� = "��"
    If coptPrint(0).Value Then
        lstr��쵥 = "��"
    ElseIf coptPrint(1).Value Then
        lstr���� = "��"
    End If
    pobjҵ�����.Sub�޸�ҵ������ "�Ƿ��ӡ��쵥", lstr��쵥
    pobjҵ�����.Sub�޸�ҵ������ "�Ƿ��ӡ����", lstr����
    
    Exit Sub
errHandler:
    sfsub������ "ְҵ�����ý���", "frmConfigureMedicalExam", "coptPrint_Click", Err.Number, Err.Description, False
End Sub

Private Sub coptQuick_Click(Index As Integer)
    On Error GoTo errHandler
    Dim lstrTemp As String 'ҵ������ֵ��
    
    If mblnSys Then Exit Sub
    If coptQuick(0).Value Then
        lstrTemp = "��"
    '    If coptCharge(0).Value Then
    '        chk���ٵǼǼƷ�.Enabled = True
    '    End If
    Else
        lstrTemp = "��"
        
    '    chk���ٵǼǼƷ�.Enabled = False
    '    chk���ٵǼǼƷ�.Value = 0
        
    End If
    pobjҵ�����.Sub�޸�ҵ������ "�Ƿ���ٵǼ�", lstrTemp
    'chk���ٵǼǼƷ�_Click
    Exit Sub
errHandler:
    sfsub������ "ְҵ�����ý���", "frmConfigureMedicalExam", "coptQuick_Click", Err.Number, Err.Description, False
End Sub
Private Sub ctxtDateNumber_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandler
    If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 13 Or KeyAscii = vbKeyBack Then
    Else
        KeyAscii = 0
        Err.Raise 6666, , "�������ڱ����������֡�"
    End If
    
    Exit Sub
errHandler:
    sfsub������ "ְҵ�����ý���", "frmConfigureMedicalExam", "ctxtDateNumber_KeyPress", Err.Number, Err.Description, False
    Exit Sub
    Resume
End Sub


Private Sub ctxtDateNumber_LostFocus()
    On Error GoTo errHandler
   '�жϸ������ڲ���Ϊ��
    If Val(ctxtdatenumber.Text) <= 0 Then
        sffuncMsg "�����븴�����ڡ��������ڱ���>0��", sf����
        ctxtdatenumber.SetFocus
        Exit Sub
    Else
        '�ж��Ƿ�Ϊ����
        If IsNumeric(Trim(ctxtdatenumber.Text)) = False Then
            sffuncMsg "��������ֻ��Ϊ���֣�������¼��", sf����
            With ctxtdatenumber
                .SelStart = 0
                .SelLength = Len(Trim(ctxtdatenumber.Text))
                .SetFocus
            End With
            Exit Sub
        Else
            '�ж��Ƿ����0
            If ctxtdatenumber.Text < 0 Then
                sffuncMsg "�������ڲ���Ϊ����������¼��", sf����
                With ctxtdatenumber
                    .SelStart = 0
                    .SelLength = Len(Trim(ctxtdatenumber.Text))
                    .SetFocus
                End With
            End If
        End If
    End If
    If pobjҵ�����.ҵ������("��������") <> Trim(ctxtdatenumber) Then
        pobjҵ�����.Sub�޸�ҵ������ "��������", Trim(ctxtdatenumber)
    End If
    
    Exit Sub
errHandler:
    sfsub������ "ְҵ�����ý���", "frmConfigureMedicalExam", "ctxtDateNumber_LostFocus", Err.Number, Err.Description, False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errHandler
    '�ȼ�����
    If Shift And vbAltMask = vbAltMask Then
        Select Case KeyCode
        Case vbKeyD
            ccmdAction_Click 0
        Case vbKeyF
            ccmdAction_Click 1
        Case vbKeyI
            ccmdAction_Click 2
        Case vbKeyO
            ccmdAction_Click 3
        Case vbKeyX
            ccmdAction_Click 4
        End Select
    End If
    Exit Sub
errHandler:
    sfsub������ "ְҵ�����ý���", "frmConfigureMedicalExam", "Form_KeyDown", Err.Number, Err.Description, False
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 39 Then
        '���������롰'����
        KeyAscii = 0
    End If

End Sub
