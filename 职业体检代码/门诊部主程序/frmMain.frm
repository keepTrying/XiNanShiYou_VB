VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "�������߹�����Ϣϵͳ"
   ClientHeight    =   11520
   ClientLeft      =   120
   ClientTop       =   15
   ClientWidth     =   15360
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   StartUpPosition =   2  '��Ļ����
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00CEE0AF&
      BorderStyle     =   0  'None
      Height          =   7515
      Left            =   0
      TabIndex        =   9
      Top             =   615
      Width           =   1755
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "�� �� �� ��"
         Height          =   270
         Left            =   450
         TabIndex        =   22
         Top             =   75
         Width           =   1035
      End
      Begin VB.Image Image3 
         Height          =   405
         Left            =   -15
         Picture         =   "frmMain.frx":0CCA
         Stretch         =   -1  'True
         Top             =   -15
         Width           =   1770
      End
      Begin VB.Label cmnuSubItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   585
         TabIndex        =   21
         Top             =   195
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label cmnuSubItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   180
         Index           =   2
         Left            =   540
         TabIndex        =   20
         Top             =   195
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label cmnuSubItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   180
         Index           =   3
         Left            =   525
         TabIndex        =   19
         Top             =   180
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label cmnuSubItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   180
         Index           =   4
         Left            =   570
         TabIndex        =   18
         Top             =   180
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label cmnuSubItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   180
         Index           =   5
         Left            =   555
         TabIndex        =   17
         Top             =   180
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label cmnuSubItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   180
         Index           =   6
         Left            =   540
         TabIndex        =   16
         Top             =   180
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label cmnuSubItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   180
         Index           =   7
         Left            =   630
         TabIndex        =   15
         Top             =   240
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label cmnuSubItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   180
         Index           =   8
         Left            =   660
         TabIndex        =   14
         Top             =   210
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label cmnuSubItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   180
         Index           =   9
         Left            =   585
         TabIndex        =   13
         Top             =   165
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label cmnuSubItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   180
         Index           =   10
         Left            =   630
         TabIndex        =   12
         Top             =   210
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label cmnuItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   165
         TabIndex        =   11
         Top             =   360
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label cmnuSubItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   405
         TabIndex        =   10
         Top             =   825
         Visible         =   0   'False
         Width           =   540
      End
   End
   Begin MSComctlLib.StatusBar cstatusBar 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   6
      Top             =   11160
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8996
            Text            =   "�û���ţ�"
            TextSave        =   "�û���ţ�"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8996
            Text            =   "�û�������"
            TextSave        =   "�û�������"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8996
            Text            =   "����վ����"
            TextSave        =   "����վ����"
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   6000
      Left            =   1800
      Top             =   1200
   End
   Begin VB.Image Image4 
      Height          =   600
      Left            =   30
      Picture         =   "frmMain.frx":1365
      Stretch         =   -1  'True
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����о������������������ල��������Ϣϵͳ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   660
      TabIndex        =   24
      Top             =   150
      Width           =   6300
   End
   Begin VB.Label clblͨ�ò��� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      Height          =   180
      Index           =   0
      Left            =   14445
      TabIndex        =   23
      Top             =   225
      Width           =   360
   End
   Begin VB.Label clblͨ�ò��� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ϵͳ����"
      Enabled         =   0   'False
      Height          =   180
      Index           =   4
      Left            =   10530
      TabIndex        =   8
      Top             =   225
      Width           =   720
   End
   Begin VB.Label clblͨ�ò��� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���ڹ���"
      Enabled         =   0   'False
      Height          =   180
      Index           =   2
      Left            =   8385
      TabIndex        =   7
      Top             =   225
      Width           =   720
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   8190
      Picture         =   "frmMain.frx":2783
      Stretch         =   -1  'True
      Top             =   9960
      Visible         =   0   'False
      Width           =   7035
   End
   Begin VB.Label clblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�û���ţ�           �û����ƣ�           ����վ����"
      Height          =   180
      Left            =   2700
      TabIndex        =   5
      Top             =   7530
      Visible         =   0   'False
      Width           =   4680
   End
   Begin VB.Label clblͨ�ò��� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�˳�"
      Height          =   180
      Index           =   7
      Left            =   13740
      TabIndex        =   4
      Top             =   225
      Width           =   360
   End
   Begin VB.Label clblͨ�ò��� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����޸�"
      Height          =   180
      Index           =   6
      Left            =   12660
      TabIndex        =   3
      Top             =   225
      Width           =   720
   End
   Begin VB.Label clblͨ�ò��� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����֪ͨ"
      Enabled         =   0   'False
      Height          =   180
      Index           =   3
      Left            =   9465
      TabIndex        =   2
      Top             =   225
      Width           =   720
   End
   Begin VB.Label clblͨ�ò��� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��    ѯ"
      Height          =   180
      Index           =   1
      Left            =   7320
      TabIndex        =   1
      Top             =   225
      Width           =   720
   End
   Begin VB.Label clblͨ�ò��� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ֵ����"
      Height          =   180
      Index           =   5
      Left            =   11595
      TabIndex        =   0
      Top             =   225
      Width           =   720
   End
   Begin VB.Image Image1 
      Height          =   6990
      Left            =   60
      Picture         =   "frmMain.frx":14861
      Stretch         =   -1  'True
      Top             =   1695
      Width           =   15135
   End
   Begin VB.Image cimgBackground 
      Height          =   705
      Left            =   15
      Picture         =   "frmMain.frx":24161
      Stretch         =   -1  'True
      Top             =   -75
      Width           =   15225
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mobj�� As Object                '��ǰƽ̨��������
Private mobj�� As Object                '��ǰƽ̨�����в�����
Private mobj���� As Object              '��ǰƽ̨�����в���
Private mobj��ѯ As Object              '��ǰƽ̨�����в�ѯ
Private mobj���� As Object              '��ǰƽ̨�����б���
Private mobj��ѯ���� As Object          '��ǰƽ̨�����в�ѯ��ϸ��Ϣ
Private mobj������� As Object          '��ǰƽ̨�����б�����ϸ��Ϣ
Private mobjSmartInfos As Object
Private mblnRe As Boolean               '�Ƿ�ȷ���˳�
Private mstr��ǰ�� As String            '��ǰѡ���������
Private mblnLoadForm As Boolean         '��ʶ�Ƿ����ڴ������塣
Private mstrOper(1 To 10) As String
Private mstrMnu(1 To 15) As String
Private mintMnu As Integer

'�޸ģ�2002-2-26���������󣩡�
Private X As Object

'�����ѯ��Ҫ�ı�����
'�޸ģ�2001-7-12��
Private mobjFrontQueryManager As Object '��Դ�����ѯ��.clsFrontQueryManager��
Private mobjSysAccObj As Object         '����������.clsSystemAccessObject��

Private Sub cimg��_Click(Index As Integer)
'    Dim i As Long
'    Dim lobjSys As New FileSystemObject
'    Dim llngCount As Long
'    Dim llngTop As Long
'
'    If Index = 0 Then Exit Sub
'    On Error Resume Next
'
'    Unload frm�ֵ��б�
'
'    If lobjSys.FileExists(App.Path & "\image\" & cimg��(Index).Tag & "2.jpg") Then
'        cimg��(Index).Picture = LoadPicture(App.Path & "\image\" & cimg��(Index).Tag & "2.jpg")
'    Else
'        cimg��(Index).Picture = LoadPicture(App.Path & "\image\" & "hot.jpg")
'    End If
'
'    llngCount = cimg��.Count
'
'    For i = 1 To llngCount - 1
'        If i <> Index Then
'            If lobjSys.FileExists(App.Path & "\image\" & cimg��(i).Tag & "1.jpg") Then
'                cimg��(i).Picture = LoadPicture(App.Path & "\image\" & cimg��(i).Tag & "1.jpg")
'            Else
'                cimg��(i).Picture = LoadPicture(App.Path & "\image\" & "normal.jpg")
'
'            End If
'        End If
'    Next
'
'    sub��ʼ�������б� cimg��(Index).Tag
'
'    Set frm�����б�.pfrmParent = Me
'
'    If frm�����б�.clbl����.Count = 2 Then
'        'ֻ��һ��������ֱ�������������档
'        Call sub��������(frm�����б�.clbl����(1).Tag)
'    Else
'        '��ʾ����ѡ���б�
'        frm�����б�.Height = frm�����б�.clbl����.Count * (frm�����б�.clbl����(0).Height + 100) + 200
'        llngTop = cimg��(Index).Top
''        If llngTop + frm�����б�.Height > Me.ScaleHeight - 200 Then
''            llngTop = ScaleHeight - frm�����б�.Height - 200
''        End If
''        If llngTop < 720 Then
''            llngTop = 720
''        End If
'        llngTop = llngTop + Me.Top + 200
'        frm�����б�.Move Me.Left + cimg��(0).Width + 300, llngTop '720
'        frm�����б�.Show , Me
'    End If
'
End Sub

'Public Enum Endway                      '���������ʽ
'    CloseAll = 1                        '�˳�ϵͳ
'    Restart = 2                         '���µ�¼
'End Enum

Private Sub cimgBackground_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    
    For i = 0 To 7
        clblͨ�ò���(i).FontUnderline = False
        clblͨ�ò���(i).ForeColor = vbBlack
    Next
End Sub

Private Sub clblͨ�ò���_Click(Index As Integer)
    On Error GoTo errHandle
    Dim lobjTemp As Object
    Dim lobj���� As Object
    Dim lobj��ѯ As Object
    Dim llngWndProc As Long
    Dim llng������ As Long
    
    clblͨ�ò���(Index).FontUnderline = False
    clblͨ�ò���(Index).ForeColor = vbBlack
    Select Case clblͨ�ò���(Index).Caption
        Case "�ֵ����"
            frm�ֵ��б�.subLoad
            If frm�ֵ��б�.clbl�ֵ�.Count = 1 Then Exit Sub  '���û���ֵ�����Ȩ������Ӧ�û�����
            If frm�ֵ��б�.clbl�ֵ�.Count = 2 Then            '���ֻ��һ���ֵ����õ�Ȩ����ʹ�õ���ʽ�˵�
                Call sub�����ֵ�(frm�ֵ��б�.clbl�ֵ�(1).Caption)
            Else                                   '�ж���ֵ����õ�Ȩ����ʹ�õ���ʽ�˵�
                Unload frm�����б�
                Set frm�ֵ��б�.pobjParent = Me
                frm�ֵ��б�.Move clblͨ�ò���(Index).Left, Me.Top + 800
                frm�ֵ��б�.Show , Me
            End If
         Case "���ڹ���"
                frm���ڹ���.Hide
                Set frm���ڹ���.pobjParent = Me
                frm���ڹ���.Move clblͨ�ò���(Index).Left, Me.Top + 800
                frm���ڹ���.Show , Me
         Case "����֪ͨ"
                frm����֪ͨ.Hide
                Set frm����֪ͨ.pobjParent = Me
                frm����֪ͨ.Move clblͨ�ò���(Index).Left, Me.Top + 800
                frm����֪ͨ.Show , Me
         Case "ϵͳ����"
                frmϵͳ����.Hide
                Set frmϵͳ����.pobjParent = Me
                frmϵͳ����.Move clblͨ�ò���(Index).Left, Me.Top + 800
                frmϵͳ����.Show , Me
           
        Case "��    ѯ"      'ѡȡ��ƽ̨���в�ѯ
                '������ѯ���档
                Dim lobjͨ�ò�ѯ As Object
                Set lobjͨ�ò�ѯ = CreateObject("ͨ�ò�ѯ.clsͨ�ò�ѯ")
                '�޸ģ�2003-7-22�����������ϵͳ��ɲ�����
                llng������ = lobjͨ�ò�ѯ.funcStart("ϵͳ����_ͨ�ò�ѯ", pstr��ϵͳ���)
                
                '�趨�򿪵Ĵ���Ϊ��������Ӵ��塣
                If llng������ <> -2 Then
                    '�򼯺��м����������
                    If Not sffunc�жϼ��ϼ�ֵ�Ƿ����(pcol��������, CStr(llng������)) Then
                        On Error Resume Next
                        SetParent llng������, Me.hWnd
                        llngWndProc = SetWindowLong(llng������, GWL_WNDPROC, AddressOf funcClassing)
                        pcolWndProc.add llngWndProc, CStr(llng������)
                        pcol��������.add "��ѯͳ��", CStr(llng������)
                        pcol�Ӵ�����.add llng������, "��ѯͳ��"
'                        Call MoveWindow(llng������, ScaleX(700, vbTwips, vbPixels), ScaleX(350, vbTwips, vbPixels), ScaleX(Me.ScaleWidth - 700, vbTwips, vbPixels), ScaleX(Me.ScaleHeight - 350, vbTwips, vbPixels), 1)
                        Call MoveWindow(llng������, ScaleX(1750, vbTwips, vbPixels), ScaleX(cimgBackground.Height - 60, vbTwips, vbPixels), ScaleX(Me.ScaleWidth - 1700, vbTwips, vbPixels), ScaleX(Me.ScaleHeight - cstatusBar.Height - cimgBackground.Height + 100, vbTwips, vbPixels), 1)
                        
                        '********************��ʱ���Դ���********************
                        'sfsubSaveSetting "ϵͳ����", "ƽ̨�������", "��ѯͳ�Ƴ�ʼ��", "hWnd��" & llng������ & " WndProc��" & llngWndProc & " ʱ��" & Format(Now, "yyyy��mm��dd��hhʱmm��ss��")
                        '****************************************************
                        Err.Clear
                    End If
                End If
        Case "ͳ�Ʊ���"       'ѡȡ��ƽ̨���б���
'                mobj����.Filter = ""
'                If mobj����.RecordCount = 0 Then Exit Sub
'                mobj�������.Filter = "��������" & "='" & mobj����.Fields("��������") & "'"
                
                '���������ѯ���档
                '�޸ģ�2001-7-13�����
                Call sub���������ѯ
                
        Case "�����޸�"
            frm�����޸�.Show vbModal, Me
        Case "�˳�"
            Unload Me
        Case "����"
            MsgBox "���������У�", vbOKOnly, "����"
        
    End Select
    Exit Sub
errHandle:
    Call sfsub������("������", "frm������", "cSSListBarͨ��_ListItemClick", Err.Number, Err.Description, False)
End Sub


Private Sub clblͨ�ò���_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    clblͨ�ò���(Index).FontUnderline = True
    clblͨ�ò���(Index).ForeColor = &H80FF&
End Sub

Private Sub cmnuItem_Click(Index As Integer)
    Dim ii As Integer, j As Integer
    
    For ii = 1 To Index
        cmnuItem(ii).Top = cmnuItem(ii - 1).Top + cmnuItem(ii - 1).Height + 150
    Next
    cmnuSubItem(0).Top = cmnuItem(ii - 1).Top '+ cmnuItem(ii - 1).Height + 150
    sub��ʼ�������б� mstrMnu(Index)
    For j = 1 To 10
        If cmnuSubItem(j) = "" Then Exit For
    Next
    If ii < mintMnu Then
        cmnuItem(ii).Top = cmnuSubItem(j - 1).Top + cmnuSubItem(j - 1).Height + 150
        For ii = ii + 1 To mintMnu
            cmnuItem(ii).Top = cmnuItem(ii - 1).Top + cmnuItem(ii - 1).Height + 150
        Next
    End If
    cmnuItem(Index).FontUnderline = False
    cmnuItem(Index).ForeColor = vbBlack
End Sub

Private Sub cmnuItem_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmnuItem(Index).FontUnderline = True
    cmnuItem(Index).ForeColor = &H80FF&
End Sub


Private Sub cmnuSubItem_Click(Index As Integer)
    If mstrOper(Index) = "ͳ�Ʊ���" Then
        Call sub���������ѯ
    Else
        sub�������� mstrOper(Index)
    End If
    cmnuSubItem(Index).FontUnderline = False
    cmnuSubItem(Index).ForeColor = vbBlack
End Sub

Private Sub cmnuSubItem_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmnuSubItem(Index).FontUnderline = True
    cmnuSubItem(Index).ForeColor = &H80FF&
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    '��ӦEsc��
    If KeyCode = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer    'ѭ������
    Dim ii As Integer   'ѭ������
    Dim j As Long
    Dim lobjSys As New FileSystemObject
    
    '�޸ģ�2001-11-16�����Ϊ���ڱʼǱ��Ͽ������У�ֻ���������е�������ܣ����������ܡ�
    On Error Resume Next
    Dim lstrServer As String
    lstrServer = sffuncGetSetting("ϵͳ����", "���ݿ�����", "��������")
    
    Me.Caption = um����վ�� & "������Ϣϵͳ"
    
    '��ȡ�������ơ�
    Dim lstrLocalName As String
    lstrLocalName = funcGetLocalName()
    'clblInfo.Caption = "�û���ţ�" & um�û���� & "       �û����ƣ�" & um�û��� & "      ����վ����" & lstrLocalName
    cstatusBar.Panels(1).Text = "�û���ţ�" & um�û����
    cstatusBar.Panels(2).Text = "�û�������" & um�û���
    cstatusBar.Panels(3).Text = "����վ����" & lstrLocalName
    'clblSys.Caption = IIf(pstrSysName = "", "�������߹�����Ϣϵͳ", pstrSysName)
    
    '�޸ģ�2002-3-8��������ע���󣬲���Ҫ�ټ����ܹ�����
    If Not pblnע�� Then
    '�޸ģ�2003-4-17����������ܹ���Ϊ���ô洢���̡�
'    If Not pblnע�� And UCase(Trim(lstrServer)) <> UCase(Trim(lstrLocalName)) Then
'        '��ȡϵͳ�����еļ��ܹ�����������
'        Dim lstrDogServer As String
'        lstrDogServer = sffuncGetSetting("ϵͳ����", "���ݿ�����", "��������������")
'        If lstrDogServer = "" Then lstrDogServer = lstrServer
'
'        Dim lobjRec As Object
'        Call dafuncGetData("sp_addlinkedserver '" & lstrDogServer & "'")
'        Err.Clear
'        Set lobjRec = dafuncGetData("exec " & lstrDogServer & ".master.dbo.ryCheck")
'        If Err.Number = 0 Then
'            Select Case lobjRec(0)
'            Case 1, 2, 3, 4, 5, 6
'                MsgBox "����������װ�����������°�ȫ���ʧ�ܡ�" & Chr(13) & Chr(10) & "�����°�װ�������������", vbCritical, "ϵͳ����"
'                End
'            Case Else
'                If (lobjRec(0) And (&H8000)) = 32768 Then
'                    '�ҵ��ˡ���ȡ��ɡ�
'                    Dim llngBit As Long
'                    Dim larrLic(1 To 10) As String
'                    larrLic(1) = "�������֤����"
'                    larrLic(2) = "�ලִ������"
'                    larrLic(3) = "������"
'                    larrLic(4) = "����֤����,����֤"
'                    larrLic(5) = "����������"
'                    larrLic(6) = "�������"
'                    larrLic(7) = "�ƻ����߹���"
'                    larrLic(8) = "���ڹ���"
'                    larrLic(9) = "�շѹ���"
'                    larrLic(10) = "վ����ѯ"
'                    pstr��ϵͳ��� = ""
'                    llngBit = &H4000
'                    For i = 1 To 10
'                        If (lobjRec(0) And llngBit) = llngBit Then
'                            pstr��ϵͳ��� = pstr��ϵͳ��� & larrLic(i) & ","
'                        End If
'                        llngBit = llngBit / 2
'                    Next
'                Else
'                    MsgBox "������������ʽ��ģ���ȫ���ʧ�ܣ�ϵͳ�޷����С�" & Chr(13) & Chr(10) & "���������Ӧ����ϵ��", vbCritical, "ϵͳ����"
'                    End
'                End If
'            End Select
'        Else
'            MsgBox "����������װ�����������°�ȫ���ʧ�ܡ�" & Chr(13) & Chr(10) & "�����°�װ������������򡣰�װǰȷ�����������������Ѱ�װSql Server2000��", vbCritical, "ϵͳ����"
'            End
'        End If
'
'        pstr��ϵͳ��� = "ϵͳ����,��λ��������," & pstr��ϵͳ���
        sub�����������
        
    End If
    
    'pstr��ϵͳ��� = "ϵͳ����,��λ��������,������,�������֤����,����֤����,����֤,"
    
    '�޸ģ�2003-7-9�������ͨ�ö����ȫ�ֱ����б�����ϵͳ��ɡ�
    um��ϵͳ��� = pstr��ϵͳ���
    
    dasubSetQueryTimeout 6000

    On Error GoTo errHandle
    
    Dim llngWndProc As Long
    llngWndProc = SetWindowLong(Me.hWnd, GWL_WNDPROC, AddressOf funcClassing)
    pcolWndProc.add llngWndProc, CStr(Me.hWnd)
    
    pobjƽ̨�ṹ.ƽ̨���� = um�û���� '��ƽ̨�ṹ���û���Ÿ�ֵ
    Set mobj�� = pobjƽ̨�ṹ.���������  'ȡ���û���ƽ̨�ṹ
    Set mobj�� = pobjƽ̨�ṹ.Operations
    Set mobj���� = pobjƽ̨�ṹ.Operation
    Set mobj��ѯ = pobjƽ̨�ṹ.Queries
    Set mobj��ѯ���� = pobjƽ̨�ṹ.Query
    Set mobj���� = pobjƽ̨�ṹ.Reports
    Set mobj������� = pobjƽ̨�ṹ.Report
    Set mobjSmartInfos = pobjƽ̨�ṹ.SmartInfos
    
    '��ʼ��������ͼƬ
'    mobj��.Filter = "��������= 'ҵ����'"
    j = 1
    '����ͼ���
    Dim lstrFileName As String, lobjPic As StdPicture
    
'    lstrFileName = Dir(App.Path & "\image\*.ico")
'    Do While lstrFileName <> ""
'        Set lobjPic = LoadPicture(App.Path & "\image\" & lstrFileName)
'        cbarOper.IconsLarge.add , Left(lstrFileName, InStr(lstrFileName, ".") - 1), lobjPic
'        lstrFileName = Dir
'    Loop
    
    Dim lcolOper As New Collection, lstrOper(1 To 14) As String, lstrAlias(1 To 14) As String
    
    lcolOper.add 8, "��λ����"
    lcolOper.add 1, "�������֤"
    lcolOper.add 3, "�ලִ��"
    lcolOper.add 4, "����"
    lcolOper.add 6, "ʳ������Ա"
    lcolOper.add 5, "ҽ�ƻ���"
    lcolOper.add 2, "����֤"
    lcolOper.add 7, "�շ�"
    'lcolOper.add 9, "����"
    'lcolOper.add 10, "����֪ͨ"
    lcolOper.add 9, "�ļ�"
    lcolOper.add 10, "�쵼��ѯ"
    'lcolOper.add 13, "ϵͳ"
    
    For ii = 1 To mobj��.RecordCount
        '�жϸ����Ƿ��в�����
        'cbarOper.Groups.add ii + 1, , mobj��("������")
        Select Case mobj��("������")
            Case "��ѯ"
                'lstrOper(lcolOper(mobj��("������"))) = "���ڹ���"
                clblͨ�ò���(1).Enabled = True
            Case "����"
                'lstrOper(lcolOper(mobj��("������"))) = "���ڹ���"
                clblͨ�ò���(2).Enabled = True
                frm���ڹ���.subLoad mobj��("������"), mobj��, mobj����
            Case "����֪ͨ"
                'lstrOper(lcolOper(mobj��("������"))) = "���ڹ���"
                clblͨ�ò���(3).Enabled = True
                frm����֪ͨ.subLoad mobj��("������"), mobj��, mobj����
            Case "ϵͳ"
                'lstrOper(lcolOper(mobj��("������"))) = "ϵͳ����"
                clblͨ�ò���(4).Enabled = True
                frmϵͳ����.subLoad mobj��("������"), mobj��, mobj����
                'sub��ʼ�������б� frm���ڹ���, mobj��("������")
            Case "�ֵ�"
                'lstrOper(lcolOper(mobj��("������"))) = "���ڹ���"
                clblͨ�ò���(5).Enabled = True
            Case "����"
                lstrOper(lcolOper(mobj��("������"))) = "�������"
                lstrAlias(lcolOper(mobj��("������"))) = mobj��("������")
            Case "����֤"
                lstrOper(lcolOper(mobj��("������"))) = "�� �� ֤"
                lstrAlias(lcolOper(mobj��("������"))) = mobj��("������")
            Case "�շ�"
                lstrOper(lcolOper(mobj��("������"))) = "�շѹ���"
                lstrAlias(lcolOper(mobj��("������"))) = mobj��("������")
            Case "�ļ�"
                lstrOper(lcolOper(mobj��("������"))) = "�ļ�����"
                lstrAlias(lcolOper(mobj��("������"))) = mobj��("������")
            Case Else
                lstrOper(lcolOper(mobj��("������"))) = mobj��("������")
                lstrAlias(lcolOper(mobj��("������"))) = mobj��("������")
        End Select
        mobj��.MoveNext
    Next
    mintMnu = 0
    For ii = 1 To lcolOper.Count
        If lstrOper(ii) <> "" Then
            'cbarOper.Groups.Add ii + 1, , "���� " + lstrOper(ii)
            Load cmnuItem(ii)
            cmnuItem(ii) = lstrOper(ii)
            cmnuItem(ii).Left = cmnuItem(0).Left
            cmnuItem(ii).Top = cmnuItem(ii - 1).Top + cmnuItem(ii - 1).Height + 150
            cmnuItem(ii).Visible = True
            mstrMnu(ii) = lstrAlias(ii)
            mintMnu = mintMnu + 1
            sub��ʼ���ֵ��б� lstrAlias(ii)
        End If
    Next
    
    'cbarOper.Groups.Remove cbarOper.Groups(1)
    
'    frm�ֵ��б�.subLoad

    '���ϵͳ���ṩ��ѯ���� ��ѯͳ�� ��ť������
    Dim lobjRec As Object
'    Set lobjRec = dafuncGetData("select * from ϵͳ����_��ѯ��Ϣ��")
'    If lobjRec.RecordCount = 0 Then
'        clblͨ�ò���(2).Visible = False
'        clblͨ�ò���(5).Left = clblͨ�ò���(4).Left
'        clblͨ�ò���(4).Left = clblͨ�ò���(3).Left
'        clblͨ�ò���(3).Left = clblͨ�ò���(2).Left
''    Else
''        cbarOper.Groups.add
''        cbarOper.Groups(cbarOper.Groups.Count).Caption = "��    ѯ"
''        cbarOper.Groups(cbarOper.Groups.Count).ListItems.add
''        cbarOper.Groups(cbarOper.Groups.Count).ListItems(1).Text = "��ѯ"
''        cbarOper.Groups(cbarOper.Groups.Count).ListItems(1).IconLarge = "��ѯ"
'   End If
    '���ϵͳ���ṩ�����򱨱�ͳ�� ��ť������
    Set lobjRec = dafuncGetData("select * from ���������Ϣ��")
    If lobjRec.RecordCount = 0 Then
'        clblͨ�ò���(3).Visible = False
'        clblͨ�ò���(5).Left = clblͨ�ò���(4).Left
'        clblͨ�ò���(4).Left = clblͨ�ò���(3).Left
    Else
'        '�޸ģ�2001-7-13�������
'        cbarOper.Groups.add
'        cbarOper.Groups(cbarOper.Groups.Count).Caption = "���� ͳ�Ʊ���"
'        cbarOper.Groups(cbarOper.Groups.Count).ListItems.add
'        cbarOper.Groups(cbarOper.Groups.Count).ListItems(1).Text = "ͳ�Ʊ���"
        mintMnu = mintMnu + 1
        Load cmnuItem(mintMnu)
        cmnuItem(mintMnu) = "ͳ�Ʊ���"
        cmnuItem(mintMnu).Left = cmnuItem(0).Left
        cmnuItem(mintMnu).Top = cmnuItem(mintMnu - 1).Top + cmnuItem(mintMnu - 1).Height + 150
        cmnuItem(mintMnu).Visible = True
        mstrMnu(mintMnu) = "ͳ�Ʊ���"
'        sub��ʼ�������ѯ����
    End If
    
   
'    If pcol�ֵ伯.Count Then
'        cbarOper.Groups.add
'        cbarOper.Groups(cbarOper.Groups.Count).Caption = "�ֵ�����"
'        For ii = 1 To pcol�ֵ伯.Count
'            cbarOper.Groups(cbarOper.Groups.Count).ListItems.add ii, "�ֵ�" & pcol�ֵ伯(ii)
'            cbarOper.Groups(cbarOper.Groups.Count).ListItems(ii).Text = pcol�ֵ伯(ii)
'            cbarOper.Groups(cbarOper.Groups.Count).ListItems(ii).IconLarge = pcol�ֵ伯(ii)
'       Next
'    End If
'    cbarOper.Groups.add
'    cbarOper.Groups(cbarOper.Groups.Count).Caption = "�޸Ŀ���"
'    cbarOper.Groups(cbarOper.Groups.Count).ListItems.add
'    cbarOper.Groups(cbarOper.Groups.Count).ListItems(1).Text = "�޸Ŀ���"
'    cbarOper.Groups(cbarOper.Groups.Count).ListItems(1).IconLarge = "�޸Ŀ���"
'    cbarOper.Groups.add
'    cbarOper.Groups(cbarOper.Groups.Count).Caption = "�˳�/ע��"
'    cbarOper.Groups(cbarOper.Groups.Count).ListItems.add
'    cbarOper.Groups(cbarOper.Groups.Count).ListItems(1).Text = "�˳�"
'    cbarOper.Groups(cbarOper.Groups.Count).ListItems(1).IconLarge = "�˳�"
'    cbarOper.Groups(cbarOper.Groups.Count).ListItems.add
'    cbarOper.Groups(cbarOper.Groups.Count).ListItems(2).Text = "ע��"
'    cbarOper.Groups(cbarOper.Groups.Count).ListItems(2).IconLarge = "ע��"
   
'    For ii = 1 To cbarOper.Groups.Count
'        For j = 1 To cbarOper.Groups(ii).ListItems.Count
'            cbarOper.Groups(ii).ListItems(j).ForeColorSource = ssUseListItem
'            cbarOper.Groups(ii).ListItems(j).ForeColor = &H0
'        Next
'    Next
    
    '�޸ģ�2002-1-10(ϵͳ����Ա��һ�����У�ֱ������������״̬���á�����
    On Error Resume Next
    If um�û���� = "0000" Then
        '�ж��Ƿ��һ�����С�
        Dim lobj���� As cls�û���������
        Set lobj���� = New cls�û���������
        lobj����.�û���� = "0000"
        lobj����.ҵ���� = "ϵͳ����"
        If lobj����.������ֵ("��һ������") <> "��" Then
'            ctxtSmartInfos.Caption = ""   '���������Ϣ
            Call sub��������("ϵͳ����_����״̬����")
            '���������й�״̬��
            lobj����.sub���Ǽ���ֵ "��һ������", "��"
        End If
    End If
    
    'Me.Caption = pstrSysName ' & "�����ð棩"
    
    On Error Resume Next
'    Image2.Picture = LoadPicture(App.Path & "\image\background(����" & pstr�汾���� & ").jpg")
    pstrSysName = Me.Caption
    Exit Sub
errHandle:
    If Err.Number = 40003 Or Err.Number = 40002 Then
    Resume Next
    Else
    Call sfsub������("������", "frm������", "Form_Load", Err.Number, Err.Description, False)
    End If
    Exit Sub
    Resume
End Sub


'���ܣ���ʼ�������ѯ���󡣣�����������ʱ���ø÷�������
'�޸ģ�2001-7-13�������
Private Sub sub��ʼ�������ѯ����()

    On Error Resume Next
    
    Set mobjSysAccObj = CreateObject("���ݷ�����.clsSystemBaseAccess")
    If Err <> 0 Then
        '�޸ģ�2002-3-4���������ע�ᡣ
        Err.Clear
        subע�ᱨ���ѯ����
        Err.Clear

        Set mobjSysAccObj = CreateObject("���ݷ�����.clsSystemBaseAccess")
    End If
    If Err <> 0 Then
        On Error GoTo errHandler
        Err.Raise 6666, , "�޷����������ݷ�����.dll���Ķ��󣬡�ͳ�Ʊ������������á���Ҫʹ�ñ������˳�ϵͳ������������������ִ��ϵͳ��װ����"
    End If
    
    '���������ѯ�������
    Set mobjFrontQueryManager = CreateObject("��Դ�����ѯ��.clsFrontQueryManager")
    If Err <> 0 Then
        On Error GoTo errHandler
        Err.Raise 6666, , "�޷���������Դ�����ѯ��.dll���Ķ��󣬡�ͳ�Ʊ������������á���Ҫʹ�ñ������˳�ϵͳ������������������ִ��ϵͳ��װ����"
    End If
    
    '��ʼ�����ݷ��ʶ���
    Dim lstrDatabase As String     '���ݿ���
    lstrDatabase = sffuncGetSetting("ϵͳ����", "���ݿ�����", "���ݿ���")
    mobjSysAccObj.ODBCConnectString = "DSN=WSFY2001;UID=user26;PWD=welcome;DATABASE=" & lstrDatabase
    If Err <> 0 Then
        On Error GoTo errHandler
        Err.Raise 6666, , "�޷���������Դ��WSFY2001�������ݿ⽨�����ӡ���ͳ�Ʊ������������á���Ҫʹ�ñ������˳�ϵͳ������������������ִ��ϵͳ��װ����"
    End If
    
    '�ж���ʱ·���Ƿ���ڡ��������ڣ�����֮��
    If Dir("c:\temp", vbDirectory) = "" Then
        MkDir "c:\temp"
    End If
    Err.Clear
    
    '��ʼ�������ѯ����
    mobjFrontQueryManager.subFrontQueryInitalize mobjSysAccObj, "", "c:\temp\" ', lobjSpec
    If Err <> 0 Then
        On Error GoTo errHandler
        Err.Raise 6666, , "��ʼ�������ѯ��ʧ�ܡ���ͳ�Ʊ������������á���Ҫʹ�ñ������˳�ϵͳ������������������ִ��ϵͳ��װ����"
    End If
    
    mobjFrontQueryManager.��ǰ�û� = um�û����
    Exit Sub
errHandler:
    Set mobjFrontQueryManager = Nothing
    Set mobjSysAccObj = Nothing
    Call sfsub������("������", "frm������", "sub��ʼ�������б�", Err.Number, Err.Description, True)
End Sub

'���ܣ����������ѯ���档���û�����ͳ�Ʊ��������˵�ʱ���ã�
'�޸ģ�2001-7-13�������
Private Sub sub���������ѯ()
    Dim lobjRec As Object
    Dim lcolReports As New Collection
    Dim lcolItem As Collection
    Dim lobjValueList As Object
    Dim lstrSql As String
    Dim i As Long
    
    On Error GoTo errHandler
    If mobjFrontQueryManager Is Nothing Then
        '�ٴγ�ʼ����
        sub��ʼ�������ѯ����
        'Err.Raise 6666, , "�����ѯ��ʼ��ʧ�ܣ��������������ѯ���档���˳�ϵͳ�����°�װϵͳ��"
    Else
    '�޸ģ�2001-8-15�������
        If mobjFrontQueryManager.ReportDataObject Is Nothing Then
            sub��ʼ�������ѯ����
        End If
    End If
    
    '������ѯ���档
    Dim llng������ As Long
    Dim llngWndProc As Long
    '�޸ģ�2003-7-22�����������ϵͳ��ɲ�����
    llng������ = mobjFrontQueryManager.funcStart(pstr��ϵͳ���)
    
    '�趨�򿪵Ĵ���Ϊ��������Ӵ��塣
    If llng������ <> -2 Then
        '�򼯺��м����������
        If Not sffunc�жϼ��ϼ�ֵ�Ƿ����(pcol��������, CStr(llng������)) Then
            On Error Resume Next
            SetParent llng������, Me.hWnd
            llngWndProc = SetWindowLong(llng������, GWL_WNDPROC, AddressOf funcClassing)
            pcolWndProc.add llngWndProc, CStr(llng������)
            pcol��������.add "����ͳ��", CStr(llng������)
            pcol�Ӵ�����.add llng������, "����ͳ��"
            
            Call MoveWindow(llng������, ScaleX(1750, vbTwips, vbPixels), ScaleX(cimgBackground.Height - 60, vbTwips, vbPixels), ScaleX(Me.ScaleWidth - 1700, vbTwips, vbPixels), ScaleX(Me.ScaleHeight - cstatusBar.Height - cimgBackground.Height + 100, vbTwips, vbPixels), 1)
            'Call MoveWindow(llng������, ScaleX(1700, vbTwips, vbPixels), ScaleX(60, vbTwips, vbPixels), ScaleX(Me.ScaleWidth - 1700, vbTwips, vbPixels), ScaleX(Me.ScaleHeight - 380, vbTwips, vbPixels), 1)

            Err.Clear
            On Error GoTo errHandler
        End If
    End If
    
    Exit Sub
errHandler:
    Call sfsub������("������", "frm������", "sub���������ѯ", Err.Number, Err.Description, True)
    Exit Sub
    Resume
End Sub

'���ܣ������û��Ĵ���
'���룺��������
'�������
'���أ���
'ע���������ô����Ѵ�������ʾ��δ��������ø���ϵͳ��ͨ�÷���funcStart����ҵ����
'���ߣ�������
'����ʱ�䣺2001-3-8
' �޸�˵�������ڴ�������ڲ�ͬ���̼䴫�ݴ������⣬�ʸ�Ϊֻ���ش�������
' �޸��ˣ�  ����
' �޸�ʱ�䣺2001-3-20
Public Sub sub��������(ByVal paraҵ������ As String)
    On Error GoTo errHandle
    Dim lobj���� As Object '�����Ӵ���Ľ������
    Dim lobjȨ�� As Object  '��ǰ�û�����Ȩ��
    Set lobjȨ�� = um����Ȩ��
    Dim llng������ As Long          '��ǰ��Ӵ���
    Dim llngWndProc As Long
    Dim lstrҵ���� As String
    If mblnLoadForm Then Exit Sub
    mblnLoadForm = True
   
    lobjȨ��.Filter = "Ȩ����" & "= '" & paraҵ������ & "'"  '�Ƚ��û���Ȩ���Ƿ��ܲ����ò���
    mobj����.Filter = 0
    mobj����.Filter = "��������" & "= '" & paraҵ������ & "'"
    If lobjȨ��.RecordCount > 0 Then '��Ȩ���и������
        '����ҵ�����
        um��ǰ������ϵͳ�� = mobj����("ҵ����")
        Set lobj���� = CreateObject(mobj����("������") & "." & mobj����("����"))
        
        If paraҵ������ = "ϵͳ����_�����ѯȨ������" Then
            llng������ = lobj����.funcStart(paraҵ������, pstr��ϵͳ���)
        Else
            llng������ = lobj����.funcStart(paraҵ������)
        End If
        
        If llng������ = -1 Then Err.Raise 6666, , "���������趨����δ�ҵ��ò�����������Ӧ�Ĵ��壡"
        '�趨�򿪵Ĵ���Ϊ��������Ӵ��塣
        If llng������ <> -2 Then
            '�򼯺��м����������
            If Not sffunc�жϼ��ϼ�ֵ�Ƿ����(pcol��������, CStr(llng������)) Then
                On Error Resume Next
                lstrҵ���� = mobj����("ҵ����")
                SetParent llng������, Me.hWnd
                llngWndProc = SetWindowLong(llng������, GWL_WNDPROC, AddressOf funcClassing)
                pcolWndProc.add llngWndProc, CStr(llng������)
                pcolҵ������.add lstrҵ����, CStr(llng������)
                pcol�Ӵ�����.add llng������, paraҵ������
                pcol��������.add paraҵ������, CStr(llng������)
                
                Call MoveWindow(llng������, ScaleX(1750, vbTwips, vbPixels), ScaleX(cimgBackground.Height - 60, vbTwips, vbPixels), ScaleX(Me.ScaleWidth - 1700, vbTwips, vbPixels), ScaleX(Me.ScaleHeight - cstatusBar.Height - cimgBackground.Height + 100, vbTwips, vbPixels), 1)

                Err.Clear
                On Error GoTo errHandle
                Call oesubSave("�û�����" & paraҵ������, "�������")
            End If
        End If
        Err.Clear
        On Error GoTo errHandle
    Else    '��Ȩ���и������
        Call sffuncMsg("��Ȩ�޽��и������", sf����)
    End If
errHandle:
    mblnLoadForm = False
    Set lobj���� = Nothing
    Set lobjȨ�� = Nothing
    If Err.Number = 0 Then Exit Sub
    If Err.Number = 429 Then
        Err.Number = 6666
        Err.Description = "�ò���δ�ڱ�����ȷ��װ��ע�ᣡ"
    End If
    Call sfsub������("������", "frm������", "sub��������", Err.Number, Err.Description, False)
End Sub


'���ܣ���������Ϣд��������
'���룺 ��
'����� ��
'���أ� ��
'ע�������
'���ߣ�������
'����ʱ�䣺2001-03-21
Private Sub WriteErrLog()
    On Error GoTo errHandler
    Dim lstr�û���� As String        '�û����
    Dim lstr����վ��� As String      '����վ���
    Dim ldat����  As Date            '�����������
    Dim lstr�����  As String
    Dim lstr�������� As String
    Dim lstr�������·�� As String     '�������·��
    Dim lstrSql As String
    Dim lstrInput As String
    If Dir("C:\ErrLog") = "" Then Exit Sub  '�����¼Ϊ�����˳�
    lstr�û���� = um�û����              'ȡ�û����
    lstr����վ��� = um����վ���         'ȡ����վ���
    Open "C:\ErrLog" For Input As #1     '�򿪴����¼��
        Do While Not EOF(1)
            Line Input #1, lstrInput
            lstr����� = Mid(lstrInput, InStr(1, lstrInput, "����ţ�") + 4, InStr(1, lstrInput, "����������") - InStr(1, lstrInput, "����ţ�") - 5)
            lstr�������� = Mid(lstrInput, InStr(1, lstrInput, "����������") + 5, InStr(1, lstrInput, "�������·����") - InStr(1, lstrInput, "����������") - 6)
            Do While InStr(1, lstr��������, "'")
                lstr�������� = Left(lstr��������, InStr(1, lstr��������, "'") - 1) & "`" & Right(lstr��������, Len(lstr��������) - InStr(1, lstr��������, "'"))
            Loop
            lstr�������� = LeftB(lstr��������, 500)
            If lstr�������� <> "" Then
                lstr�������� = Replace(lstr��������, "'", "''")     '�����������г��ֵ�"'"ת����"''"
            End If
            lstr�������·�� = Mid(lstrInput, InStr(1, lstrInput, "�������·����") + 7, InStr(1, lstrInput, "���ڣ�") - InStr(1, lstrInput, "�������·����") - 8)
            ldat���� = Format(Mid(lstrInput, InStr(1, lstrInput, "���ڣ�") + 3, Len(lstrInput) - InStr(1, lstrInput, "���ڣ�")), "yyyy/mm/dd hh:mm:ss")
            'д�����ݿ�
            lstrSql = "Insert Into ϵͳ����_ϵͳ�����¼�� Values('" & _
            lstr�û���� & "' ,'" & _
            lstr����վ��� & "','" & _
            ldat���� & "','" & _
            lstr����� & "','" & _
            lstr�������� & "','" & _
            lstr�������·�� & "')"
            dafuncGetData (lstrSql)
        Loop
    Close #1
    Kill "C:\ErrLog"
    Exit Sub
errHandler:
    If Err.Number <> 3000 Then
        Resume Next
    Else
        Close #1
        Kill "C:\ErrLog"
    End If
End Sub


Private Sub subע�ᱨ���ѯ����()
    Dim lstrPath As String
    Dim lstrFile As String
    Dim llngRes As Long
    Dim lstrLongPath As String
    Dim lstrShortPath As String
    
    On Error Resume Next
    
    '�ѳ�·��ת��Ϊ��·����
    lstrLongPath = App.Path & "\�������\"
    lstrPath = String$(165, 0)
    
    lstrFile = lstrLongPath & "���ݷ�����.dll"
    llngRes = GetShortPathName(lstrFile, lstrPath, 164)
    If llngRes > 0 Then
        lstrFile = Left$(lstrPath, llngRes)
    End If
    Shell "Regsvr32 /s " & lstrFile, vbNormalFocus
    
    lstrFile = lstrShortPath & "FileToDatabase.dll"
    llngRes = GetShortPathName(lstrFile, lstrPath, 164)
    If llngRes > 0 Then
        lstrFile = Left$(lstrPath, llngRes)
    End If
    Shell "Regsvr32 /s " & lstrFile, vbNormalFocus
    
    lstrFile = lstrShortPath & "�����ѯ����.dll"
    llngRes = GetShortPathName(lstrFile, lstrPath, 164)
    If llngRes > 0 Then
        lstrFile = Left$(lstrPath, llngRes)
    End If
    Shell "Regsvr32 /s " & lstrFile, vbNormalFocus
    
    lstrFile = lstrShortPath & "��Դ�����ѯ��.dll"
    llngRes = GetShortPathName(lstrFile, lstrPath, 164)
    If llngRes > 0 Then
        lstrFile = Left$(lstrPath, llngRes)
    End If
    Shell "Regsvr32 /s " & lstrFile, vbNormalFocus
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If (X > 3500 Or X < 1800) And Y > 900 Then
'        cfrm�ֵ�.Visible = False
'    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mblnLoadForm Then Cancel = True: Exit Sub
    On Error Resume Next
    Dim llng������ As Long
    Dim lstrTemp As String
'    Me.WindowState = 0
    frmExit.Show vbModal, Me
    If pblnCancel Then
        Cancel = True
        Exit Sub
    End If
    Dim i As Integer
    Dim lstr������ As String
    Dim lobj���� As Object
    Dim lint�������� As Integer
    lint�������� = pcol��������.Count
    For i = 1 To lint��������
        lstr������ = pcol��������(1)
        If lstr������ <> "ƽ̨����" Then
            If lstr������ = "�ֵ����" Then
                Set lobj���� = CreateObject("�ֵ����.clsalldictionarys")
        
            ElseIf lstr������ = "����ͳ��" Then
                '�޸ģ�2001-7-13�����
                Set lobj���� = mobjFrontQueryManager
            ElseIf lstr������ = "��ѯͳ��" Then
                '�޸ģ�2001-7-24�����죩
                Set lobj���� = CreateObject("ͨ�ò�ѯ.clsͨ�ò�ѯ")
            Else
                mobj����.Filter = "��������" & "= '" & lstr������ & "'"
                '����ҵ�����
                Set lobj���� = CreateObject(mobj����("������") & "." & mobj����("����"))
            End If
            On Error GoTo Goon
            lobj����.funcClose lstr������
            If sffunc�жϼ��ϼ�ֵ�Ƿ����(pcol�Ӵ�����, lstr������) Then
                Cancel = True
                Exit For
            End If
        
        Else
            Unload frmƽ̨����
            If sffunc�жϼ��ϼ�ֵ�Ƿ����(pcol�Ӵ�����, lstr������) Then
                Cancel = True
                Exit For
            End If
        End If
Goon:
    If Err.Number <> 0 Then Err.Clear
    On Error Resume Next
    Next i
    If Cancel = True Then
        mblnRe = False
        Exit Sub
    Else
        mblnRe = True
        Set pobjƽ̨�ṹ = Nothing
    End If
    WriteErrLog                                '��������Ϣд�����ݿ�
    Call oesubSave("�û��˳�ϵͳ", "�˳�ϵͳ") '��¼������־
    SetWindowLong Me.hWnd, GWL_WNDPROC, pcolWndProc(CStr(Me.hWnd))
    If Cancel <> True Then
        Me.Hide
        Set pcolWndProc = Nothing
        Set pcol�������� = Nothing
        Set pcolҵ������ = Nothing
        Set pcol�Ӵ����� = Nothing
        Set mobjFrontQueryManager = Nothing
        If Not pblnExit And mblnRe Then
            pblnע�� = True
            Call oesubSave("�û�ע�����½���ϵͳ", "ע��")
            Unload frm����֪ͨ
            Unload frm���ڹ���
            Unload frmϵͳ����
            Unload frm�ֵ��б�
            Call Main
        Else
            '�޸ģ�2002-2-26���˳�����
            subExit
        End If
    End If
End Sub
Private Sub subExit()
    On Error Resume Next
    X.subCloseDatabase
    Unload frm����֪ͨ
    Unload frm���ڹ���
    Unload frmϵͳ����
    Unload frm�ֵ��б�
End Sub
Private Sub Form_Resize()
    On Error Resume Next
'    Image1.Left = 0
'    Image1.Top = 0
'    Image1.Width = Me.ScaleWidth
'    Image1.Height = Me.ScaleHeight
''    clblClose.Left = Me.ScaleWidth - 375
'    Image1.ZOrder 1
'    clblInfo.Top = Me.ScaleHeight - 500
    Frame1.Height = Me.ScaleHeight - cstatusBar.Height - cimgBackground.Height + 100
    cimgBackground.Width = Me.ScaleWidth - cimgBackground.Left
    subResizeChild
End Sub
Private Sub sub��ʼ���ֵ��б�(para���� As String)
    Dim i As Integer, j As Integer
    Dim lobjRec As Object
    
    On Error GoTo errHandle
    
    mobj��.Filter = ""
    mobj��.Filter = "��������" & " ='" & para���� & "' "
    For i = 1 To mobj��.RecordCount
        '�޸ģ�2003-7-9������жϵ�ǰ��������ҵ�����Ƿ��ڼ��ܹ���ɷ�Χ�ڡ�
        mobj����.Filter = ""
        mobj����.Filter = "��������" & "='" & mobj��.Fields("��������") & "'"
        If mobj����.RecordCount > 0 Then
            If pstr��ϵͳ��� = "" Or InStr(pstr��ϵͳ���, mobj����.Fields("ҵ����") & ",") > 0 Then
                If Not sffunc�жϼ��ϼ�ֵ�Ƿ����(pcol�ֵ伯, mobj����.Fields("ҵ����").Value) Then
                    '�жϸ�ҵ���Ƿ��в��������ֵ䡣
                    Set lobjRec = dafuncGetData("select * from ϵͳ����_�ֵ�_�ֵ���б� where ҵ����='" & mobj����.Fields("ҵ����").Value & "' and ����='������'")
                    If lobjRec.RecordCount > 0 Then
                        pcol�ֵ伯.add mobj����.Fields("ҵ����").Value, mobj����.Fields("ҵ����").Value
                    End If
                End If
            End If
        End If
        mobj��.MoveNext
    Next i

    Exit Sub
errHandle:
    If Err.Number = 40002 Or Err.Number = 40003 Or Err.Number = 40006 Then
        Resume Next
    Else
        Call sfsub������("������", "frm������", "sub��ʼ�������б�", Err.Number, Err.Description, False)
    End If
End Sub


Private Sub sub��ʼ�������б�(para���� As String)
    Dim i As Integer, j As Integer
    Dim ii As Integer
    Dim lobjRec As Object
    On Error GoTo errHandle
    
    'frm�����б�.subClear

    '����ò�����Ĳ���
    'frm�����б�.clblTitle.Caption = para����
    ii = 1
    mobj��.Filter = ""
    mobj��.Filter = "��������" & " ='" & para���� & "' "
    For i = 1 To 10
        cmnuSubItem(i).Visible = False
        cmnuSubItem(i) = ""
    Next
    'cbarOper.Groups(cbarOper.Groups.Count).ListItems.Clear
    If para���� = "ͳ�Ʊ���" Then
        cmnuSubItem(ii) = "ͳ�Ʊ���"
        cmnuSubItem(ii).Top = cmnuSubItem(ii - 1).Top + cmnuSubItem(ii - 1).Height + 150
        cmnuSubItem(ii).Left = cmnuSubItem(0).Left
        cmnuSubItem(ii).Visible = True
        mstrOper(ii) = "ͳ�Ʊ���"
        Exit Sub
    End If
    For i = 1 To mobj��.RecordCount
        '�޸ģ�2003-7-9������жϵ�ǰ��������ҵ�����Ƿ��ڼ��ܹ���ɷ�Χ�ڡ�
        mobj����.Filter = ""
        mobj����.Filter = "��������" & "='" & mobj��.Fields("��������") & "'"
        If mobj����.RecordCount > 0 Then
            If pstr��ϵͳ��� = "" Or InStr(pstr��ϵͳ���, mobj����.Fields("ҵ����") & ",") > 0 Then
                cmnuSubItem(ii) = mobj��("��������")
                cmnuSubItem(ii).Top = cmnuSubItem(ii - 1).Top + cmnuSubItem(ii - 1).Height + 150
                cmnuSubItem(ii).Left = cmnuSubItem(0).Left
                cmnuSubItem(ii).Visible = True
                mstrOper(ii) = mobj��.Fields("��������")
                ii = ii + 1
            
            End If
        End If
        mobj��.MoveNext
    Next i

    Exit Sub
errHandle:
    If Err.Number = 40002 Or Err.Number = 40003 Or Err.Number = 40006 Then
        Resume Next
    Else
        Call sfsub������("������", "frm������", "sub��ʼ�������б�", Err.Number, Err.Description, False)
    End If
End Sub

Public Sub sub�����ֵ�(ByVal para��ϵͳ�� As String)
    On Error Resume Next
'    ctxtSmartInfos.Caption = ""   '���������Ϣ
    On Error GoTo errHandle
    Dim lobj���� As Object '�����Ӵ���Ľ������
    Dim llng������ As Long          '��ǰ��Ӵ���
    Dim llngWndProc As Long
    '����ҵ�����
    
    um��ǰ������ϵͳ�� = para��ϵͳ�� 'clbl�ֵ�(Index).Caption
    Set lobj���� = CreateObject("�ֵ����.clsalldictionarys")
    llng������ = lobj����.funcStart(para��ϵͳ��)
    If llng������ = -1 Then Err.Raise 6666, , "���������趨����δ�ҵ��ò�����������Ӧ�Ĵ��壡"
    '�趨�򿪵Ĵ���Ϊ��������Ӵ��塣
    If llng������ <> -2 Then
        '�򼯺��м����������
        If Not sffunc�жϼ��ϼ�ֵ�Ƿ����(pcol��������, CStr(llng������)) Then
            On Error Resume Next
            SetParent llng������, Me.hWnd
            llngWndProc = SetWindowLong(llng������, GWL_WNDPROC, AddressOf funcClassing)
            pcolWndProc.add llngWndProc, CStr(llng������)
            pcol��������.add "�ֵ����", CStr(llng������)
            pcol�Ӵ�����.add llng������, "�ֵ����"
            Call MoveWindow(llng������, ScaleX(1750, vbTwips, vbPixels), ScaleX(cimgBackground.Height - 60, vbTwips, vbPixels), ScaleX(Me.ScaleWidth - 1700, vbTwips, vbPixels), ScaleX(Me.ScaleHeight - cstatusBar.Height - cimgBackground.Height + 100, vbTwips, vbPixels), 1)
            
            '********************��ʱ���Դ���********************
            'sfsubSaveSetting "ϵͳ����", "ƽ̨�������", "�ֵ�����ʼ��", "hWnd��" & llng������ & " WndProc��" & llngWndProc & " ʱ��" & Format(Now, "yyyy��mm��dd��hhʱmm��ss��")
            '****************************************************
            Err.Clear
            On Error GoTo errHandle
            Call oesubSave("�û������ֵ����", "�������")
        End If
    End If
errHandle:
    Set lobj���� = Nothing
    If Err.Number = 0 Then Exit Sub
    If Err.Number = 429 Then
        Err.Number = 6666
        Err.Description = "�ò���δ�ڱ�����ȷ��װ��ע�ᣡ"
    End If
    Call sfsub������("������", "frm������", "sub�����ֵ�", Err.Number, Err.Description, False)
End Sub


Private Sub sub�����������()
    Dim lstrTime As String
    
    lstrTime = "2005-12-31"
    
    If lstrTime < Format(Now, "yyyy-mm-dd") Then
        MsgBox "�Բ���������������ѵ������������Ӧ����ϵ��", vbCritical, "ϵͳ��ʾ"
        End
    End If
End Sub


Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    
    For i = 1 To mintMnu
        cmnuItem(i).FontUnderline = False
        cmnuItem(i).ForeColor = vbBlack
    Next
    For i = 1 To 10
        cmnuSubItem(i).FontUnderline = False
        cmnuSubItem(i).ForeColor = vbBlack
    Next
End Sub

Private Sub Timer1_Timer()
    sub�����������
End Sub


Public Sub subResizeChild()
    Dim llngHwnd As Long
    Dim i As Long
    
    For i = 1 To pcol�Ӵ�����.Count
        llngHwnd = pcol�Ӵ�����(i)

        Call MoveWindow(llngHwnd, ScaleX(1750, vbTwips, vbPixels), ScaleX(cimgBackground.Height - 60, vbTwips, vbPixels), ScaleX(Me.ScaleWidth - 1700, vbTwips, vbPixels), ScaleX(Me.ScaleHeight - cstatusBar.Height - cimgBackground.Height + 100, vbTwips, vbPixels), 1)
        'Call MoveWindow(llngHwnd, ScaleX(700, vbTwips, vbPixels), ScaleX(350, vbTwips, vbPixels), ScaleX(Me.ScaleWidth - 700, vbTwips, vbPixels), ScaleX(Me.ScaleHeight - 400, vbTwips, vbPixels), 1)

    Next
End Sub

