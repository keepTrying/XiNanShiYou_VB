VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "�������ƹ�����Ϣϵͳ"
   ClientHeight    =   10995
   ClientLeft      =   180
   ClientTop       =   465
   ClientWidth     =   13710
   ClipControls    =   0   'False
   ForeColor       =   &H00B9F7D3&
   Icon            =   "frmMain1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10995
   ScaleWidth      =   13710
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2640
      Top             =   3360
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   2280
      Top             =   5040
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00F7F2E1&
      Caption         =   "���칤��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   9000
      TabIndex        =   24
      Top             =   7800
      Visible         =   0   'False
      Width           =   5055
      Begin VB.TextBox clblMessage 
         Appearance      =   0  'Flat
         BackColor       =   &H00F7F2E1&
         BorderStyle     =   0  'None
         Height          =   2655
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         Top             =   240
         Width           =   4815
      End
   End
   Begin MSComctlLib.StatusBar cstatusBar 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   10635
      Width           =   13710
      _ExtentX        =   24183
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4817
            MinWidth        =   1060
            Text            =   "�û���ţ�"
            TextSave        =   "�û���ţ�"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4817
            MinWidth        =   1060
            Text            =   "�û�������"
            TextSave        =   "�û�������"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4815
            MinWidth        =   1058
            Text            =   "����վ����"
            TextSave        =   "����վ����"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9049
            MinWidth        =   5292
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   1800
      Top             =   1200
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   7635
      Left            =   -45
      TabIndex        =   7
      Top             =   500
      Width           =   1635
      Begin VB.Image cimgPhoto 
         Height          =   1335
         Left            =   240
         Stretch         =   -1  'True
         Top             =   6240
         Width           =   1095
      End
      Begin VB.Image cimgButton 
         Height          =   300
         Index           =   7
         Left            =   0
         Picture         =   "frmMain1.frx":0CCA
         Top             =   0
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Image cimgButton 
         Height          =   300
         Index           =   6
         Left            =   0
         Picture         =   "frmMain1.frx":111C
         Top             =   0
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Image cimgButton 
         Height          =   300
         Index           =   5
         Left            =   0
         Picture         =   "frmMain1.frx":156E
         Top             =   0
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Image cimgButton 
         Height          =   300
         Index           =   4
         Left            =   0
         Picture         =   "frmMain1.frx":19C0
         Top             =   0
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Image cimgButton 
         Height          =   300
         Index           =   3
         Left            =   0
         Picture         =   "frmMain1.frx":1E12
         Top             =   0
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Image cimgButton 
         Height          =   300
         Index           =   2
         Left            =   0
         Picture         =   "frmMain1.frx":2264
         Top             =   0
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Image cimgButton 
         Height          =   300
         Index           =   1
         Left            =   0
         Picture         =   "frmMain1.frx":26B6
         Top             =   0
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Image Image5 
         Height          =   765
         Left            =   15
         Picture         =   "frmMain1.frx":2B08
         Stretch         =   -1  'True
         Top             =   6705
         Width           =   1725
      End
      Begin VB.Image cimgButton 
         Height          =   330
         Index           =   10
         Left            =   150
         Picture         =   "frmMain1.frx":5B12
         Stretch         =   -1  'True
         Top             =   285
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Image cimgButton 
         Height          =   300
         Index           =   8
         Left            =   300
         Picture         =   "frmMain1.frx":5F64
         Top             =   120
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Image cimgButton 
         Height          =   300
         Index           =   0
         Left            =   105
         Picture         =   "frmMain1.frx":63B6
         Top             =   570
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Label cmnuItem 
         BackColor       =   &H009EE9C4&
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   22
         Top             =   -240
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "�� �� �� ��"
         Height          =   165
         Left            =   615
         TabIndex        =   19
         Top             =   -100
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Image Image3 
         Height          =   405
         Left            =   255
         Picture         =   "frmMain1.frx":6808
         Stretch         =   -1  'True
         Top             =   1620
         Visible         =   0   'False
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         Index           =   4
         Left            =   570
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   210
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
         Left            =   465
         TabIndex        =   8
         Top             =   615
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image cpicButton 
         Height          =   360
         Index           =   0
         Left            =   60
         Picture         =   "frmMain1.frx":6EA3
         Top             =   -285
         Visible         =   0   'False
         Width           =   1530
      End
   End
   Begin VB.Label clblͨ�ò��� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�Զ�����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   7
      Left            =   12840
      TabIndex        =   28
      Top             =   120
      Width           =   840
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6960
      TabIndex        =   27
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5280
      TabIndex        =   26
      Top             =   7800
      Width           =   4575
   End
   Begin VB.Label clblͨ�ò��� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ͳ�Ʊ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   2
      Left            =   6720
      TabIndex        =   23
      Top             =   120
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Image cimgButton 
      Height          =   300
      Index           =   9
      Left            =   45
      Picture         =   "frmMain1.frx":9918
      Top             =   645
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Image4 
      Height          =   1575
      Left            =   1600
      Picture         =   "frmMain1.frx":9D6A
      Top             =   500
      Width           =   2535
   End
   Begin VB.Label clblSysName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����ල��������Ϣϵͳ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   3300
   End
   Begin VB.Label clblͨ�ò��� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   0
      Left            =   11880
      TabIndex        =   20
      Top             =   120
      Width           =   420
   End
   Begin VB.Label clblͨ�ò��� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ϵͳ����"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   3
      Left            =   8400
      TabIndex        =   6
      Top             =   120
      Width           =   840
   End
   Begin VB.Label clblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�û���ţ�           �û����ƣ�           ����վ����"
      Height          =   180
      Left            =   2700
      TabIndex        =   4
      Top             =   7530
      Visible         =   0   'False
      Width           =   4680
   End
   Begin VB.Label clblͨ�ò��� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�˳�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   6
      Left            =   11280
      TabIndex        =   3
      Top             =   120
      Width           =   420
   End
   Begin VB.Label clblͨ�ò��� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����޸�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   5
      Left            =   10320
      TabIndex        =   2
      Top             =   120
      Width           =   840
   End
   Begin VB.Label clblͨ�ò��� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ѯ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   1
      Left            =   7800
      TabIndex        =   1
      Top             =   120
      Width           =   420
   End
   Begin VB.Label clblͨ�ò��� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ֵ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   4
      Left            =   9360
      TabIndex        =   0
      Top             =   120
      Width           =   840
   End
   Begin VB.Image Image1 
      Height          =   7695
      Left            =   3525
      Picture         =   "frmMain1.frx":E268
      Stretch         =   -1  'True
      Top             =   1590
      Width           =   7875
   End
   Begin VB.Image cimgBackground 
      Height          =   585
      Left            =   0
      Picture         =   "frmMain1.frx":1F0A2
      Stretch         =   -1  'True
      Top             =   -90
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

'�޸ģ���������
Private X As Object

Private mblnAutoUpgrade As Boolean      '��ǰҪ�����Զ�����

'�����ѯ��Ҫ�ı�����
Private mobjFrontQueryManager As Object
Private mobjSysAccObj As Object         '����������.clsSystemAccessObject��

Private mintMinutes As Integer
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cimg��_Click(Index As Integer)

End Sub


Private Sub cimgBackground_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    
    For i = 0 To clblͨ�ò���.Count - 1
        clblͨ�ò���(i).FontUnderline = False
        clblͨ�ò���(i).ForeColor = vbWhite
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
    clblͨ�ò���(Index).ForeColor = vbWhite
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
         
         Case "ϵͳ����"
            frmϵͳ����.Hide

            mobj��.Filter = ""
            
            If um�û���� = "0000" Then
                mobj��.Filter = "������='ϵͳ' or ������='ϵͳ����'"

            Else
                mobj��.Filter = "��������= 'ҵ����' and ������='ϵͳ'"
            End If
            
            If mobj��.RecordCount > 0 Then
                frmϵͳ����.subLoad mobj��("������"), mobj��, mobj����
                    
                Set frmϵͳ����.pobjParent = Me
                frmϵͳ����.Move clblͨ�ò���(Index).Left, Me.Top + 800
                frmϵͳ����.Show , Me
            End If
            mobj��.Filter = "��������= 'ҵ����'"
            
        Case "��ѯ"      'ѡȡ��ƽ̨���в�ѯ
                '������ѯ���档
                Dim lobjͨ�ò�ѯ As Object
                Set lobjͨ�ò�ѯ = CreateObject("ͨ�ò�ѯ.clsͨ�ò�ѯ")

                llng������ = lobjͨ�ò�ѯ.funcStart("ϵͳ����_ͨ�ò�ѯ", pstr��ϵͳ���)
                
                '�趨�򿪵Ĵ���Ϊ��������Ӵ��塣
                If llng������ <> -2 Then
                    '�򼯺��м����������
                    If Not sffunc�жϼ��ϼ�ֵ�Ƿ����(pcol��������, CStr(llng������)) Then
                        On Error Resume Next
                        SetParent llng������, Me.hWnd
                        llngWndProc = SetWindowLong(llng������, GWL_WNDPROC, AddressOf funcClassing)
                        pcolWndProc.Add llngWndProc, CStr(llng������)
                        pcol��������.Add "��ѯͳ��", CStr(llng������)
                        pcol�Ӵ�����.Add llng������, "��ѯͳ��"

                        Call MoveWindow(llng������, ScaleX(1600, vbTwips, vbPixels), ScaleX(cimgBackground.Height - 60, vbTwips, vbPixels), ScaleX(Me.ScaleWidth - 1650, vbTwips, vbPixels), ScaleX(Me.ScaleHeight - cstatusBar.Height - cimgBackground.Height + 100, vbTwips, vbPixels), 1)
                        Err.Clear
                    End If
                End If
        Case "ͳ�Ʊ���"       'ѡȡ��ƽ̨���б���
            '���������ѯ���档
            Call sub���������ѯ
                
        Case "�����޸�"
            frm�����޸�.Show vbModal, Me
        Case "�˳�"
            Unload Me
        Case "����"
            'MsgBox "���������У�", vbOKOnly, "����"
            ShellExecute Me.hWnd, "Open", App.Path + "\�û��ֲ�\manual.chm", "", "", 1
        Case "�Զ�����"
            If MsgBox("�����Զ�����ʱ���������˳�ϵͳ��ȷ�������Ѿ�����ã������˳�����", vbQuestion + vbYesNo, "ϵͳѯ��") = vbNo Then Exit Sub
            Shell App.Path & "\autoupgrade.exe '" & pstr�û���� & "'", vbNormalFocus
            mblnAutoUpgrade = True
            Unload Me
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
    
    If cmnuSubItem(1).Visible And cmnuSubItem(1).Top = cpicButton(Index).Top + 450 Then '���˵�������
        For ii = 1 To 10
            cmnuSubItem(ii).Visible = False
            cimgButton(ii).Visible = False
        Next
        For ii = 1 To mintMnu
            cpicButton(ii).Top = cpicButton(ii - 1).Top + cpicButton(ii - 1).Height + 30
            cmnuItem(ii).Top = cpicButton(ii).Top + 90
        Next
        Exit Sub
    End If
    For ii = 1 To Index
        cpicButton(ii).Top = cpicButton(ii - 1).Top + cpicButton(ii - 1).Height + 30
        cmnuItem(ii).Top = cpicButton(ii).Top + 90
    Next

    cmnuSubItem(0).Top = cpicButton(ii - 1).Top + 120
    sub��ʼ�������б� mstrMnu(Index)
    For j = 1 To 10
        If cmnuSubItem(j) = "" Then Exit For
    Next
    If ii <= mintMnu Then
        cpicButton(ii).Top = cmnuSubItem(j - 1).Top + cmnuSubItem(j - 1).Height + 100
        cmnuItem(ii).Top = cpicButton(ii).Top + 90
        For ii = ii + 1 To mintMnu
            cpicButton(ii).Top = cpicButton(ii - 1).Top + cpicButton(ii - 1).Height + 30
            cmnuItem(ii).Top = cpicButton(ii).Top + 90
        Next
    End If
    cmnuItem(Index).FontUnderline = False
    cmnuItem(Index).ForeColor = vbBlack
End Sub

Private Sub cmnuItem_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    
    For i = 0 To cmnuItem.Count - 1
        If i = Index Then
            cmnuItem(Index).FontUnderline = True
            cmnuItem(Index).ForeColor = &H80FF&
        Else
            cmnuItem(i).FontUnderline = False
            cmnuItem(i).ForeColor = vbBlack
        End If
    Next
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
    Dim i As Integer
    
    For i = 0 To cmnuSubItem.Count - 1
        If cmnuSubItem(i).Visible Then
            If i = Index Then
                cmnuSubItem(Index).FontUnderline = True
                cmnuSubItem(Index).ForeColor = &H80FF&
            Else
                cmnuSubItem(i).FontUnderline = False
                cmnuSubItem(i).ForeColor = vbBlack
            End If
        End If
    Next
End Sub

Private Sub Form_Activate()
    sub��ʾ���칤��
End Sub

Public Sub sub��ʾ���칤��()
    Dim lobjRec As Object
    On Error Resume Next
    
    '��ȡ���칤����
    Set lobjRec = dafuncGetData("exec ϵͳ����_��ȡ���칤�� '" & um�û���� & "'")
    clblMessage.Text = ""
    If lobjRec.RecordCount > 0 Then
        clblMessage.Text = IIf(IsNull(lobjRec(0)), "", lobjRec(0))
    End If
    If clblMessage.Text = "" Then
        Frame2.Visible = False
    Else
        Frame2.Visible = True
    End If
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
    
    On Error Resume Next
    Dim lstrServer As String
    lstrServer = sffuncGetSetting("ϵͳ����", "���ݿ�����", "��������")
    
    clblSysName = um����վ�� & "������Ϣϵͳ"
    
    '��ȡ�������ơ�
    Dim lstrLocalName As String
    Dim lobjRec As Object
    
    lstrLocalName = funcGetLocalName()

    cstatusBar.Panels(1).Text = "�û���ţ�" & um�û����
    cstatusBar.Panels(2).Text = "�û�������" & um�û���
    cstatusBar.Panels(3).Text = "����վ����" & lstrLocalName
    
    mblnAutoUpgrade = False
    
    If Not pblnע�� Then
        '�жϸ��û���ע����Ϣ
        Dim lobjCheck As New cls�û����, lstrExpireDate As String
'        lstrExpireDate = lobjCheck.funcGetExpireDate()
'        If lstrExpireDate = "" Then
'            frm�����Ϣ¼��.Show 1
'            lstrExpireDate = lobjCheck.funcGetExpireDate()
'            If lstrExpireDate = "" Then
'                MsgBox "������ϵͳ����ʽ�û���ϵͳ�޷����У�", vbCritical, "ϵͳ��ʾ"
'                End
'            End If
'        End If
'        If CDate(lstrExpireDate) < Date Then
'            MsgBox "����ϵͳ�Ѿ�������ʹ�����ޣ���������ṩ����ϵ��", vbCritical, "ϵͳ��ʾ"
'            End
'        End If
'        If CDate(lstrExpireDate) < CDate("2050-01-01") Then
'            If DateDiff("d", Date, CDate(lstrExpireDate)) < 30 Then
'                MsgBox "����ʹ�����޻�ʣ��" & DateDiff("d", Date, CDate(lstrExpireDate)) & "�죡", vbInformation, "ϵͳ��ʾ"
'            End If
'        End If
    End If
    '��ȡ���
    Dim larrLic(1 To 16) As String
    Dim lstrSubSystem As String

    lstrSubSystem = lobjCheck.funcGetSubSystem()
    If lstrSubSystem = "" Then
        MsgBox "��û��ʹ���κ�ϵͳ���ܵ�Ȩ�ޣ����������Ӧ����ϵ��", vbInformation, "ϵͳ��ʾ"
        End
    End If

    larrLic(1) = "�������֤����,�������"
    larrLic(2) = "�ලִ������,ͻ���¼�"
    larrLic(3) = "������"
    larrLic(4) = "����֤����,����֤"
    larrLic(5) = "����������,�������"
    larrLic(6) = "�������"
    larrLic(7) = "�ƻ����߹���"
    larrLic(8) = "���ڹ���,�����칫����"
    larrLic(9) = "�շѹ���"
    larrLic(10) = "վ����ѯ,�쵼��ѯ"

    larrLic(11) = "�ʿع���"
    larrLic(12) = "����֪ͨ"
    larrLic(13) = "ҽ�ƻ�������"
    larrLic(14) = "ְҵ��������"
    larrLic(15) = "�������,����"
    larrLic(16) = "�������"

    pstr��ϵͳ��� = "ϵͳ����,��λ��������,"
    For i = 1 To Len(lstrSubSystem)
        If Mid(lstrSubSystem, i, 1) = "1" Then pstr��ϵͳ��� = pstr��ϵͳ��� + larrLic(i) + ","
    Next
    
'    If Not pblnע�� And Not pbln���� Then
'        '��ȡϵͳ�����еļ��ܹ�����������
'        Dim lstrDogServer As String
'        lstrDogServer = sffuncGetSetting("ϵͳ����", "���ݿ�����", "��������������")
'        If lstrDogServer = "" Then lstrDogServer = lstrServer
'
'        Call dafuncGetData("sp_addlinkedserver '" & lstrDogServer & "'")
'        Err.Clear
'        Set lobjRec = dafuncGetData("exec [" & lstrDogServer & "].master.dbo.ryCheck")
'        If Err.Number = 0 Then
'            Select Case lobjRec(0)
'            Case 1, 2, 3, 4, 5, 6
'                MsgBox "����������װ�����������°�ȫ���ʧ�ܡ�" & Chr(13) & Chr(10) & "�����°�װ�������������", vbCritical, "ϵͳ����"
'                End
'            Case Else
''                If (lobjRec(0) And (&H8000)) = 32768 Then
'                    '�ҵ��ˡ���ȡ��ɡ�
'                    Dim llngBit As Long
'                    Dim larrLic(1 To 15) As String
'                    larrLic(1) = "�������֤����,�������"
'                    larrLic(2) = "�ලִ������,ͻ���¼�"
'                    larrLic(3) = "������"
'                    larrLic(4) = "����֤����,����֤"
'                    larrLic(5) = "����������,�������"
'                    larrLic(6) = "�������"
'                    larrLic(7) = "�ƻ����߹���"
'                    larrLic(8) = "���ڹ���,�����칫����"
'                    larrLic(9) = "�շѹ���"
'                    larrLic(10) = "վ����ѯ,�쵼��ѯ"
'
'                    larrLic(11) = "�ʿع���"
'                    larrLic(12) = "����֪ͨ"
'                    larrLic(13) = "ҽ�ƻ�������"
'                    larrLic(14) = "ְҵ��������"
'                    larrLic(15) = "�������,����"
'                    pstr��ϵͳ��� = ""
'                    llngBit = &H4000
'                    For i = 1 To 15
'                        If (lobjRec(0) And llngBit) = llngBit Then
'                            pstr��ϵͳ��� = pstr��ϵͳ��� & larrLic(i) & ","
'                        End If
'                        llngBit = llngBit / 2
'                    Next
''                Else
''                    MsgBox "������������ʽ��ģ���ȫ���ʧ�ܣ�ϵͳ�޷����С�" & Chr(13) & Chr(10) & "���������Ӧ����ϵ��", vbCritical, "ϵͳ����"
''                    End
''                End If
'            End Select
'        Else
'            MsgBox "����������װ�����������°�ȫ���ʧ�ܡ�" & Chr(13) & Chr(10) & "�����°�װ������������򡣰�װǰȷ�����������������Ѱ�װSql Server2000��", vbCritical, "ϵͳ����"
'            End
'        End If
'        pstr��ϵͳ��� = "ϵͳ����,��λ��������," & pstr��ϵͳ���
'    ElseIf pbln���� Then
'        pstr��ϵͳ��� = ""
'        sub�����������
'        Timer1.Enabled = True
'    End If
    
    um��ϵͳ��� = pstr��ϵͳ���
    
    dasubSetQueryTimeout 6000

    On Error GoTo errHandle
    
    Dim llngWndProc As Long
    llngWndProc = SetWindowLong(Me.hWnd, GWL_WNDPROC, AddressOf funcClassing)
    pcolWndProc.Add llngWndProc, CStr(Me.hWnd)
    
    pobjƽ̨�ṹ.ƽ̨���� = um�û���� '��ƽ̨�ṹ���û���Ÿ�ֵ
    Set mobj�� = pobjƽ̨�ṹ.���������  'ȡ���û���ƽ̨�ṹ
    Set mobj�� = pobjƽ̨�ṹ.Operations
    Set mobj���� = pobjƽ̨�ṹ.Operation
    Set mobj��ѯ = pobjƽ̨�ṹ.Queries
    Set mobj��ѯ���� = pobjƽ̨�ṹ.Query
    Set mobj���� = pobjƽ̨�ṹ.Reports
    Set mobj������� = pobjƽ̨�ṹ.Report
    Set mobjSmartInfos = pobjƽ̨�ṹ.SmartInfos
    
    j = 1
    
    mintMnu = 1
    
    '��ʼ��������ͼƬ
    If um�û���� = "0000" Then
        mobj��.Filter = "������='ϵͳ' or ������='ϵͳ����'"
'        clblͨ�ò���(3).Enabled = True
    Else
        mobj��.Filter = "��������= 'ҵ����'"
    End If
    For ii = 1 To mobj��.RecordCount
        Select Case mobj��("������")
            Case "ϵͳ", "ϵͳ����"
                clblͨ�ò���(3).Enabled = True

            Case Else
        
                Load cmnuItem(mintMnu)
                Load cpicButton(mintMnu)
                cpicButton(mintMnu).Left = cpicButton(0).Left
                cpicButton(mintMnu).Top = cpicButton(mintMnu - 1).Top + cpicButton(mintMnu - 1).Height + 30
                cpicButton(mintMnu).Visible = True
                cmnuItem(mintMnu) = mobj��("������")
                cmnuItem(mintMnu).Left = cmnuItem(0).Left
                cmnuItem(mintMnu).Top = cpicButton(mintMnu).Top + 90
                cmnuItem(mintMnu).Visible = True
                mstrMnu(mintMnu) = mobj��("������")
                mintMnu = mintMnu + 1
                
        End Select
        sub��ʼ���ֵ��б� mobj��("������")
        
        mobj��.MoveNext
    Next
    mintMnu = mintMnu - 1
    
    '�ж��û��Ƿ��ֵ����Ա
    Set lobjRec = dafuncGetData("select * from ϵͳ����_�ֵ�_�û�������� where �û����='" & um�û���� & "'")
    If lobjRec.RecordCount = 0 Then clblͨ�ò���(4).Enabled = False
    
    mobj��.Filter = "��������= 'ҵ����'"
    
  
    On Error Resume Next
    If um�û���� = "0000" Then
        '�ж��Ƿ��һ�����С�
        Dim lobj���� As cls�û���������
        Set lobj���� = New cls�û���������
        lobj����.�û���� = "0000"
        lobj����.ҵ���� = "ϵͳ����"
        If lobj����.������ֵ("��һ������") <> "��" Then
            Call sub��������("ϵͳ����_����״̬����")
            '���������й�״̬��
            lobj����.sub���Ǽ���ֵ "��һ������", "��"
        End If
    End If
    
    Me.Caption = pstrSysName
    
    On Error Resume Next
    cimgPhoto.Picture = pmfunc��ȡͼƬ(um�û����, "Ա������")
    If cimgPhoto.Picture <> 0 Then
        Image5.Visible = False
    End If
    Image1.Picture = LoadPicture(App.Path & "\��ҳ-����.jpg")
    '�������ж�ע����Ϣ����¼ϵͳʱ�Ѿ��жϹ�һ��
    'Timer2.Enabled = True
    
    '������ʹ����
    'Timer3.Enabled = True
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
Private Sub sub��ʼ�������ѯ����()

    On Error Resume Next
    
    Set mobjSysAccObj = CreateObject("���ݷ�����.clsSystemBaseAccess")
    If Err <> 0 Then
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
    Else

        If mobjFrontQueryManager.ReportDataObject Is Nothing Then
            sub��ʼ�������ѯ����
        End If
    End If


    '������ѯ���档
    Dim llng������ As Long
    Dim llngWndProc As Long
    
    llng������ = mobjFrontQueryManager.funcStart(pstr��ϵͳ���)
    
    '�趨�򿪵Ĵ���Ϊ��������Ӵ��塣
    If llng������ <> -2 Then
        '�򼯺��м����������
        If Not sffunc�жϼ��ϼ�ֵ�Ƿ����(pcol��������, CStr(llng������)) Then
            On Error Resume Next
            SetParent llng������, Me.hWnd
            llngWndProc = SetWindowLong(llng������, GWL_WNDPROC, AddressOf funcClassing)
            pcolWndProc.Add llngWndProc, CStr(llng������)
            pcol��������.Add "����ͳ��", CStr(llng������)
            pcol�Ӵ�����.Add llng������, "����ͳ��"
            
            Call MoveWindow(llng������, ScaleX(1600, vbTwips, vbPixels), ScaleX(cimgBackground.Height - 60, vbTwips, vbPixels), ScaleX(Me.ScaleWidth - 1650, vbTwips, vbPixels), ScaleX(Me.ScaleHeight - cstatusBar.Height - cimgBackground.Height + 100, vbTwips, vbPixels), 1)
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
        If InStr(paraҵ������, "����֤����") > 0 Then
            Dim lobjFrm As New clsManageForm
            If lobjFrm.funcCheck("aaf3") <> "ab!&d3290" Then
                MsgBox "����汾����ȷ��", vbCritical, "ϵͳ��ʾ"
                Exit Sub
            End If
            llng������ = lobjFrm.funcStart(paraҵ������)
        Else
            Set lobj���� = CreateObject(mobj����("������") & "." & mobj����("����"))
            
            If paraҵ������ = "ϵͳ����_�����ѯȨ������" Then
                llng������ = lobj����.funcStart(paraҵ������, pstr��ϵͳ���)
            Else
                llng������ = lobj����.funcStart(paraҵ������)
            End If
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
                pcolWndProc.Add llngWndProc, CStr(llng������)
                pcolҵ������.Add lstrҵ����, CStr(llng������)
                pcol�Ӵ�����.Add llng������, paraҵ������
                pcol��������.Add paraҵ������, CStr(llng������)
                
                Call MoveWindow(llng������, ScaleX(1600, vbTwips, vbPixels), ScaleX(cimgBackground.Height - 60, vbTwips, vbPixels), ScaleX(Me.ScaleWidth - 1650, vbTwips, vbPixels), ScaleX(Me.ScaleHeight - cstatusBar.Height - cimgBackground.Height + 100, vbTwips, vbPixels), 1)

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
    If mblnAutoUpgrade = False Then     'ֻ�����������˳�ʱ�ŵ����ý��棬�Զ���������Ҫ
        frmExit.Show vbModal, Me
        If pblnCancel Then
            Cancel = True
            Exit Sub
        End If
    Else
        pblnExit = True
    End If
    
    sub�˳���ʹ����
    
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

                Set lobj���� = mobjFrontQueryManager
            ElseIf lstr������ = "��ѯͳ��" Then

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
            'Unload frmƽ̨����
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
'    WriteErrLog                                '��������Ϣд�����ݿ�
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
'            Unload frm����֪ͨ
'            Unload frm���ڹ���
            Unload frmϵͳ����
            Unload frm�ֵ��б�
            Call Main
        Else

            subExit
            
        End If
    End If
End Sub
Private Sub subExit()
    On Error Resume Next
    X.subCloseDatabase
'    Unload frm����֪ͨ
'    Unload frm���ڹ���
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
    Frame1.Height = Me.ScaleHeight - cstatusBar.Height - cimgBackground.Height + 80
    Image5.Top = Frame1.Height - Image5.Height - 100
    cimgBackground.Width = Me.ScaleWidth - cimgBackground.Left
    Frame2.Top = Me.ScaleHeight - cstatusBar.Height - Frame2.Height - 120
    Frame2.Left = Me.ScaleWidth - Frame2.Width - 120
    cimgPhoto.Top = Frame1.Height - cimgPhoto.Height - 100
    
    subResizeChild
End Sub
Private Sub sub��ʼ���ֵ��б�(para���� As String)
    Dim i As Integer, j As Integer
    Dim lobjRec As Object
    
    On Error GoTo errHandle
    
    mobj��.Filter = ""
    mobj��.Filter = "��������" & " ='" & para���� & "' "
    For i = 1 To mobj��.RecordCount
        '�޸ģ��жϵ�ǰ��������ҵ�����Ƿ��ڼ��ܹ���ɷ�Χ�ڡ�
        mobj����.Filter = ""
        mobj����.Filter = "��������" & "='" & mobj��.Fields("��������") & "'"
        If mobj����.RecordCount > 0 Then
            If pstr��ϵͳ��� = "" Or InStr(pstr��ϵͳ���, mobj����.Fields("ҵ����") & ",") > 0 Then
                If Not sffunc�жϼ��ϼ�ֵ�Ƿ����(pcol�ֵ伯, mobj����.Fields("ҵ����").Value) Then
                    '�жϸ�ҵ���Ƿ��в��������ֵ䡣
                    Set lobjRec = dafuncGetData("select * from ϵͳ����_�ֵ�_�ֵ���б� where ҵ����='" & mobj����.Fields("ҵ����").Value & "' and ����='������'")
                    If lobjRec.RecordCount > 0 Then
                        pcol�ֵ伯.Add mobj����.Fields("ҵ����").Value, mobj����.Fields("ҵ����").Value
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
    
    '����ò�����Ĳ���
    ii = 1
    mobj��.Filter = ""
    mobj��.Filter = "��������" & " ='" & para���� & "' "
    For i = 1 To 10
        cmnuSubItem(i).Visible = False
        cimgButton(i).Visible = False
        cmnuSubItem(i) = ""
    Next

    For i = 1 To mobj��.RecordCount
        '�޸ģ��жϵ�ǰ��������ҵ�����Ƿ��ڼ��ܹ���ɷ�Χ�ڡ�
        mobj����.Filter = ""
        mobj����.Filter = "��������" & "='" & mobj��.Fields("��������") & "'"
        If mobj����.RecordCount > 0 Then
            If pstr��ϵͳ��� = "" Or InStr(pstr��ϵͳ���, mobj����.Fields("ҵ����") & ",") > 0 Then
                cmnuSubItem(ii) = IIf(Len(mobj��("��������")) > 6, Left(mobj��("��������"), 6) & "...", mobj��("��������"))
                cmnuSubItem(ii).Top = cmnuSubItem(ii - 1).Top + cmnuSubItem(ii - 1).Height + 150
                cmnuSubItem(ii).Left = cmnuSubItem(0).Left
                cmnuSubItem(ii).Visible = True
                cimgButton(ii).Top = cmnuSubItem(ii).Top - 30
                cimgButton(ii).Visible = True
                cimgButton(ii).Left = cimgButton(0).Left
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
            pcolWndProc.Add llngWndProc, CStr(llng������)
            pcol��������.Add "�ֵ����", CStr(llng������)
            pcol�Ӵ�����.Add llng������, "�ֵ����"
            Call MoveWindow(llng������, ScaleX(1600, vbTwips, vbPixels), ScaleX(cimgBackground.Height - 60, vbTwips, vbPixels), ScaleX(Me.ScaleWidth - 1650, vbTwips, vbPixels), ScaleX(Me.ScaleHeight - cstatusBar.Height - cimgBackground.Height + 100, vbTwips, vbPixels), 1)
            
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
    
    lstrTime = "2008-12-31"
    
    If lstrTime < Format(Now, "yyyy-mm-dd") Then
        MsgBox "�Բ������ʹ�������ѵ������������Ӧ����ϵ��", vbCritical, "ϵͳ��ʾ"
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

        Call MoveWindow(llngHwnd, ScaleX(1600, vbTwips, vbPixels), ScaleX(cimgBackground.Height - 60, vbTwips, vbPixels), ScaleX(Me.ScaleWidth - 1650, vbTwips, vbPixels), ScaleX(Me.ScaleHeight - cstatusBar.Height - cimgBackground.Height + 100, vbTwips, vbPixels), 1)
        'Call MoveWindow(llngHwnd, ScaleX(700, vbTwips, vbPixels), ScaleX(350, vbTwips, vbPixels), ScaleX(Me.ScaleWidth - 700, vbTwips, vbPixels), ScaleX(Me.ScaleHeight - 400, vbTwips, vbPixels), 1)

    Next
End Sub

Private Sub Timer2_Timer()
    On Error Resume Next
    Dim lforeWin As Long
    '����û�����Ч����
    Dim lstrDate As String, lstrExpireDate As String
    Dim lobjCheck As New cls�û����

    lstrDate = (dafuncGetData("select getdate()").Fields(0))
    lstrExpireDate = lobjCheck.funcGetExpireDate()
    If lstrExpireDate = "" Then
        MsgBox "������ϵͳ����ʽ�û���ϵͳ�޷����У�", vbCritical, "ϵͳ��ʾ"
        End
    End If
    If lstrExpireDate = "��֤��Ϣ����" Then
        MsgBox "ϵͳ����֤��Ϣ�����޷��������У�", vbCritical, "ϵͳ��ʾ"
        End
    End If
    If CDate(lstrDate) > CDate(lstrExpireDate) Then
        MsgBox "��ǰϵͳ�Ѿ�ʧЧ���޷����У�", vbCritical, "ϵͳ��ʾ"
        End
    End If
    
'    lforeWin = GetForegroundWindow()
'    If Me.hWnd = lforeWin Then
'
'        sub��ʾ���칤��
'
'    End If
    Timer2.Enabled = False
End Sub

Private Sub Timer3_Timer()
    On Error Resume Next
    Timer3.Enabled = False
    sub��¼��ʹ����
End Sub
