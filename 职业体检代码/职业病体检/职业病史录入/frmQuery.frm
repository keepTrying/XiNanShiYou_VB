VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmQuery 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "��ѯ"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   5025
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame1 
      Caption         =   "��ѯ����"
      Height          =   3375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4815
      Begin VB.TextBox ctxtϵͳ��� 
         Height          =   350
         Left            =   2040
         TabIndex        =   19
         Top             =   2760
         Width           =   2535
      End
      Begin VB.CheckBox cchkType 
         Caption         =   "ϵͳ���(�����)"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   18
         Top             =   2800
         Width           =   1935
      End
      Begin VB.CheckBox cchkType 
         Caption         =   "�Թܱ�ţ�"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   17
         Top             =   3720
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox ctxt�Թܱ�� 
         Height          =   350
         Left            =   1680
         TabIndex        =   16
         Top             =   3720
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.TextBox ctxt��쵥�� 
         Height          =   350
         Left            =   2040
         TabIndex        =   15
         Top             =   3240
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CheckBox cchkType 
         Caption         =   "��쵥�ţ�"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   14
         Top             =   3280
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox ctxt���� 
         Height          =   350
         Left            =   2040
         TabIndex        =   13
         Top             =   2300
         Width           =   2535
      End
      Begin VB.CheckBox cchkType 
         Caption         =   "��    ����"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   12
         Top             =   2300
         Width           =   1455
      End
      Begin VB.CheckBox cchkType 
         Caption         =   "������ڴӣ�"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   1455
      End
      Begin VB.CheckBox cchkType 
         Caption         =   "�������ƣ�"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   1455
      End
      Begin VB.ComboBox ccmbSheet 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2040
         TabIndex        =   6
         Top             =   1320
         Width           =   2535
      End
      Begin VB.CheckBox cchkType 
         Caption         =   "��λ���ƣ�"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox ctxtUnit 
         Enabled         =   0   'False
         Height          =   350
         Left            =   2040
         TabIndex        =   4
         Top             =   1800
         Width           =   2055
      End
      Begin VB.CommandButton ccmd��λ 
         Caption         =   "..."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "��λ��λ"
         Top             =   1800
         Width           =   495
      End
      Begin MSComCtl2.DTPicker cdtp��ʼ���� 
         Height          =   300
         Left            =   2040
         TabIndex        =   9
         Top             =   360
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   59834368
         CurrentDate     =   37056
      End
      Begin MSComCtl2.DTPicker cdtp�������� 
         Height          =   300
         Left            =   2040
         TabIndex        =   10
         Top             =   840
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   59834368
         CurrentDate     =   37056
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Index           =   1
         Left            =   1080
         TabIndex        =   11
         Top             =   960
         Width           =   360
      End
   End
   Begin VB.CommandButton ccmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   375
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton ccmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3600
      Width           =   1215
   End
End
Attribute VB_Name = "frmQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'��ѯ����
Public pstr��ʼ���� As String
Public pstr��ֹ���� As String
Public pstr�������� As String
Public pstr��λ���� As String
Public pstr���� As String
Public pstr��쵥�� As String
Public pstr�Թܱ�� As String
Public pstrϵͳ��� As String
Public pstr���֤�� As String
Public pblnOk As Boolean

Private Sub cchkType_Click(Index As Integer)
    On Error Resume Next
    If cchkType(0).Value = 1 Then
        cdtp��ʼ����.Enabled = True
        cdtp��������.Enabled = True
        cdtp��ʼ����.SetFocus
    Else
        cdtp��ʼ����.Enabled = False
        cdtp��������.Enabled = False
    End If
    If cchkType(1).Value = 1 Then
        ccmbSheet.Enabled = True
        ccmbSheet.SetFocus
    Else
        ccmbSheet.Enabled = False
    End If
    If cchkType(2).Value = 1 Then
        ctxtUnit.Enabled = True
        ccmd��λ.Enabled = True
        ctxtUnit.SetFocus
    Else
        ctxtUnit.Enabled = False
        ccmd��λ.Enabled = False
    End If
    
    If cchkType(3).Value = 1 Then
        ctxt����.Enabled = True
        ctxt����.SetFocus
    Else
        ctxt����.Enabled = False
    End If
    
    If cchkType(4).Value = 1 Then
        ctxt��쵥��.Enabled = True
        ctxt��쵥��.SetFocus
    Else
        ctxt��쵥��.Enabled = False
    End If
    
    If cchkType(5).Value = 1 Then
        ctxt�Թܱ��.Enabled = True
        ctxt�Թܱ��.SetFocus
    Else
        ctxt�Թܱ��.Enabled = False
    End If
    If cchkType(6).Value = 1 Then
        ctxtϵͳ���.Enabled = True
        ctxtϵͳ���.SetFocus
    Else
        ctxtϵͳ���.Enabled = False
    End If
End Sub

Private Sub ccmdCancel_Click()
    pblnOk = False
    Unload Me

End Sub

Private Sub ccmdOk_Click()
    On Error GoTo errHandler
    If cchkType(0).Value = 1 Then
        pstr��ʼ���� = Format(cdtp��ʼ����.Value, "yyyy-mm-dd")
        pstr��ֹ���� = Format(cdtp��������.Value, "yyyy-mm-dd")
    Else
        pstr��ʼ���� = ""
        pstr��ֹ���� = ""
    End If
    If cchkType(1).Value = 1 Then
        pstr�������� = ccmbSheet.Text
    Else
        pstr�������� = ""
    End If
    If cchkType(2).Value = 1 Then
        pstr��λ���� = ctxtUnit.Text
    Else
        pstr��λ���� = ""
    End If
    
    If cchkType(3).Value = 1 Then
        pstr���� = ctxt����.Text
    Else
        pstr���� = ""
    End If
    
    If cchkType(4).Value = 1 Then
        pstr��쵥�� = ctxt��쵥��.Text
    Else
        pstr��쵥�� = ""
    End If
    
    If cchkType(5).Value = 1 Then
        pstr�Թܱ�� = ctxt�Թܱ��.Text
    Else
        pstr�Թܱ�� = ""
    End If
    If cchkType(6).Value = 1 Then
        pstrϵͳ��� = ctxtϵͳ���.Text
    Else
        pstrϵͳ��� = ""
    End If
    pblnOk = True
    Unload Me
    Exit Sub
errHandler:
    sfsub������ "ְҵ��ʷ¼�����", "frmquery", "ccmdok_click", Err.Number, Err.Description, True
End Sub

Private Sub ccmd��λ_Click()
    Dim lobj�ӿ� As Object
    Dim lobjRec As Object
    
    On Error GoTo errHandler
    
    Set lobj�ӿ� = CreateObject("��λ����ҵ��.ClsUnitInterface")
    Set lobjRec = lobj�ӿ�.func��λ�򵥶�λ(Screen.Width / 2, Screen.Height / 2)
    
    If lobjRec Is Nothing Then
        ctxtUnit.SetFocus
        Exit Sub
    End If
    
    ctxtUnit = lobjRec!��λ����
    ctxtUnit.SetFocus
    Exit Sub
errHandler:
    sfsub������ "ְҵ��ʷ¼�����", "frmquery", "ccmd��λ_click", Err.Number, Err.Description, True
End Sub

Private Sub ctxtϵͳ���_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And ctxtϵͳ��� <> "" Then
        ccmdOk_Click
    End If
End Sub

Private Sub Form_Load()
    Dim i As Long
    On Error GoTo errHandler
    If pstr��ʼ���� = "" Then
        cdtp��ʼ����.Value = Format(DateAdd("d", -7, Date), "yyyy/mm/dd") '��ʼ����Ϊ��ǰ���ڵ�ǰ7��
        cchkType(0).Value = 0
    Else
        cdtp��ʼ����.Value = Format(pstr��ʼ����, "yyyy/mm/dd")
        cchkType(0).Value = 1
    End If
    If pstr��ֹ���� = "" Then
        cdtp��������.Value = Date
    Else
        cdtp��������.Value = Format(pstr��ֹ����, "yyyy/mm/dd")
    End If
    
    
    '��ȡ�����������ơ�
    Dim lobj����ģ�弯 As Object
    Dim lcolInfo As Collection
    Set lobj����ģ�弯 = CreateObject("ְҵ������.ClsMedicalExamTemplateSet")
    Set lcolInfo = lobj����ģ�弯.Ԫ�ؼ�
    
    ccmbSheet.Clear
    If lcolInfo.Count > 0 Then
        ccmbSheet.AddItem ""
    End If
    For i = 1 To lcolInfo.Count
        ccmbSheet.AddItem lcolInfo(i)
    Next i
    ccmbSheet.Text = pstr��������
    If pstr�������� = "" Then
        cchkType(1).Value = 0
    Else
        cchkType(1).Value = 1
    End If
    
    ctxtUnit = pstr��λ����
    If pstr��λ���� = "" Then
        cchkType(2).Value = 0
    Else
        cchkType(2).Value = 1
    End If
    
    ctxt���� = pstr����
    If pstr���� = "" Then
        cchkType(3).Value = 0
    Else
        cchkType(3).Value = 1
    End If
    
    ctxt��쵥�� = pstr��쵥��
    If pstr��쵥�� = "" Then
        cchkType(4).Value = 0
    Else
        cchkType(4).Value = 1
    End If

    ctxt�Թܱ�� = pstr�Թܱ��
    If pstr�Թܱ�� = "" Then
        cchkType(5).Value = 0
    Else
        cchkType(5).Value = 1
    End If
    Exit Sub
errHandler:
    sfsub������ "ְҵ��ʷ¼�����", "frmquery", "form_load", Err.Number, Err.Description, True
End Sub
