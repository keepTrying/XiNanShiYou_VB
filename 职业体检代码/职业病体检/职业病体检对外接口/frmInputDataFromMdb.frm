VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInputDataFromMdb 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����λ���ݵ���"
   ClientHeight    =   7455
   ClientLeft      =   1170
   ClientTop       =   1215
   ClientWidth     =   10470
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   10470
   Begin VB.Frame Frame2 
      Caption         =   "���ݷ���������"
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   555
      Left            =   120
      TabIndex        =   23
      Top             =   1260
      Width           =   5235
      Begin VB.OptionButton coptInUnit 
         Caption         =   "վ�����ݷ�����"
         Height          =   315
         Left            =   2880
         TabIndex        =   25
         Top             =   180
         Value           =   -1  'True
         Width           =   1755
      End
      Begin VB.OptionButton coptOutUnit 
         Caption         =   "վ�����ݷ�����"
         Height          =   315
         Left            =   180
         TabIndex        =   24
         Top             =   180
         Width           =   1995
      End
   End
   Begin VB.CommandButton ccmdPrefech 
      Caption         =   "Ԥ��ȡ����(&W)"
      Height          =   375
      Left            =   6960
      TabIndex        =   2
      Top             =   840
      Width           =   1515
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ԥ������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   3195
      Left            =   120
      TabIndex        =   20
      Top             =   3780
      Width           =   10335
      Begin VSFlex6DAOCtl.vsFlexGrid cgrdPreviewData 
         Height          =   2835
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   5001
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   14737632
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   12
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0   'False
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
      End
   End
   Begin MSComctlLib.ProgressBar cprgDatatranform 
      Height          =   375
      Left            =   0
      TabIndex        =   18
      Top             =   7080
      Visible         =   0   'False
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame cfraFiltrateCondition 
      Caption         =   "���ݵ�������"
      ForeColor       =   &H00800000&
      Height          =   1785
      Left            =   120
      TabIndex        =   16
      Top             =   1920
      Width           =   6645
      Begin MSComCtl2.DTPicker cdtpBeginDate 
         Height          =   375
         Left            =   1320
         TabIndex        =   4
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   23527425
         CurrentDate     =   36951
      End
      Begin VB.CheckBox cchkSystemCode 
         Caption         =   "ϵͳ���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CheckBox cchkUnitName 
         Caption         =   "��λ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox cchkMedicalDate 
         Caption         =   "�������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   300
         Value           =   1  'Checked
         Width           =   1245
      End
      Begin VB.CommandButton ccmdLocateUnit 
         Caption         =   "��λ��λ(&L)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5280
         TabIndex        =   8
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox ctxtUnit 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   7
         Top             =   720
         Width           =   3975
      End
      Begin VB.TextBox ctxtBeginCode 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   10
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox ctxtEndCode 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   11
         Top             =   1200
         Width           =   2415
      End
      Begin MSComCtl2.DTPicker cdtpEndDate 
         Height          =   375
         Left            =   3480
         TabIndex        =   5
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   23527425
         CurrentDate     =   36951
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   3120
         TabIndex        =   22
         Top             =   360
         Width           =   180
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   3840
         TabIndex        =   17
         Top             =   1320
         Width           =   180
      End
   End
   Begin VB.Frame cfraSelectInputData 
      Caption         =   "ѡ��Ҫ���������"
      ForeColor       =   &H00800000&
      Height          =   1905
      Left            =   6840
      TabIndex        =   15
      Top             =   1800
      Width           =   3555
      Begin VB.ListBox clstDataType 
         Height          =   1530
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   12
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.TextBox ctxtDataSource 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1260
      TabIndex        =   0
      Top             =   840
      Width           =   4155
   End
   Begin VB.CommandButton ccmdBrowse 
      Caption         =   "����ļ�(&B)"
      Height          =   375
      Left            =   5460
      TabIndex        =   1
      Top             =   840
      Width           =   1275
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8340
      Top             =   540
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar csbMain 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   19
      Top             =   7080
      Width           =   10470
      _ExtentX        =   18468
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   18415
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar ctbMain 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   10470
      _ExtentX        =   18468
      _ExtentY        =   1111
      ButtonWidth     =   820
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "cimg��ťͼ��"
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList cimg��ťͼ�� 
      Left            =   9300
      Top             =   660
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog ccdgBrowse 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label ccmdDataSource 
      AutoSize        =   -1  'True
      Caption         =   "������Դ��"
      Height          =   180
      Left            =   120
      TabIndex        =   14
      Top             =   960
      Width           =   900
   End
End
Attribute VB_Name = "frmInputDataFromMdb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mobjGUI  As cls����ͨ�ö��� '���ڳ�ʼ����������
Attribute mobjGUI.VB_VarHelpID = -1
Private mobj����ӿ� As ClsManageTransmission  'ҵ�����

Private mstrϵͳ��Ź̶�����  As String

Public pblnInUse As Boolean                    '���������Ƿ��Ѽ��ء�������Ҫʹ�á�

Private Sub coptInUnit_Click()
    On Error GoTo errHandler
    
    MousePointer = 11
    '������ݷ��ࡣ
    clstDataType.Clear
    
    '������ʱ���ܲ�����
    cfrafiltrateCondition.Enabled = False
    cfraSelectInputData.Enabled = False
            
    subԤ��ȡ����
        
    '�ָ����档
    ctbMain.Buttons(1).Enabled = True
    cfrafiltrateCondition.Enabled = True
    cfraSelectInputData.Enabled = True
    
    MousePointer = 0
    Exit Sub
errHandler:
    MousePointer = 0
    sfsub������ "������ӿڲ���", "frmInputDataFromMdb", "coptInUnit_Click", Err.Number, Err.Description, False
    Exit Sub
    Resume
End Sub

Private Sub coptOutUnit_Click()
    On Error GoTo errHandler
    
    MousePointer = 11
    '������ݷ�Χ��
    clstDataType.Clear
    cfrafiltrateCondition.Enabled = False
    cfraSelectInputData.Enabled = False
            
    subԤ��ȡ����
        
    ctbMain.Buttons(1).Enabled = True
    
    cfrafiltrateCondition.Enabled = True
    cfraSelectInputData.Enabled = True
    
    MousePointer = 0
    Exit Sub
errHandler:
    MousePointer = 0
    sfsub������ "������ӿڲ���", "frmInputDataFromMdb", "coptOutUnit_Click", Err.Number, Err.Description, False
    Exit Sub
    Resume
End Sub



Private Sub ctxtBeginCode_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ctxtEndCode.SetFocus
    End If
End Sub

Private Sub ctxtDataSource_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        If ctxtDataSource.Text <> "" Then
            ccmdPrefech.Enabled = True
            ccmdPrefech.SetFocus
        End If
    End If
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    ctxtDataSource.SetFocus
    ctxtDataSource.SelStart = Len(ctxtDataSource)
    ctxtDataSource.SelLength = 0
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 39 Then
        '���������롰'����
        KeyAscii = 0
    End If

End Sub
'���ܣ���ʼ�����档
'���ߣ����ơ�
Private Sub Form_Load()
    Dim lobjRec As Object
    Dim lcol��������ť As New Collection
    
    On Error GoTo errHandler
    pblnInUse = True
    
    '��������ͨ�ö��󣬳�ʼ����������
    Set mobjGUI = New cls����ͨ�ö���
    With lcol��������ť
        .Add "Ԥ��(&R)108"
        .Add "����(&I)112"
        .Add "|"
        .Add "�˳�"
    End With
    With mobjGUI
        Set .Form = Me
        Set .c������ = ctbMain
    End With
    mobjGUI.subInitialize lcol��������ť, ""
    
    '����������ӿڶ���
    Set mobj����ӿ� = CreateObject("������ӿڲ���.ClsManageTransmission")
    
    '�ӹ���վ�����ļ��л�ȡ��¼���ϴε����ļ���
    ctxtDataSource.Text = mobj����ӿ�.����վ����.�ڲ������ļ�
    
    '�ڸ�������check��δ��ѡ��ʱ,����������򲻿���.
    ctxtBeginCode.Enabled = False
    ctxtEndCode.Enabled = False
    ctxtUnit.Enabled = False
    ccmdLocateUnit.Enabled = False
    cdtpBeginDate.Value = Format(Date, "yyyy-mm-dd")
    cdtpEndDate.Value = Format(Date, "yyyy-mm-dd")
    
    '����,Ԥ����ť��ѡ���ļ���ű�Ϊ����.
    If Len(ctxtDataSource.Text) = 0 Then
        ccmdPrefech.Enabled = False
    End If
    If ctxtDataSource.Text = "" Then
        ccmdPrefech.Enabled = False
    End If
    ctbMain.Buttons(2).Enabled = False
    cfrafiltrateCondition.Enabled = False
    cfraSelectInputData.Enabled = False
    ctbMain.Buttons(1).Enabled = False
    
    cdtpBeginDate.Value = Date
    cdtpEndDate.Value = Date
    
    '��ȡϵͳ��Ź̶����֡�
    Dim lobj��� As Object '�����󣬻�ȡϵͳ��ŵĹ̶����֡�
    Set lobj��� = CreateObject("�����󲿼�.clsMedicalExam")
    mstrϵͳ��Ź̶����� = lobj���.ϵͳ��Ź̶�����
    Set lobj��� = Nothing
    
    Exit Sub
errHandler:
    sfsub������ "������ӿڲ���", "frmInputDataFromMdb", "Form_Load", Err.Number, Err.Description, False
    Exit Sub
    Resume
End Sub

'���ܣ� �û�����������check��ʱ��ʱ������ؼ��ڿ��úͲ����õ�״̬֮��仯��
'���ߣ� ���ơ�
Private Sub cchkMedicalDate_Click()
    On Error GoTo errHandler
    If cchkMedicalDate.Value = 1 Then
        cdtpBeginDate.Enabled = True
        cdtpEndDate.Enabled = True
        cdtpBeginDate.SetFocus
    Else
        cdtpBeginDate.Enabled = False
        cdtpEndDate.Enabled = False
    End If
    Exit Sub
errHandler:
    sfsub������ "������ӿڲ���", "frmInputDataFromMdb", "cchkMedicalDate_Click", Err.Number, Err.Description, False
End Sub

'���ܣ� ���û����ϵͳ���check��ʱ������ϵͳ���¼���Ŀ���״̬��
'���ߣ� ���ơ�
Private Sub cchkSystemCode_Click()
    On Error GoTo errHandler
    If cchkSystemCode.Value = 1 Then
        ctxtBeginCode.Enabled = True
        ctxtEndCode.Enabled = True
        ctxtBeginCode.SetFocus
        ctxtBeginCode.SelStart = Len(ctxtBeginCode)
        ctxtBeginCode.SelLength = 0
    Else
        ctxtBeginCode.Enabled = False
        ctxtEndCode.Enabled = False
    End If
    Exit Sub
errHandler:
    sfsub������ "������ӿڲ���", "frmInputDataFromMdb", "cchkSystemCode_Click", Err.Number, Err.Description, False
End Sub

Private Sub ctxtBeginCode_GotFocus()
    On Error Resume Next
    If Trim(ctxtBeginCode) = "" Then
        ctxtBeginCode.Text = mstrϵͳ��Ź̶�����
        ctxtBeginCode.SelStart = Len(ctxtBeginCode)
        ctxtBeginCode.SelLength = 0
    End If
End Sub

Private Sub ctxtEndCode_GotFocus()
    On Error Resume Next
    If Trim(ctxtEndCode) = "" Then
        ctxtEndCode.Text = mstrϵͳ��Ź̶�����
        ctxtEndCode.SelStart = Len(ctxtEndCode)
        ctxtEndCode.SelLength = 0
    End If

End Sub

'���ܣ� ���û�����λ����check��ʱ��������λ����¼���Ŀ���״̬��
'���ߣ� ���ơ�
Private Sub cchkUnitName_Click()
    On Error GoTo errHandler
    If cchkUnitName.Value = 1 Then
        ctxtUnit.Enabled = True
        ccmdLocateUnit.Enabled = True
        ctxtUnit.SetFocus
    Else
        ctxtUnit.Enabled = False
        ccmdLocateUnit.Enabled = False
    End If
    Exit Sub
errHandler:
    sfsub������ "������ӿڲ���", "frmInputDataFromMdb", "cchkUnitName_Click", Err.Number, Err.Description, False
End Sub

'���ܣ�����ͨ����λ��λ�õ��ĵ�λ���ƣ���ʾ�ڵ�λ�����ı����У�����λ����֮����Ӣ�Ķ��ŷָ���
'���ߣ� ���ơ�
Private Sub ccmdLocateUnit_Click()
    On Error GoTo errHandler
    Dim lobj������ As Object
    Dim lobjRec As Object
    Dim lstrUnit As String
    
    Set lobj������ = CreateObject("�����󲿼�.clsManageMedicalExam")
    Set lobjRec = lobj������.func��λ��λ
    
    If Not lobjRec Is Nothing Then
        If lobjRec.RecordCount > 0 Then
            lstrUnit = lobjRec.Fields("��λ����").Value
        End If
    End If
    If Len(lstrUnit) >= 1 Then
        If Len(ctxtUnit.Text) >= 1 Then
            ctxtUnit.Text = ctxtUnit.Text & "," & lstrUnit
        Else
            ctxtUnit.Text = lstrUnit
        End If
    End If
    Exit Sub
errHandler:
    sfsub������ "������ӿڲ���", "frmInputDataFromMdb", "ccmdLocateUnit_Click", Err.Number, Err.Description, False
End Sub

'���ܣ����ҵ�λ���Ƽ��ִ��е����Ķ��ţ�����Ӣ�Ķ����滻����
'���ߣ� ���ơ�
Private Sub ctxtUnit_LostFocus()
    On Error GoTo errHandler
    Dim i As Integer
    Dim lstrUnit As String
    
    lstrUnit = ctxtUnit.Text
    For i = 1 To Len(lstrUnit)
        If Mid(lstrUnit, i, 1) = "��" Then
            Mid(lstrUnit, i, 1) = ","
        End If
    Next i
    ctxtUnit.Text = lstrUnit
    Exit Sub
errHandler:
    sfsub������ "������ӿڲ���", "frmInputDataFromMdb", "ccmdLocateUnit_Click", Err.Number, Err.Description, False
End Sub

'����:  Ԥ����MDB�ļ��еĵ��������͵�������,����ʾ�ڽ����ϡ�
'���ߣ� ���ơ�
Private Sub ccmdPrefech_Click()
    On Error GoTo errHandler
    
    MousePointer = 11
    cfrafiltrateCondition.Enabled = False
    cfraSelectInputData.Enabled = False
    Frame2.Enabled = False
    
    '��ָ�����ļ�������copy������ִ�е�·�ļ����¡�
    If Len(ctxtDataSource.Text) = 0 Then
        Err.Number = 6666
        Err.Description = "Ԥ��ȡ����ǰ����������ȷ�ĵ����ļ���·����"
        GoTo errHandler
    Else
        mobj����ӿ�.sub����׼�� ctxtDataSource.Text
    End If
        
    subԤ��ȡ����
        
    ctbMain.Buttons(1).Enabled = True
    ctbMain.Buttons(2).Enabled = True
    cfrafiltrateCondition.Enabled = True
    cfraSelectInputData.Enabled = True
    Frame2.Enabled = True
    MousePointer = 0
    Exit Sub
errHandler:
    MousePointer = 0
    sfsub������ "������ӿڲ���", "frmInputDataFromMdb", "ccmdPrefech_Click", Err.Number, Err.Description, False
    Exit Sub
    Resume
End Sub
Private Sub subԤ��ȡ����()
    Dim lobjRec As Object
    Dim i As Integer
    
    On Error GoTo errHandler
    '��ʼ��"���ݵ�������"frame�ڵĸ��
    cdtpBeginDate.Value = Date
    cdtpEndDate.Value = Date
    ctxtBeginCode = ""
    ctxtEndCode = ""
    ctxtUnit = ""
    Set lobjRec = mobj����ӿ�.Func��ȡmdb���ݽ����ļ��е����ݷ�Χ
    While Not lobjRec.EOF
        Select Case lobjRec.Fields("��Χ��").Value
            Case "��ʼ����"
                cdtpBeginDate.Year = DatePart("yyyy", lobjRec.Fields("��Χֵ").Value)
                cdtpBeginDate.Month = DatePart("m", lobjRec.Fields("��Χֵ").Value)
                cdtpBeginDate.Day = DatePart("d", lobjRec.Fields("��Χֵ").Value)
            Case "��������"
                cdtpEndDate.Year = DatePart("yyyy", lobjRec.Fields("��Χֵ").Value)
                cdtpEndDate.Month = DatePart("m", lobjRec.Fields("��Χֵ").Value)
                cdtpEndDate.Day = DatePart("d", lobjRec.Fields("��Χֵ").Value)
            Case "��λ���Ƽ�"
                ctxtUnit.Text = lobjRec.Fields("��Χֵ").Value
            Case "��ϵͳ���"
                If Len(lobjRec.Fields("��Χֵ").Value) <> 0 Then
                    ctxtBeginCode.Text = lobjRec.Fields("��Χֵ").Value
                End If
            Case "��ϵͳ���"
                If Len(lobjRec.Fields("��Χֵ").Value) <> 0 Then
                    ctxtEndCode.Text = lobjRec.Fields("��Χֵ").Value
                End If
        End Select
        lobjRec.MoveNext
    Wend
        
    '��MDB���ݿ��а��������ݷ���������������ʱ��¼�ڱ����������ͱ��У��������ݷ����б���С�
    Set lobjRec = mobj����ӿ�.Func��ȡmdb�ļ��е����ݷ����嵥(IIf(coptInUnit, "վ�����ݷ�����", "վ�����ݷ�����"))
    clstDataType.Clear
    While Not lobjRec.EOF
        clstDataType.AddItem lobjRec.Fields("���ݷ�����").Value
        clstDataType.Selected(clstDataType.NewIndex) = True
        lobjRec.MoveNext
    Wend
    
    Exit Sub
errHandler:
    sfsub������ "������ӿڲ���", "frmInputDataFromMdb", "subԤ��ȡ����", Err.Number, Err.Description, True
    Exit Sub
    Resume

End Sub

'���ܣ� �����ļ����Ҵ��ڡ�
'���ߣ� ���ơ�
Private Sub ccmdBrowse_Click() ' ���á�CancelError��Ϊ True
    On Error GoTo errHandler
    ccdgBrowse.CancelError = True
    ' ���ñ�־
    ccdgBrowse.Flags = cdlOFNHideReadOnly
    ' ���ù�����
    ccdgBrowse.Filter = "All Files (*.*)|*.*|Access file" & _
        "(*.mdb)|*.mdb|Batch Files (*.bat)|*.bat"
    ccdgBrowse.FilterIndex = 2
    ccdgBrowse.ShowOpen
    ctxtDataSource.Text = ccdgBrowse.FileName
    
    If ctxtDataSource.Text = "" Then
        ccmdPrefech.Enabled = False
    Else
        ccmdPrefech.Enabled = True
    End If
    Exit Sub
errHandler:
    ' �û����ˡ�ȡ������ť
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    Set mobj����ӿ� = Nothing
    Set mobjGUI = Nothing
    pblnInUse = False
    
End Sub

'���ܣ� ����,Ԥ�������ļ��е�����.
'���ߣ� ���ơ�
Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    On Error GoTo errHandler
    Dim lcolRange As Collection '��ŵ����Ԥ�����ݵĹ���������
    Dim lobjRec As Object '�������������
    Dim lcolType As Collection
    Dim i As Integer
    
    Set lcolRange = funcCalCon
    Select Case Operate
        Case "Ԥ��"
            '��ȡ����������ʾ��Ԥ�����ݿ���
            cgrdPreviewData.Rows = 1
            Set lobjRec = mobj����ӿ�.Func�鿴����(lcolRange, 0) '(0����/ 1����)
            gfsubLoadGridFromRec cgrdPreviewData, lobjRec, , "�����������,ϵͳ���,������ݺ���,����,�Ա�,��������,��λ����,�������,��������,������,��Ϻʹ������,���ҽʦ"
            cgrdPreviewData.Rows = cgrdPreviewData.Rows - 1
            
            cprgDatatranform.Value = 0
            Cancel = True
            
        Case "����"
            cprgDatatranform.Value = 0
            cprgDatatranform.Visible = True
            MousePointer = 11
            csbMain.Panels(1) = "���ڵ��룬���Ժ�..."
            'ȡ����Ҫ��������ݷ�������
            Set lcolType = New Collection
            For i = 0 To clstDataType.ListCount - 1
                If clstDataType.Selected(i) Then
                    lcolType.Add clstDataType.List(i), clstDataType.List(i)
                End If
            Next i
            '׼�����루�����ļ�����ʱ�ļ�����
            'mobj����ӿ�.sub����׼�� ctxtDataSource.Text
            
            '��ʼ���롣
            mobj����ӿ�.Sub���ݵ��� lcolRange, lcolType, cprgDatatranform
            
            '����ɹ���
            csbMain.Panels(1) = "����ɹ���"
            MousePointer = 0
            cprgDatatranform.Visible = False
            Cancel = True
    End Select
    
    
    Exit Sub
errHandler:
    sfsub������ "������ӿڲ���", "frmInputDataFromMdb", "mobjGUI_BeforeOperate", Err.Number, Err.Description, False
    cprgDatatranform.Visible = False
    csbMain.Panels(1) = ""
    MousePointer = 0
    Exit Sub
    Resume
End Sub

'���ܣ�  �û�������MDB�ļ�����,"Ԥ��ȡ����" ��ť��Ϊ����
'���ߣ� ���ơ�
Private Sub ctxtDataSource_LostFocus()
    Dim lstrErrDes As String
    
    On Error GoTo errHandler
    
    
    ctbMain.Buttons(1).Enabled = False
    ctbMain.Buttons(2).Enabled = False
    
    If Len(ctxtDataSource.Text) <> 0 Then
        If UCase(Right(ctxtDataSource.Text, 3)) = "MDB" And Dir(ctxtDataSource.Text) <> "" Then
            ccmdPrefech.Enabled = True
        Else
            ccmdPrefech.Enabled = False
            Err.Raise 6666, , "������ļ������Ϸ������������룡"
        End If
    End If
    
    Exit Sub
errHandler:
    sfsub������ "������ӿڲ���", "frmInputDataFromMdb", "ctxtDataSource_LostFocus", Err.Number, Err.Description, False
End Sub

'���ܣ� ͨ��"���ݵ�������"frame���е����ã������ݵ����Ԥ��ǰ��ȡ���ݹ���������
'���ߣ� ���ơ�
Private Function funcCalCon() As Object
    Dim lobjRange As Collection '[���ݷ�Χ�������ݷ�Χֵ] "���ݵ�������"
    Dim lcolItem As Collection
    On Error GoTo errHandler
    
    Set lobjRange = New Collection
    If cchkMedicalDate.Value = 1 Then
        Set lcolItem = New Collection
        lcolItem.Add cdtpBeginDate.Value, "���ݷ�Χֵ"
        lcolItem.Add "��ʼ����", "���ݷ�Χ��"
        lobjRange.Add lcolItem, "��ʼ����"
                
        Set lcolItem = New Collection
        lcolItem.Add cdtpEndDate.Value, "���ݷ�Χֵ"
        lcolItem.Add "��������", "���ݷ�Χ��"
        lobjRange.Add lcolItem, "��������"
    End If
            
    If cchkSystemCode.Value = 1 Then
        Set lcolItem = New Collection
        If Len(ctxtBeginCode.Text) > 0 Then
            lcolItem.Add ctxtBeginCode.Text, "���ݷ�Χֵ"
            lcolItem.Add "��ϵͳ��", "���ݷ�Χ��"
            lobjRange.Add lcolItem, "��ϵͳ���"
        End If
                
        If Len(ctxtEndCode.Text) > 0 Then
            Set lcolItem = New Collection
            lcolItem.Add ctxtEndCode.Text, "���ݷ�Χֵ"
            lcolItem.Add "��ϵͳ��", "���ݷ�Χ��"
            lobjRange.Add lcolItem, "��ϵͳ���"
        End If
    End If
            
    If cchkUnitName.Value = 1 Then
        Set lcolItem = New Collection
        lcolItem.Add ctxtUnit.Text, "���ݷ�Χֵ"
        lcolItem.Add "��λ���Ƽ�", "���ݷ�Χ��"
        lobjRange.Add lcolItem, "��λ���Ƽ�"
    End If
    
    Set funcCalCon = lobjRange
    Exit Function
errHandler:
    sfsub������ "������ӿڲ���", "frmInputDataFromMdb", "funcCalCon", Err.Number, Err.Description, True
End Function












