VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOutputData 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "վ��������ݵ���"
   ClientHeight    =   7800
   ClientLeft      =   1395
   ClientTop       =   1215
   ClientWidth     =   11055
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   11055
   Begin MSComctlLib.Toolbar ctbMain 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   1111
      ButtonWidth     =   900
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "cimg��ťͼ��"
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog ccdgBrowse 
      Left            =   8040
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ԥ������"
      ClipControls    =   0   'False
      ForeColor       =   &H00800000&
      Height          =   3735
      Left            =   120
      TabIndex        =   19
      Top             =   3280
      Width           =   10695
      Begin VSFlex6DAOCtl.vsFlexGrid cgrdPreviewData 
         Height          =   3315
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   5847
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
   Begin VB.CommandButton ccmdBrowse 
      Caption         =   "����ļ�(&B)"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox ctxtOutputDestination 
      Height          =   315
      Left            =   1380
      TabIndex        =   0
      Top             =   840
      Width           =   4335
   End
   Begin VB.Frame cfraSelectInputData 
      Caption         =   "ѡ��Ҫ����������"
      ForeColor       =   &H00800000&
      Height          =   2025
      Left            =   7440
      TabIndex        =   15
      Top             =   1200
      Width           =   3375
      Begin VB.ListBox clstDataType 
         Height          =   1740
         ItemData        =   "frmOutputData.frx":0000
         Left            =   120
         List            =   "frmOutputData.frx":0002
         Style           =   1  'Checkbox
         TabIndex        =   11
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame cfrafiltrateCondition 
      Caption         =   "���ݵ�������"
      ForeColor       =   &H00800000&
      Height          =   2025
      Left            =   120
      TabIndex        =   13
      Top             =   1200
      Width           =   7005
      Begin VB.CheckBox cchkOver 
         Caption         =   "��������"
         Height          =   255
         Left            =   5400
         TabIndex        =   24
         Top             =   1560
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.ComboBox ccmbTemplate 
         Height          =   300
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   1560
         Width           =   2895
      End
      Begin VB.CheckBox cchkTemplate 
         Caption         =   "������"
         Height          =   285
         Left            =   120
         TabIndex        =   22
         Top             =   1560
         Width           =   1035
      End
      Begin VB.TextBox ctxtEndCode 
         Height          =   315
         Left            =   4320
         TabIndex        =   10
         Top             =   1080
         Width           =   2475
      End
      Begin VB.TextBox ctxtBeginCode 
         Height          =   315
         Left            =   1380
         TabIndex        =   9
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox ctxtUnit 
         Height          =   315
         Left            =   1380
         TabIndex        =   6
         Top             =   690
         Width           =   4155
      End
      Begin VB.CheckBox cchkSystemCode 
         Caption         =   "ϵͳ���"
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   1140
         Width           =   1035
      End
      Begin VB.CheckBox cchkUnitName 
         Caption         =   "��λ����"
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox cchkMedicalDate 
         Caption         =   "�������"
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   270
         Value           =   1  'Checked
         Width           =   1155
      End
      Begin VB.CommandButton ccmdLocateUnit 
         Caption         =   "��λ��λ(&L)"
         Height          =   375
         Left            =   5580
         TabIndex        =   7
         Top             =   660
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker cdtpBeginDate 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "yyyy-MM-dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   3
         EndProperty
         Height          =   375
         Left            =   1380
         TabIndex        =   3
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
         Format          =   23658496
         CurrentDate     =   36951
      End
      Begin MSComCtl2.DTPicker cdtpEndDate 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "yyyy-MM-dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   3
         EndProperty
         Height          =   375
         Left            =   3600
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
         Format          =   23658496
         CurrentDate     =   36951
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Height          =   180
         Left            =   3240
         TabIndex        =   21
         Top             =   270
         Width           =   180
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         Height          =   180
         Left            =   4020
         TabIndex        =   14
         Top             =   1200
         Width           =   180
      End
   End
   Begin MSComctlLib.ProgressBar cprgDatatranform 
      Height          =   375
      Left            =   0
      TabIndex        =   17
      Top             =   7080
      Visible         =   0   'False
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.StatusBar csbMain 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   18
      Top             =   7425
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   19447
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList cimg��ťͼ�� 
      Left            =   6960
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "�����������������콡��֤ҵ��ϵͳ�ּҵ�������Ա�ּҺ�Ľ���֤ϵͳ����Ҫ��֤�Ĵ�ҵ��Ա��������"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   7080
      Width           =   10095
   End
   Begin VB.Label clabOutputDestintion 
      AutoSize        =   -1  'True
      Caption         =   "�����ļ���"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   180
      TabIndex        =   16
      Top             =   840
      Width           =   900
   End
End
Attribute VB_Name = "frmOutputData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mobjGUI  As cls����ͨ�ö��� '���ڳ�ʼ����������
Attribute mobjGUI.VB_VarHelpID = -1
Private mobj����ӿ� As ClsManageTransmission '������ӿڶ���

Private mstrϵͳ��Ź̶����� As String

Public pblnInUse As Boolean

Private Sub ctxtBeginCode_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ctxtEndCode.SetFocus
    End If

End Sub

Private Sub Form_Load()
    Dim lcolInfo As New Collection
    Dim lobjRec As Object
    Dim i As Integer
    On Error GoTo errHandler
    pblnInUse = True
    
    '��������ͨ�ö��󣬳�ʼ����������
    Set mobjGUI = New cls����ͨ�ö���
    With lcolInfo
        .Add "Ԥ��(&R)108"
        .Add "����(&E)113"
        .Add "|"
        .Add "�˳�"
    End With
    With mobjGUI
        Set .Form = Me
        Set .c������ = ctbMain
    End With
    mobjGUI.subInitialize lcolInfo, ""
    
    '����������ӿڶ���
    Set mobj����ӿ� = CreateObject("������ӿڲ���.ClsManageTransmission")
    
    '��ȡ�����ڹ���վ�����ļ��м�¼���ϴε������ļ���
    ctxtOutputDestination.Text = mobj����ӿ�.����վ����.�ڲ������ļ�
    
    '��ȡ���п��ܵ�������ݷ��ࡣ
    Set lobjRec = mobj����ӿ�.�������ݷ����嵥
    While Not lobjRec.EOF
        clstDataType.AddItem lobjRec.Fields("���ݷ�����")
        lobjRec.MoveNext
    Wend
    If clstDataType.ListCount > 0 Then
        clstDataType.Selected(0) = True
    End If
    '��ȡ�����������ơ�
    Dim lobj����ģ�弯 As Object
    Set lobj����ģ�弯 = CreateObject("�����󲿼�.clsMedicalExamTemplateSet")
    Set lcolInfo = lobj����ģ�弯.Ԫ�ؼ�
    For i = 1 To lcolInfo.Count
        ccmbTemplate.AddItem lcolInfo(i)
    Next
    Set lobj����ģ�弯 = Nothing
    
    'ȱʡѡ���һ����
    If ccmbTemplate.ListCount > 0 Then
        ccmbTemplate.ListIndex = 0
    End If
    
       
    '������������check��δ��ѡ��ʱ,����򲻿ɡ�.
    ctxtBeginCode.Enabled = False
    ctxtEndCode.Enabled = False
    ctxtUnit.Enabled = False
    ccmbTemplate.Enabled = False
    ccmdLocateUnit.Enabled = False
        
    '�������ڵĳ�ʼֵ��
    cdtpBeginDate.Value = Format(Date, "yyyy-mm-dd")
    cdtpEndDate.Value = Format(Date, "yyyy-mm-dd")
    
    '����,Ԥ����ť��ѡ���ļ���ű�Ϊ����.
    If Len(ctxtOutputDestination.Text) = 0 Then
        ctbMain.Buttons(1).Enabled = False
    End If
    
    '��ʼʱ����ť�����á�
    If ctxtOutputDestination.Text = "" Then
        ctbMain.Buttons(1).Enabled = False
    End If
    
    '��ȡϵͳ��Ź̶����֡�
    Dim lobj��� As Object '�����󣬻�ȡϵͳ��ŵĹ̶����֡�
    Set lobj��� = CreateObject("�����󲿼�.clsMedicalExam")
    mstrϵͳ��Ź̶����� = lobj���.ϵͳ��Ź̶�����
    Set lobj��� = Nothing
    
    ctbMain.Buttons(2).Enabled = False
    Exit Sub
errHandler:
    sfsub������ "������ӿڲ���", "frmOutputData", "form_load", Err.Number, Err.Description, False
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

Private Sub ctxtOutputDestination_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        If cchkMedicalDate.Value = 1 Then
            cdtpBeginDate.SetFocus
        ElseIf cchkSystemCode.Value = 1 Then
            ctxtBeginCode.SetFocus
        ElseIf cchkUnitName.Value = 1 Then
            ctxtUnit.SetFocus
        Else
            clstDataType.SetFocus
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 39 Then
        '���������롰'����
        KeyAscii = 0
    End If

End Sub
Private Sub Form_Activate()
    On Error Resume Next
    ctxtOutputDestination.SetFocus
End Sub

'���ܣ� ����������check��ʱ��������������ؼ���״̬��
'���ߣ� ����
Private Sub cchkMedicalDate_Click()
    On Err GoTo errHandler
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
    sfsub������ "������ӿڲ���", "frmOutputData", "cchkMedicalDate_Click", Err.Number, Err.Description, False
End Sub

'���ܣ� ���ϵͳ���check��ʱ������ϵͳ�������ؼ���״̬��
'���ߣ� ����
Private Sub cchkSystemCode_Click()
    On Err GoTo errHandler
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
    sfsub������ "������ӿڲ���", "frmOutputData", "cchkSystemCode_Click", Err.Number, Err.Description, False
End Sub

'���ܣ� �����λ����check��ʱ�����õ�λ��������ؼ���״̬��
'���ߣ� ����
Private Sub cchkUnitName_Click()
    On Err GoTo errHandler
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
    sfsub������ "������ӿڲ���", "frmOutputData", "cchkUnitName_Click", Err.Number, Err.Description, False
End Sub

'���ܣ� ���������check��ʱ����������������ؼ���״̬��
'���ߣ� ����
Private Sub cchkTemplate_click()
    On Err GoTo errHandler
    If cchkTemplate.Value = 1 Then
        ccmbTemplate.Enabled = True
    Else
        ccmbTemplate.Enabled = False
    End If
    Exit Sub
errHandler:
    sfsub������ "������ӿڲ���", "frmOutputData", "ccmdTemplate_click", Err.Number, Err.Description, False
End Sub

'����: �����ļ����Ҵ���.
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
    
    ctxtOutputDestination.Text = ccdgBrowse.FileName
    
    '�ж�����ĺϷ��ԡ�
    ctxtOutputDestination_LostFocus
    
    Exit Sub
errHandler:
    Exit Sub
End Sub

Private Sub ctxtOutputDestination_LostFocus()
    On Error GoTo errHandler
    
    '�ж������Ŀ���ļ��Ƿ���mdb�ļ���
    ctbMain.Buttons(2).Enabled = False
    ctbMain.Buttons(1).Enabled = False
    If ctxtOutputDestination.Text <> "" Then
        If UCase(Right(Trim(ctxtOutputDestination.Text), 3)) <> "MDB" Then
            sffuncMsg "������Ϸ������ݵ���Ŀ���ļ�������mdb��׺��", sf����
        
            ctxtOutputDestination.Text = ""
        Else
            ctbMain.Buttons(1).Enabled = True
            ctbMain.Buttons(2).Enabled = True
        End If
    End If
    Exit Sub
errHandler:
    sfsub������ "������ӿڲ���", "frmOutputData", "ctxtOutputDestination_Validate", Err.Number, Err.Description, False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    pblnInUse = False
    Set mobjGUI = Nothing
    Set mobj����ӿ� = Nothing
End Sub

'����:  ����,Ԥ�������ļ��е�����.
Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    On Error GoTo errHandler
    Dim lobjRange As Collection
    Dim lobjRec As Object '�������������
    Dim lcolType As Collection
    Dim i As Integer
    
    Select Case Operate
        Case "Ԥ��"
            '��ȡҪԤ�������ݷ�Χ��
            ctxtOutputDestination.SetFocus
            Set lobjRange = funcCalCon
            
            '��ȡ����������ʾ��Ԥ�����ݿ���
            cgrdPreviewData.Rows = 1
            Set lobjRec = mobj����ӿ�.Func�鿴����(lobjRange, 1) '(0����/ 1����)
            gfsubLoadGridFromRec cgrdPreviewData, lobjRec, , "�����������,ϵͳ���,������ݺ���,����,�Ա�,��������,��λ����,�������,��������,������,��Ϻʹ������,���ҽʦ"
            cgrdPreviewData.Rows = cgrdPreviewData.Rows - 1
            
            ctbMain.Buttons(2).Enabled = True
            Cancel = True
        Case "����"
            cprgDatatranform.Value = 0
            cprgDatatranform.Visible = True
            MousePointer = 11
            csbMain.Panels(1) = "����׼�����������Ժ�..."
            
            ctbMain.Enabled = False
            cfrafiltrateCondition.Enabled = False
            cfraSelectInputData.Enabled = False
            DoEvents
            '���б��ȡ����Ҫ���������ݷ�����
            Set lcolType = New Collection
            For i = 0 To clstDataType.ListCount - 1
                If clstDataType.Selected(i) Then
                    lcolType.Add clstDataType.List(i), clstDataType.List(i)
                End If
            Next i
            If lcolType.Count = 0 Then
                Err.Raise 6666, , "��ѡ��Ҫ���������ݷ��࣡"
            End If
            
            '��ȡҪԤ�������ݷ�Χ��
            Set lobjRange = funcCalCon
            
            '����ǰ���������ļ���
            csbMain.Panels(1) = "���ڿ����ļ������Ժ�..."
            mobj����ӿ�.sub����׼�� ctxtOutputDestination.Text
            DoEvents
            
            '��ʼ����������ʾ���ȡ�
            csbMain.Panels(1) = "���ڵ��������Ժ�..."
            mobj����ӿ�.Sub���ݵ��� lobjRange, lcolType, cprgDatatranform
            DoEvents
            
            cprgDatatranform.Visible = False
            ctbMain.Enabled = True
            cfrafiltrateCondition.Enabled = True
            cfraSelectInputData.Enabled = True
            csbMain.Panels(1) = "�����ɹ���"
            MousePointer = 0
            Cancel = True
    End Select
    
    Exit Sub
errHandler:
    sfsub������ "������ӿڲ���", "frmOutputData", "mobjGUI_BeforeOperate", Err.Number, Err.Description, False
    If Operate = "����" Then
        ctbMain.Enabled = True
        cfrafiltrateCondition.Enabled = True
        cfraSelectInputData.Enabled = True
        csbMain.Panels(1) = "����ʧ�ܡ�"
    End If
    cprgDatatranform.Visible = False
    MousePointer = 0
End Sub

'���ܣ�����ͨ����λ��λ�õ��ĵ�λ���ƣ���ʾ�ڵ�λ�����ı����У�����λ����֮����Ӣ�Ķ��ŷָ���
Private Sub ccmdLocateUnit_Click()
    On Error GoTo errHandler
    Dim lobj������ As Object
    Dim lobjRec As Object
    Dim lstrUnit As String
    
    '����������ҵ�����
    Set lobj������ = CreateObject("�����󲿼�.clsManageMedicalExam")
    
    '��λ��λ��
    Set lobjRec = lobj������.func��λ��λ
    
    '�Ѷ�λ���ĵ�λ������ʾ�ڵ�λ����¼����С�
    If Not lobjRec Is Nothing Then
        If lobjRec.RecordCount > 0 Then
            lstrUnit = lobjRec.Fields("��λ����").Value
        End If
    End If
    If Trim(lstrUnit) <> "" Then
        If Trim(ctxtUnit.Text) <> "" Then
            ctxtUnit.Text = Trim(ctxtUnit.Text) & "," & lstrUnit
        Else
            ctxtUnit.Text = lstrUnit
        End If
    End If
    
    Exit Sub
errHandler:
    sfsub������ "������ӿڲ���", "frmInputDataFromMdb", "ccmdLocateUnit_Click", Err.Number, Err.Description, False
End Sub

'���ܣ����ҵ�λ���Ƽ��ִ��е����Ķ��ţ�����Ӣ�Ķ����滻����
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

'���ܣ��ȴ�"���ݵ�������"frame��ָ���ķ�Χֵ���㵼�����ݹ�������
Private Function funcCalCon() As Object
    Dim lobjRange As Collection '[���ݷ�Χ�������ݷ�Χֵ] "���ݵ�������"
    Dim lcolItem As Collection
    On Error GoTo errHandler
    
    Set lobjRange = New Collection
    If cchkMedicalDate.Value = 1 Then
        Set lcolItem = New Collection
        lcolItem.Add Format(cdtpBeginDate.Value, "yyyy-mm-dd"), "���ݷ�Χֵ"
        lcolItem.Add "��ʼ����", "���ݷ�Χ��"
        lobjRange.Add lcolItem, "��ʼ����"
                
        Set lcolItem = New Collection
        lcolItem.Add Format(cdtpEndDate.Value, "yyyy-mm-dd"), "���ݷ�Χֵ"
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
            
    If cchkUnitName.Value = 1 And Len(ctxtUnit.Text) > 0 Then
        Set lcolItem = New Collection
        lcolItem.Add ctxtUnit.Text, "���ݷ�Χֵ"
        lcolItem.Add "��λ���Ƽ�", "���ݷ�Χ��"
        lobjRange.Add lcolItem, "��λ���Ƽ�"
    End If
    
    If cchkTemplate.Value = 1 Then
        Set lcolItem = New Collection
        lcolItem.Add ccmbTemplate.List(ccmbTemplate.ListIndex), "���ݷ�Χֵ"
        lcolItem.Add "������", "���ݷ�Χ��"
        lobjRange.Add lcolItem, "������"
    End If
    
    If cchkOver.Value = 1 Then
        Set lcolItem = New Collection
        lcolItem.Add True, "���ݷ�Χֵ"
        lcolItem.Add "��������", "���ݷ�Χ��"
        lobjRange.Add lcolItem, "��������"
    End If
    Set funcCalCon = lobjRange
    Exit Function

errHandler:
    sfsub������ "������ӿڲ���", "frmInputDataFromMdb", "funcCalCon", Err.Number, Err.Description, True
End Function

