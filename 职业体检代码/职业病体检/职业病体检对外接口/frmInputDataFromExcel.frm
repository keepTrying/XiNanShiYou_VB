VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmInputDataFromExcel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�ⵥλ���ݱ���"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10470
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   10470
   StartUpPosition =   3  '����ȱʡ
   Begin MSComDlg.CommonDialog ccdgBrowse 
      Left            =   9600
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar cstbMain 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   8
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar ctbMain 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   7
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
   Begin MSComctlLib.ProgressBar cprgDataTransform 
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   7440
      Visible         =   0   'False
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton ccmdBrowse 
      Caption         =   "����ļ�(&B)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   4
      Top             =   840
      Width           =   1275
   End
   Begin VB.Frame cfraPreview 
      Caption         =   "Ԥ������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   4155
      Left            =   120
      TabIndex        =   3
      Top             =   2880
      Width           =   10335
      Begin VSFlex6DAOCtl.vsFlexGrid cgrdPreview 
         Height          =   3855
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   6800
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
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   "��λ����    |����   |�Ա� |���� "
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
   Begin VB.Frame cfraSelectMedicalTemplate 
      Caption         =   "ѡ������"
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1395
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   4935
      Begin VB.ListBox clstTemplate 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   4620
      End
   End
   Begin VB.TextBox ctxtDataSource 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   840
      Width           =   4035
   End
   Begin MSComctlLib.ImageList cimg��ťͼ�� 
      Left            =   8520
      Top             =   660
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Label clabDataSource 
      AutoSize        =   -1  'True
      Caption         =   "������Դ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   900
   End
End
Attribute VB_Name = "frmInputDataFromExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private mobj����ӿ� As ClsManageTransmission
Private WithEvents mobjGUI  As cls����ͨ�ö���
Attribute mobjGUI.VB_VarHelpID = -1
Public pblnInUse As Boolean

'����: ����û��޸ĵ������Ƿ�Ϸ�.
Private Sub cgrdPreview_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    On Error Resume Next
    If Not IsNumeric(cgrdPreview.TextMatrix(Row, Col)) And Col = 3 Then
        sffuncMsg "�����ֶ�ֻ��Ϊ����", sf����
    End If
    
    If cgrdPreview.TextMatrix(Row, Col) <> "��" And cgrdPreview.TextMatrix(Row, Col) <> "Ů" _
        And Len(cgrdPreview.TextMatrix(Row, Col)) <> 0 And Col = 2 Then
        sffuncMsg "�Ա��ֶ���������", sf����
    End If

End Sub

Private Sub ctxtDataSource_LostFocus()
    On Err GoTo errHandler
    If Len(ctxtDataSource.Text) <> 0 And Right(ctxtDataSource.Text, 3) = "xls" Then
        ctbMain.Buttons(1).Enabled = True
    End If
    Exit Sub
errHandler:
    sfsub������ "������ӿڲ���", "frmInputFromExcel", "ctxtDataSource_LostFocus", Err.Number, Err.Description, False
End Sub

'��ʼ������,
Private Sub Form_Load()
    Dim lcol��������ť As New Collection
    Dim lcolTemplateSet As Object
    Dim lcolInfo As Collection
    
    Dim i As Integer
    
    On Error GoTo errHandler
    pblnInUse = True
    
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
    
    Set mobj����ӿ� = CreateObject("������ӿڲ���.ClsManageTransmission")
    
    '��ʼ��"ctxtDataSource"�ļ���.(excel�ļ�·��)
    ctxtDataSource.Text = mobj����ӿ�.����վ����.Excel�ļ�
    
    '��ʼ��ѡ�������б��
    Set lcolTemplateSet = CreateObject("�����󲿼�.ClsMedicalExamTemplateSet")
    lcolTemplateSet.�������� = 3        '����Ϊ���������
    Set lcolInfo = lcolTemplateSet.Ԫ�ؼ�
    For i = 1 To lcolInfo.Count
        clstTemplate.AddItem lcolInfo(i)
    Next i
    
    '����,Ԥ����ť��ѡ���ļ���ű�Ϊ����.
    If Len(ctxtDataSource.Text) = 0 Then
        ctbMain.Buttons(1).Enabled = False
    End If
    ctbMain.Buttons(2).Enabled = False
    If clstTemplate.ListCount > 0 Then
        clstTemplate.ListIndex = 0
    End If
    cprgDataTransform.ZOrder 0
    Exit Sub
errHandler:
    sfsub������ "������ӿڲ���", "frmInputFromExcel", "Form_Load", Err.Number, Err.Description, False
End Sub

'����: �����ļ����Ҵ���.
Private Sub ccmdBrowse_Click() ' ���á�CancelError��Ϊ True
    ccdgBrowse.CancelError = True
    On Error GoTo errHandler
    ' ���ñ�־
    ccdgBrowse.Flags = cdlOFNHideReadOnly
    ' ���ù�����
    ccdgBrowse.Filter = "All Files (*.*)|*.*|Excel file" & _
        "(*.xls)|*.xls|Batch Files (*.bat)|*.bat"
    ccdgBrowse.FilterIndex = 2
    ccdgBrowse.ShowOpen
    ctxtDataSource.Text = ccdgBrowse.FileName
    If Len(ctxtDataSource.Text) <> 0 Then
        ctbMain.Buttons(1).Enabled = True
    End If
    
    Exit Sub
errHandler:
    ' �û����ˡ�ȡ������ť
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set mobjGUI = Nothing
    Set mobj����ӿ� = Nothing
    pblnInUse = False
End Sub

'����:  ����,Ԥ�������ļ��е�����.
Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    Dim lcolItem As Collection  '����һ����Ҫ�����������
    Dim lcolInfo As Collection  '���������Ҫ���������
    Dim i As Integer
    Dim j As Integer            '��¼�ӱ����ļ��ж����Ŀ��е�����
    On Error GoTo errHandler
    MousePointer = 11
    
    Select Case Operate
        Case "Ԥ��"
            '�ж��û������.xls�ļ��Ƿ���ڡ�
            If Dir(ctxtDataSource.Text) = "" Then
                Err.Raise 6666, , "�����EXCEL �ļ����ڣ����������룡"
            End If
            cstbMain.Panels(1) = "����Ԥ���ⵥλ�����ļ������ݣ����Ժ�..."
            MousePointer = 11
            'Ԥ���ļ����ݡ�
            Set lcolItem = mobj����ӿ�.Func�������ļ��������Ա��Ϣ(ctxtDataSource.Text)
            With cgrdPreview
                .Rows = lcolItem.Count + 1
                .Cols = 4
                .TextMatrix(0, 0) = "��λ����"
                .TextMatrix(0, 1) = "����"
                .TextMatrix(0, 2) = "�Ա�"
                .TextMatrix(0, 3) = "����"
            End With
            j = 1
            For i = 1 To lcolItem.Count
                If lcolItem(i)("����") = "" Then
                    '��������ļ����п��У����������С�
                Else
                    cgrdPreview.TextMatrix(j, 0) = lcolItem(i)("��λ����")
                    cgrdPreview.TextMatrix(j, 1) = lcolItem(i)("����")
                    cgrdPreview.TextMatrix(j, 2) = lcolItem(i)("�Ա�")
                    cgrdPreview.TextMatrix(j, 3) = lcolItem(i)("����")
                End If
                j = j + 1
            Next
            cgrdPreview.Rows = j
            
            ctbMain.Buttons(2).Enabled = True
            cgrdPreview.Editable = True
            cstbMain.Panels(1) = ""
            MousePointer = 0
            Cancel = True
        Case "����"
            '�ж��û������.xls�ļ��Ƿ���ڡ�
            If Dir(ctxtDataSource.Text) = "" Then
                Err.Raise 6666, , "�����EXCEL �ļ����ڣ����������룡"
            End If
            If cgrdPreview.Rows = 1 Then
                Err.Raise 6666, , "�����EXCEL �ļ��������ݿ��Ե��룡"
            End If
            cstbMain.Panels(1) = "���ڵ��������������Ա���������Ժ�..."
            cprgDataTransform.Visible = True
            '����Ԥ�����ݿ��ڵ�����.
            Set lcolInfo = New Collection
            For i = 1 To cgrdPreview.Rows - 1
                Set lcolItem = New Collection
                lcolItem.Add cgrdPreview.TextMatrix(i, 0), "��λ����"
                lcolItem.Add cgrdPreview.TextMatrix(i, 1), "����"
                lcolItem.Add cgrdPreview.TextMatrix(i, 2), "�Ա�"
                lcolItem.Add cgrdPreview.TextMatrix(i, 3), "����"
                
                lcolInfo.Add lcolItem, CStr(i)
            Next i
            mobj����ӿ�.Sub���������Ա�Ǽ� clstTemplate.List(clstTemplate.ListIndex), cprgDataTransform, lcolInfo
            
            On Error Resume Next
            mobj����ӿ�.sub�����ļ��� ctxtDataSource.Text
            
            cstbMain.Panels(1) = "����ɹ���"
            cgrdPreview.Rows = 1
            ctbMain.Buttons(1).Enabled = False
            ctbMain.Buttons(2).Enabled = False
            
            cprgDataTransform.Visible = False
            MousePointer = 0
            Cancel = True
    End Select
    
    
    Exit Sub
errHandler:
    sfsub������ "������ӿڲ���", "frmInputFromExcel", "mobjGUI_BeforeOperate", Err.Number, Err.Description, False
    cprgDataTransform.Visible = False
    cstbMain.Panels(1) = ""
    MousePointer = 0
End Sub


