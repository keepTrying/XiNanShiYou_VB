VERSION 5.00
Object = "{8099FCC2-0A81-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F48D4DEC-1198-11D5-91BE-0050BA06B70C}#5.9#0"; "¼��ؼ�.ocx"
Begin VB.Form frm���ô��� 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "���ô���"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   9735
   ClipControls    =   0   'False
   Icon            =   "frm���ô���.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   9735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   6135
      Left            =   0
      TabIndex        =   7
      Top             =   840
      Width           =   9615
      Begin VB.Frame Frame8 
         Caption         =   "���ۿ�����"
         Height          =   1755
         Left            =   4680
         TabIndex        =   9
         Top             =   4080
         Width           =   4680
         Begin VB.OptionButton copt���� 
            Caption         =   "���Դ��ۣ����ϸ����"
            Height          =   180
            Index           =   2
            Left            =   150
            TabIndex        =   5
            Top             =   1185
            Width           =   2100
         End
         Begin VB.OptionButton copt���� 
            Caption         =   "���Դ��ۣ������ϸ����"
            Height          =   180
            Index           =   1
            Left            =   150
            TabIndex        =   4
            Top             =   735
            Width           =   2280
         End
         Begin VB.OptionButton copt���� 
            Caption         =   "������"
            Height          =   180
            Index           =   0
            Left            =   150
            TabIndex        =   3
            Top             =   300
            Value           =   -1  'True
            Width           =   840
         End
      End
      Begin VB.CommandButton ccmd��λ��λ 
         Caption         =   "��λ��λ"
         Height          =   375
         Left            =   8520
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox ctxt���� 
         Height          =   360
         Left            =   6000
         TabIndex        =   1
         Text            =   "1.00"
         Top             =   960
         Width           =   1065
      End
      Begin VB.VScrollBar cvsl���� 
         Height          =   360
         Left            =   6600
         Max             =   1
         Min             =   100
         TabIndex        =   2
         Top             =   960
         Value           =   100
         Width           =   675
      End
      Begin ¼��ؼ�.ctlInputBox cinp���ѵ�λ 
         Height          =   360
         Left            =   4680
         TabIndex        =   6
         Top             =   360
         Width           =   3645
         _ExtentX        =   6429
         _ExtentY        =   635
         BackColor       =   16777215
         ForeColor       =   0
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
         LeftOfTextbox   =   760
         Text            =   ""
         Label           =   "���ѵ�λ"
         Enabled         =   0   'False
         ����            =   ""
         ����            =   0
         ����������ֵ  =   0   'False
         ���������Сֵ  =   0   'False
         �����ѡ        =   0   'False
      End
      Begin VSFlex6Ctl.vsFlexGrid cFlg���� 
         Height          =   5685
         Left            =   120
         TabIndex        =   8
         Top             =   300
         Width           =   4395
         _cx             =   20192840
         _cy             =   20195116
         _ConvInfo       =   1
         Appearance      =   0
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
         BackColorBkg    =   16777215
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
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
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
         DataMember      =   ""
      End
      Begin ¼��ؼ�.ctlInputBox cinb��λ��� 
         Height          =   360
         Left            =   5160
         TabIndex        =   10
         Top             =   360
         Visible         =   0   'False
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   635
         BackColor       =   16777215
         ForeColor       =   0
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
         LeftOfTextbox   =   760
         Text            =   ""
         Label           =   "��λ���"
         Enabled         =   0   'False
         ����            =   ""
         ����            =   0
         ����������ֵ  =   0   'False
         ���������Сֵ  =   0   'False
         �����ѡ        =   0   'False
         BackgroundColor =   15791081
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���۱���"
         Height          =   180
         Left            =   4680
         TabIndex        =   12
         Top             =   1080
         Width           =   720
      End
      Begin VB.Label clbl����˵�� 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1785
         Left            =   4560
         TabIndex        =   11
         Top             =   1800
         Width           =   4965
      End
   End
   Begin MSComctlLib.ImageList cimg��ťͼ�� 
      Left            =   5760
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar ctlb���� 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "cimg��ťͼ��"
      _Version        =   393216
   End
End
Attribute VB_Name = "frm���ô���"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public pblnInUse As Boolean

Private WithEvents mobjGUI As cls����ͨ�ö���
Attribute mobjGUI.VB_VarHelpID = -1

Private Sub ccmd��λ��λ_Click()
    On Error GoTo errhandler
    
    Dim lobj��λ��Ϣ As Object
    Set lobj��λ��Ϣ = pobj��λ��λ.func��λ�򵥶�λ(8600, 1000)
    If Not (lobj��λ��Ϣ Is Nothing) Then   '�ѻ�ȡ��λ��Ϣ
        cinp���ѵ�λ.Text = lobj��λ��Ϣ.Fields!��λ����
        cinb��λ���.Text = IIf(IsNull(lobj��λ��Ϣ.Fields!������), "", lobj��λ��Ϣ.Fields!������)
    Else
    
    End If

    ctxt����.SetFocus
    Exit Sub
errhandler:
    sfsub������ "�շѽ��沿��", "frm���ô���", "ccmd��λ��λ_Click", Err.Number, Err.Description, False
End Sub

Private Sub cFlg����_Click()
    On Error Resume Next
    cinb��λ���.Text = cFlg����.TextMatrix(cFlg����.Row, 1)
    cinp���ѵ�λ.Text = cFlg����.TextMatrix(cFlg����.Row, 2)
    ctxt����.Text = cFlg����.TextMatrix(cFlg����.Row, 3)
    cvsl����.Value = Int(Val(cFlg����.TextMatrix(cFlg����.Row, 3)) * 100)
End Sub

Private Sub cvsl����_Change()
    
    ctxt����.Text = cvsl����.Value / 100
End Sub

Private Sub Form_Load()
    On Error GoTo errhandler
    
    If pblnInUse = True Then Exit Sub
    
    pblnInUse = True

    '��ʼ��������
    Dim lcol��������ť As Collection
    Set lcol��������ť = New Collection
    Set mobjGUI = New cls����ͨ�ö���
    Set mobjGUI.Form = Me
    Set mobjGUI.c������ = ctlb����

    lcol��������ť.Add "���"
    lcol��������ť.Add "ɾ��"
    lcol��������ť.Add "|"
    lcol��������ť.Add "����"
    lcol��������ť.Add "|"
    lcol��������ť.Add "�˳�"
    
    mobjGUI.subInitialize lcol��������ť, ""
    
    Dim lint���ۿ���  As Integer
    lint���ۿ��� = Val(pobj�շѹ���.ҵ������("���ۿ���"))
    copt����(lint���ۿ���).Value = True
    
    Dim lobjRec As Object
    Set lobjRec = dafuncGetData("select a.��λ���,b.��λ����,a.���۱��� from �շѹ���_������Ϣ�� a inner join ��λ����_��λ������Ϣ�� b on a.��λ���=b.������")
    Set cFlg����.DataSource = lobjRec
    
    cFlg����.Row = 0
    
    Exit Sub
errhandler:
    sfsub������ "�շѽ��沿��", "frm���ô���", "Form_Load", Err.Number, Err.Description, False
    Exit Sub
    Resume
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    pblnInUse = False
End Sub

Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    On Error GoTo errhandler
    Select Case Operate
    Case "���"
        Cancel = True
        ctxt����.Enabled = True
        ccmd��λ��λ.Enabled = True
        cvsl����.Enabled = True
        
        cinb��λ���.Text = ""
        cinp���ѵ�λ.Text = ""
        ctxt����.Text = "1.00"
        
   Case "ɾ��"
        Cancel = True
        If cinb��λ���.Text <> "" Then
            pobj�շѹ���.subɾ��������Ϣ LTrim(RTrim(cinb��λ���.Text))
            cFlg����.RemoveItem cFlg����.RowSel
        Else
            MsgBox "��ѡ�����ĵ�λ��", vbExclamation, "ϵͳ����"
        End If
   Case "����"
        Dim lint���ۿ��� As Integer
        Cancel = True
        If copt����(0).Value Then
            lint���ۿ��� = 1
        ElseIf copt����(1).Value Then
            lint���ۿ��� = 1
        Else
            lint���ۿ��� = 1
        End If
        pobj�շѹ���.ҵ������("���ۿ���") = lint���ۿ���
        
        If cinb��λ���.Text <> "" Then
            pobj�շѹ���.sub���������Ϣ cinb��λ���.Text, IIf(IsNull(ctxt����.Text), "1", ctxt����.Text)
            'ˢ������
            Dim lobjRec As Object
            Set lobjRec = dafuncGetData("select a.��λ���,b.��λ����,a.���۱��� from �շѹ���_������Ϣ�� a inner join ��λ����_��λ������Ϣ�� b on a.��λ���=b.������")
            Set cFlg����.DataSource = lobjRec
            
            cinb��λ���.Text = ""
            cinp���ѵ�λ.Text = ""
            ccmd��λ��λ.SetFocus
            
        End If
        
   End Select
   
    Exit Sub
errhandler:
    sfsub������ "�շѽ��沿��", "frm���ô���", "mobjGUI_BeforeOperate", Err.Number, Err.Description, False
    Exit Sub
    Resume
   
End Sub
