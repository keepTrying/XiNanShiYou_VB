VERSION 5.00
Object = "{8099FCC2-0A81-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frm���� 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "����"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   11070
   ClipControls    =   0   'False
   Icon            =   "frm����.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   11070
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      Caption         =   "�������޸ĵ� "
      ForeColor       =   &H80000008&
      Height          =   6120
      Left            =   120
      TabIndex        =   18
      Top             =   1800
      Width           =   10860
      Begin VB.Frame Frame1 
         Height          =   5500
         Left            =   6360
         TabIndex        =   19
         Top             =   120
         Width           =   4335
         Begin VB.OptionButton coptFind 
            Caption         =   "�����Ƿ�����"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   5040
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.TextBox ctxtFind 
            Height          =   375
            Left            =   1680
            TabIndex        =   24
            Top             =   5040
            Width           =   1935
         End
         Begin VB.ComboBox ccmb�շѱ�׼ 
            Height          =   300
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   240
            Width           =   2655
         End
         Begin VB.ComboBox Ccbo�շ���Ŀ���� 
            Height          =   300
            Left            =   1440
            TabIndex        =   6
            Top             =   600
            Width           =   2655
         End
         Begin VSFlex6Ctl.vsFlexGrid cgrdItem 
            Height          =   3975
            Left            =   120
            TabIndex        =   26
            Top             =   960
            Width           =   4000
            _cx             =   60955536
            _cy             =   60955491
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
            BackColorAlternate=   15791081
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   3
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   "�շ���Ŀ            |���         |���Ƿ�     "
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
         Begin VB.Label Clab�շ���Ŀ���� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�շ���Ŀ����"
            Height          =   180
            Left            =   120
            TabIndex        =   22
            Top             =   720
            Width           =   1200
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�շѱ�׼"
            Height          =   180
            Index           =   6
            Left            =   120
            TabIndex        =   21
            Top             =   360
            Width           =   720
         End
      End
      Begin VSFlex6Ctl.vsFlexGrid cgrdDetail 
         Height          =   5340
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   6165
         _cx             =   115354234
         _cy             =   115352779
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
         BackColorAlternate=   16777215
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
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
         TabBehavior     =   1
         OwnerDraw       =   0
         Editable        =   -1  'True
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
      Begin VB.Label clblTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ܽ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         TabIndex        =   27
         Top             =   5640
         Width           =   1200
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����Del��������ɾ����ǰѡ�е���Ŀ"
         Height          =   180
         Left            =   120
         TabIndex        =   23
         Top             =   3960
         Width           =   2970
      End
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      Caption         =   "������Ϣ"
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   10845
      Begin VB.TextBox ctxtInput 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   960
         TabIndex        =   9
         Top             =   200
         Width           =   1665
      End
      Begin VB.TextBox ctxtInput 
         Height          =   300
         Index           =   2
         Left            =   3600
         TabIndex        =   0
         Top             =   200
         Width           =   1560
      End
      Begin VB.TextBox ctxtInput 
         Height          =   300
         Index           =   3
         Left            =   6360
         TabIndex        =   1
         Top             =   200
         Width           =   2715
      End
      Begin VB.ComboBox ccmb���ܿ��� 
         Height          =   300
         Left            =   6360
         TabIndex        =   5
         Text            =   "Combo1"
         Top             =   600
         Width           =   1695
      End
      Begin VB.ComboBox ccmb�������� 
         Height          =   300
         Left            =   960
         TabIndex        =   3
         Top             =   600
         Width           =   1695
      End
      Begin VB.ComboBox ccmbƬ�� 
         Height          =   300
         Left            =   3600
         TabIndex        =   4
         Top             =   600
         Width           =   1695
      End
      Begin VB.CommandButton ccmd��λ 
         Caption         =   "..."
         Height          =   255
         Left            =   9240
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�շѱ��"
         Height          =   180
         Index           =   1
         Left            =   105
         TabIndex        =   15
         Top             =   225
         Width           =   720
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Index           =   2
         Left            =   3000
         TabIndex        =   14
         Top             =   225
         Width           =   540
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ѵ�λ"
         Height          =   180
         Index           =   3
         Left            =   5520
         TabIndex        =   13
         Top             =   225
         Width           =   720
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ܿ���"
         Height          =   180
         Index           =   4
         Left            =   5520
         TabIndex        =   12
         Top             =   705
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   180
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ƭ��"
         Height          =   180
         Left            =   3000
         TabIndex        =   10
         Top             =   720
         Width           =   360
      End
   End
   Begin MSComctlLib.ImageList cimg��ťͼ�� 
      Left            =   9840
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar ctlb������ 
      Align           =   1  'Align Top
      Height          =   555
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   11070
      _ExtentX        =   19526
      _ExtentY        =   979
      ButtonWidth     =   1455
      ButtonHeight    =   926
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "cimg��ťͼ��"
      _Version        =   393216
      Begin VB.CheckBox cchk��� 
         Caption         =   "��������"
         Height          =   255
         Left            =   8520
         TabIndex        =   17
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frm����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public pblnInUse As Boolean

'����������
Public pstr�շѱ�� As String
Public pstr��λ��� As String
Public pstrҵ����� As String

Dim WithEvents mobj����ͨ�ö��� As cls����ͨ�ö���
Attribute mobj����ͨ�ö���.VB_VarHelpID = -1


Private Const �շ�_�շѱ�� = 1
Private Const �շ�_������ = 2
Private Const �շ�_���ѵ�λ = 3

Private Const �����嵥_�շ���Ŀ��� = 0
Private Const �����嵥_�շ���Ŀ���� = 1
Private Const �����嵥_���� = 2
Private Const �����嵥_���� = 3
Private Const �����嵥_��� = 4


Dim mstrUndoCount As String          '���ڱ�������ԭ�����ַ���,�Ա������벻�Ϸ�ʱ�ܹ���ԭ
Dim mstrUndoMoney As String          '���ڱ�������ԭ�����ַ���,�Ա������벻�Ϸ�ʱ�ܹ���ԭ
Dim mcur��С���� As Currency
Dim mcur��󵥼� As Currency

Dim mstr���ѵ�λ��� As String  '�ӵ�λ��λ�ӿڵõ��Ľ��ѵ�λ�ı��
Dim mint���ѷ�ʽ��� As Integer '���ѷ�ʽ�ı��

Dim mint��Ŀ���� As Integer

Private Sub Ccbo�շ���Ŀ����_Click()
    On Error GoTo errhandler
   
    Dim lobjRec As Object            '���������¼���ݼ�
    
    '�����շ���Ŀ��������,��ȡ�շѱ��ǰ׺
    Set lobjRec = dafuncGetData("select �շ���Ŀ��� from �շѹ���_�շ���Ŀ�ֵ�� where �շ���Ŀ����= '" & Ccbo�շ���Ŀ����.Text & "'")
    
    '��ȡ�¼��շ���Ŀ
    Set lobjRec = dafuncGetData("select �շ���Ŀ���� as �շ���Ŀ,�շ���Ŀ��� as ���,���Ƿ� from �շѹ���_�շ���Ŀ�ֵ�� where left(�շ���Ŀ���,3)='" & Left$(lobjRec("�շ���Ŀ���"), 3) & "' and len(�շ���Ŀ���)>3")
    
    cgrdItem.FormatString = ""
    Set cgrdItem.DataSource = lobjRec
    cgrdItem.ColWidth(0) = 2000
    cgrdItem.Row = 0
    On Error Resume Next
    ctxtFind.SetFocus
    ctxtFind.Text = ""
    Exit Sub
errhandler:
    MsgBox "��ȡ����ʾָ��������շ���Ŀʧ�ܣ�" & Error, vbOKOnly + vbExclamation, "ϵͳ��ʾ"
End Sub

Private Sub ccmb�շѱ�׼_Click()
    Dim lrds�շѱ�׼ As Object
    Dim i As Integer
    Dim lcurMoney As Currency
    
    On Error GoTo errhandler
    
    Set lrds�շѱ�׼ = dafuncGetData("select a.�շ���Ŀ���,b.�շ���Ŀ����,a.����,a.����,b.������λ,���=a.����*a.���� from �շѹ���_�շѱ�׼��Ϣ�� a,�շѹ���_�շ���Ŀ�ֵ�� b where b.�շ���Ŀ���=a.�շ���Ŀ��� and �շѱ�׼����='" & ccmb�շѱ�׼.Text & "'")
    
    If lrds�շѱ�׼.EOF Then
        sffuncMsg "�շѱ�׼�����շ���Ŀ��", sf����
        Exit Sub
    Else
        lrds�շѱ�׼.MoveFirst
        Dim llngItemCount As Long
        For i = 0 To lrds�շѱ�׼.RecordCount - 1
            If Not func�����Ŀ�Ƿ���ѡ(lrds�շѱ�׼("�շ���Ŀ���")) Then
                sub�����Ŀ lrds�շѱ�׼("�շ���Ŀ���")
                llngItemCount = llngItemCount + 1
            End If
            lrds�շѱ�׼.MoveNext
        Next
        
        sub�����ܽ��
        
        If llngItemCount = lrds�շѱ�׼.RecordCount Then
            MsgBox "�շѱ�׼�е������շ���Ŀ(" & llngItemCount & "��)����ӵ������嵥�У�" & vbCrLf & vbCrLf & "(���ι�������� " & lrds�շѱ�׼.RecordCount & " ���е� " & llngItemCount & " ���շ���Ŀ��)", vbInformation, "ϵͳ��ʾ"
        ElseIf llngItemCount = 0 Then
            MsgBox "�շѱ�׼�е������շ���Ŀ�ڷ����嵥������ӣ�" & vbCrLf & vbCrLf & "(���ι�������� " & lrds�շѱ�׼.RecordCount & " ���е� " & llngItemCount & " ���շ���Ŀ��)", vbInformation, "ϵͳ��ʾ"
        Else
            MsgBox "�շѱ�׼�в����շ���Ŀ�ڷ����嵥�������,����� " & llngItemCount & " ������ӵ������嵥��" & vbCrLf & vbCrLf & "(���ι�������� " & lrds�շѱ�׼.RecordCount & " ���е� " & llngItemCount & " ���շ���Ŀ��)", vbInformation, "ϵͳ��ʾ"
        End If
    End If
                
    Exit Sub
errhandler:
    sfsub������ "�շѽ��沿��", "frm����", "ccmb�շѱ�׼_Click", Err.Number, Err.Description, False
End Sub

Private Sub ccmd��λ_Click()
    Dim lrds������Ϣ As Object               '��λ�Ĵ�����Ϣ
    Dim lrdsTemp As Object
    
    On Error GoTo errhandler
    
    '���õ�λ�����Ķ�λ�ӿڻ�ȡ��λ��Ϣ
    Set lrdsTemp = pobj��λ��λ.func��λ�򵥶�λ(100, 100)
    If Not (lrdsTemp Is Nothing) Then
        If lrdsTemp.RecordCount > 0 Then
            '��ʾ��λ����`
            ctxtInput(�շ�_���ѵ�λ).Text = lrdsTemp("��λ����")
            '��ʾ�������ࡢƬ��
            ccmb��������.Text = lrdsTemp("��������")
            ccmbƬ��.Text = IIf(IsNull(lrdsTemp("Ƭ��")), "", lrdsTemp("Ƭ��"))
            
            '���浥λ��������
            mstr���ѵ�λ��� = lrdsTemp("������")
            ctxtInput(�շ�_���ѵ�λ).SetFocus
        End If
    End If


    Exit Sub
errhandler:
    sfsub������ "�շѽ��沿��", "frm����", "ccmd��λ_Click", Err.Number, Err.Description, False
End Sub



Private Sub cgrdDetail_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    Dim lcurMoney As Currency
    
    On Error GoTo errhandle
    ctlb������.Buttons(1).Enabled = True
    Select Case cgrdDetail.TextMatrix(0, Col)
        Case "����"
            '�ж�������Ƿ���ֵ
            If Len(cgrdDetail.TextMatrix(Row, Col)) > 4 Then
                cgrdDetail.TextMatrix(Row, Col) = mstrUndoCount
            Else
                If IsNumeric(cgrdDetail.TextMatrix(Row, Col)) And Val(cgrdDetail.TextMatrix(Row, Col)) > 0 Then
                    '����ֵ
                    '������
                    cgrdDetail.TextMatrix(Row, �����嵥_���) = cgrdDetail.TextMatrix(Row, �����嵥_����) * cgrdDetail.TextMatrix(Row, �����嵥_����)
                Else
                    '������ֵ
                    'Undo
                    cgrdDetail.TextMatrix(Row, Col) = mstrUndoCount
                End If
            End If
        Case "����"
            Dim lcur���� As Currency
            If mcur��С���� = mcur��󵥼� Then
                sffuncMsg "���շ���Ŀ�����Ѷ�,�����޸ģ�", sf����
                cgrdDetail.TextMatrix(Row, Col) = mstrUndoMoney
                Exit Sub
            End If
            
            If IsNumeric(cgrdDetail.TextMatrix(Row, Col)) Then
                If Val(cgrdDetail.TextMatrix(Row, Col)) > 0 Then
                    If Val(cgrdDetail.TextMatrix(Row, Col)) <= mcur��󵥼� And Val(cgrdDetail.TextMatrix(Row, Col)) >= mcur��С���� Then
                        cgrdDetail.TextMatrix(Row, �����嵥_���) = cgrdDetail.TextMatrix(Row, �����嵥_����) * cgrdDetail.TextMatrix(Row, �����嵥_����)
                    Else
                        sffuncMsg "����ĵ��۳�����Χ��", sf����
                        cgrdDetail.TextMatrix(Row, Col) = mstrUndoMoney
                    End If
                Else
                    cgrdDetail.TextMatrix(Row, Col) = mstrUndoMoney
                End If
            Else
                cgrdDetail.TextMatrix(Row, Col) = mstrUndoMoney
            End If
        Case Else
    End Select
    
    sub�����ܽ��
    
    Exit Sub
errhandle:
    sfsub������ "�շѽ��沿��", "frm����", "cing�����嵥_AfterEdit", Err.Number, Err.Description, False
    
End Sub

Private Sub cgrdDetail_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error Resume Next
     
    Select Case Col
        Case �����嵥_����
            ctlb������.Buttons(1).Enabled = False
            mstrUndoCount = cgrdDetail.TextMatrix(Row, Col)
            
        Case �����嵥_����
            ctlb������.Buttons(1).Enabled = False
                        
            '��ȡ��С����,��󵥼�.
            Dim lobjRec As Object
            Set lobjRec = dafuncGetData("select * from �շѹ���_�շ���Ŀ�ֵ�� where �շ���Ŀ���='" & cgrdDetail.TextMatrix(Row, 0) & "'")
            If lobjRec.RecordCount > 0 Then
                mcur��С���� = IIf(IsNull(lobjRec("��С����").Value), 0, lobjRec("��С����").Value)
                mcur��󵥼� = IIf(IsNull(lobjRec("��󵥼�").Value), 99999999, lobjRec("��󵥼�").Value)
            Else
                sffuncMsg "δ�ҵ����շ���Ŀ��������Ϣ����������Ϣ�����ѱ��޸Ļ�ɾ�������˳��շѽ��棬���½��룡"
            End If
            mstrUndoMoney = cgrdDetail.TextMatrix(Row, Col)
        Case Else
            ctlb������.Buttons(1).Enabled = True
            Cancel = True
    End Select
End Sub



Private Sub cgrdDetail_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case vbKeyDelete
            mobj����ͨ�ö���_BeforeOperate "ɾ��", False
    End Select

End Sub

Private Sub cgrdDetail_LostFocus()
    On Error Resume Next
    ctlb������.Buttons(1).Enabled = True
End Sub


Private Sub cgrdItem_Click()
    On Error Resume Next
    ctxtFind.Text = ""
End Sub

Private Sub cgrdItem_DblClick()
    Dim lstrCode As String
    
    On Error GoTo errhandler
    '����շ���Ŀ
    lstrCode = cgrdItem.TextMatrix(cgrdItem.Row, 1)
    lstrCode = Right(lstrCode, Len(lstrCode) - InStr(lstrCode, " "))
    If InStr(lstrCode, " ") > 0 Then lstrCode = Right(lstrCode, Len(lstrCode) - InStr(lstrCode, " "))
    lstrCode = Trim(lstrCode)
    
    If Not func�����Ŀ�Ƿ���ѡ(lstrCode) Then
        sub�����Ŀ lstrCode
    End If
                    
Exit Sub
errhandler:
    sfsub������ "�շѽ��沿��", "frm����", "clst�շ���Ŀ_DblClick", Err.Number, Err.Description, False
          
End Sub

Private Sub sub�����Ŀ(ByVal paraCode As String)
    Dim lobjRec As Object
    
    Set lobjRec = dafuncGetData("select * from �շѹ���_�շ���Ŀ�ֵ�� where �շ���Ŀ���='" & paraCode & "'")
    cgrdDetail.AddItem paraCode & vbTab & lobjRec("�շ���Ŀ����") & vbTab & _
                    lobjRec("����") & vbTab & "1" & vbTab & lobjRec("����")
    
    sub�����ܽ��
End Sub


Private Sub sub�����ܽ��()
    Dim lcurMoney As Double
    Dim i As Long
    
    For i = 1 To cgrdDetail.Rows - 1
        lcurMoney = Format(lcurMoney + cgrdDetail.ValueMatrix(i, �����嵥_���), "0.00")
    Next

    clblTotal.Caption = "�ܽ�" & lcurMoney

End Sub


'���ܣ�������������Ƿ�������Ŀ��
Private Sub ctxtFind_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Dim i As Long
    If KeyCode = 13 Then
        '�س�ѡ����Ŀ��
        If cgrdItem.Row > 0 Then
            cgrdItem_DblClick
        End If
    Else
        '��λ��
        Dim lCol As Long
        cgrdItem.Row = 0
        If ctxtFind.Text <> "" Then
            For i = 1 To cgrdItem.Rows - 1
                If UCase(Left(cgrdItem.TextMatrix(i, 2), Len(ctxtFind))) = UCase(ctxtFind.Text) Then
                    cgrdItem.Select i, 0, i, cgrdItem.Cols - 1
                    cgrdItem.TopRow = i
                    Exit For
                End If
            Next
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 39 Then
        KeyAscii = 0
    ElseIf KeyAscii = 13 Then
        If ActiveControl = ctxtFind Then
        Else
            SendKeys Chr(vbKeyTab)
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim lcol������ As Collection
    Dim i As Long
    Dim lobjRec As Object
    
    On Error GoTo errhandle
    
    If pblnInUse Then Exit Sub
    pblnInUse = True
        
    Set mobj����ͨ�ö��� = New cls����ͨ�ö���
    Set mobj����ͨ�ö���.Form = Me
    Set mobj����ͨ�ö���.c������ = ctlb������
    
    Set lcol������ = New Collection
    
    lcol������.Add "����"
    lcol������.Add "|"
    lcol������.Add "ɾ��"
    lcol������.Add "���"
    lcol������.Add "|"
    lcol������.Add "�˳�"
    
    mobj����ͨ�ö���.subInitialize lcol������, ""
    
    mint��Ŀ���� = Val(pobj�շѹ���.ҵ������("��Ŀ����"))
    
    sub��ʼ������
    
    If pstr�շѱ�� <> "" Then
        '�ڲ��շ�,��ʾ������Ϣ��
        sub��ʾ������Ϣ
    Else
        mstr���ѵ�λ��� = pstr��λ���
        If pstr��λ��� <> "" Then
            Set lobjRec = dafuncGetData("select * from ��λ����_��λ������Ϣ�� where ������='" & pstr��λ��� & "'")
            If lobjRec.RecordCount > 0 Then
                ctxtInput(�շ�_���ѵ�λ) = lobjRec!��λ����
                ccmb��������.Text = lobjRec!��������
                ccmbƬ��.Text = IIf(IsNull(lobjRec!Ƭ��), "", lobjRec!Ƭ��)
            End If
        End If
    End If
    
    
    Exit Sub
errhandle:
    sfsub������ "�շѽ��沿��", "frm����", "Form_Load", Err.Number, Err.Description, False
    Exit Sub
    Resume
End Sub

Private Sub sub��ʾ������Ϣ()
    Dim lobjRec As Object
    Dim i As Long
    
    On Error GoTo errhandler
    
    If pstr�շѱ�� <> "" Then
        '�޸��շѼ�¼��
        Set lobjRec = dafuncGetData("select a.�շ�����,a.�շѱ��,a.�շ���Ŀ���,�շ���Ŀ����=(select �շ���Ŀ���� from �շѹ���_�շ���Ŀ�ֵ�� where �շ���Ŀ���=a.�շ���Ŀ���),a.����,������λ=(select ������λ from �շѹ���_�շ���Ŀ�ֵ�� where �շ���Ŀ���=a.�շ���Ŀ���),a.����,a.���,a.�շ�״̬,a.���ѷ�ʽ,a.������,a.���ѵ�λ���,���ѵ�λ=(select ��λ���� from ��λ����_��λ������Ϣ�� where ������=a.���ѵ�λ���),a.��������,a.�˷�����,�շ��˱��=a.�շ���,�շ���=(select ���� from ϵͳ����_Ա��������Ϣ�� where ���=a.�շ���),�˷��˱��=a.�˷���,�˷���=(select ���� from ϵͳ����_Ա��������Ϣ�� where ���=a.�˷���) ,���ܿ��Ҿ����˱��=a.���ܿ��Ҿ�����,���ܿ��Ҿ�����=(select ���� from ϵͳ����_Ա��������Ϣ�� where ���=a.���ܿ��Ҿ�����),���ܿ��ұ��,���ܿ���=(select ���� from ϵͳ����_�����ֵ�� where ���=a.���ܿ��ұ��),���۱���,��ע1,��ע2  from �շѹ���_������Ϣ�� a where �շѱ��='" & pstr�շѱ�� & "'  and �շ�״̬=0")
        
        cgrdDetail.Rows = 1
    
        Do While Not lobjRec.EOF
            cgrdDetail.AddItem lobjRec("�շ���Ŀ���") & vbTab & _
                lobjRec("�շ���Ŀ����") & vbTab & _
                lobjRec("����") & vbTab & _
                lobjRec("����") & vbTab & _
                lobjRec("���")
            lobjRec.MoveNext
        Loop
        If lobjRec.RecordCount > 0 Then
            lobjRec.MoveFirst
            mstr���ѵ�λ��� = IIf(IsNull(lobjRec("���ѵ�λ���").Value), "", lobjRec("���ѵ�λ���").Value)
        
            ctxtInput(�շ�_�շѱ��).Text = lobjRec("�շѱ��")
            
            If IIf(IsNull(lobjRec("���ܿ���")), "", lobjRec("���ܿ���")) <> "" Then
                For i = 0 To ccmb���ܿ���.ListCount - 1
                    If ccmb���ܿ���.List(i) = IIf(IsNull(lobjRec("���ܿ���")), "", lobjRec("���ܿ���")) Then
                        ccmb���ܿ���.ListIndex = i
                        Exit For
                    End If
                Next
            Else
                ccmb���ܿ���.ListIndex = -1
            End If
            
            ccmb��������.Text = IIf(IsNull(lobjRec("��ע1").Value), "", lobjRec("��ע1").Value)
            ccmbƬ��.Text = IIf(IsNull(lobjRec("��ע2").Value), "", lobjRec("��ע2").Value)
        
        
            ctxtInput(�շ�_������).Text = lobjRec("������")
            ctxtInput(�շ�_���ѵ�λ).Text = IIf(IsNull(lobjRec("���ѵ�λ").Value), "", lobjRec("���ѵ�λ").Value)
        
        End If
        
    End If

    Exit Sub
errhandler:
    sfsub������ "�շѽ��沿��", "frm����", "sub��ʾ������Ϣ", Err.Number, Err.Description, True
    Exit Sub
    Resume

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    pblnInUse = False
    Set mobj����ͨ�ö��� = Nothing
    
End Sub


Private Sub sub�������()
    Dim i As Integer
    
    On Error GoTo errhandle
    clblTotal = "�ܽ�0"
    
    '�ڱ������Ҫ����շѱ��;�켽��;2002/9/30
    mstr���ѵ�λ��� = ""
  
    Dim lobjCtrl As Control
    For Each lobjCtrl In ctxtInput
        lobjCtrl.Text = ""
    Next
    cgrdDetail.Rows = 1
    
    ccmb���ܿ���.Text = um�û���������
    
    ctxtInput(�շ�_������).SetFocus
    
    Exit Sub
errhandle:
    sfsub������ "�շѽ��沿��", "frm����", "sub����շѽ���", Err.Number, Err.Description, True
End Sub


Private Sub sub��ʼ������()
On Error GoTo errhandle
    Dim lobj�շѱ�׼ As Object
    Dim lobj���� As Object
    Dim lobj���ѷ�ʽ As Object

    mstrUndoCount = ""
    mstrUndoMoney = ""
    mstr���ѵ�λ��� = ""
    mint���ѷ�ʽ��� = 0
    
    Dim i As Long
    Dim j As Long
    
    Set lobj�շѱ�׼ = dafuncGetData("select �շѱ�׼����,���Ƿ� from �շѹ���_�շѱ�׼��Ϣ�� group by ���Ƿ�,�շѱ�׼����")
    Set lobj���� = dafuncGetData("select * from ϵͳ����_�����ֵ��")
    Set lobj���ѷ�ʽ = dafuncGetData("select * from �շѹ���_���ѷ�ʽ�ֵ��")
    
    
    '��ʼ�� "cing�����嵥"
    With cgrdDetail
        .Cols = 5
        .Rows = 1
        .TextMatrix(0, �����嵥_�շ���Ŀ���) = "�շ���Ŀ���"
        .ColWidth(�����嵥_�շ���Ŀ���) = 1310
        .ColAlignment(�����嵥_�շ���Ŀ���) = flexAlignCenterCenter
        
        .TextMatrix(0, �����嵥_�շ���Ŀ����) = "�շ���Ŀ����"
        .ColWidth(�����嵥_�շ���Ŀ����) = 1320
        
        .TextMatrix(0, �����嵥_����) = "����"
        .ColWidth(�����嵥_����) = 480
        
        .TextMatrix(0, �����嵥_����) = "����"
        .ColWidth(�����嵥_����) = 500
        
        .TextMatrix(0, �����嵥_���) = "���"
        .ColWidth(�����嵥_���) = 570
    End With
    
    '��ʼ�� "�շѱ�׼"
    ccmb�շѱ�׼.Clear
    Do While Not lobj�շѱ�׼.EOF
        ccmb�շѱ�׼.AddItem lobj�շѱ�׼("�շѱ�׼����").Value
        lobj�շѱ�׼.MoveNext
    Loop
    
    '��ʼ�� "���ܿ���"�б�
    ccmb���ܿ���.Clear
    If Not (lobj���� Is Nothing) Then
        Do While Not lobj����.EOF
            ccmb���ܿ���.AddItem lobj����("����").Value
            lobj����.MoveNext
        Loop
    End If
    ccmb���ܿ���.Text = um�û���������
    
    '��ȡ�շ���Ŀ���ࡣ
    Dim lobjRec As Object
    Set lobjRec = dafuncGetData("select �շ���Ŀ���,�շ���Ŀ���� from �շѹ���_�շ���Ŀ�ֵ�� where len(�շ���Ŀ���)=3  order by �շ���Ŀ��� ")
    Do While Not lobjRec.EOF
        Ccbo�շ���Ŀ����.AddItem lobjRec("�շ���Ŀ����")
        lobjRec.MoveNext
    Loop
    
    Ccbo�շ���Ŀ����.ListIndex = 0
    
    '��ȡ��������
    Set lobjRec = dafuncGetData("select * from ϵͳ����_���������ֵ���ͼ order by ���")
    ccmb��������.Clear
    ccmb��������.AddItem ""
    Do While Not lobjRec.EOF
        ccmb��������.AddItem lobjRec("����").Value
        lobjRec.MoveNext
    Loop
    
    '��ȡƬ��
    Set lobjRec = dafuncGetData("select * from ϵͳ����_Ƭ���ֵ���ͼ order by ���")
    ccmbƬ��.Clear
    ccmbƬ��.AddItem ""
    Do While Not lobjRec.EOF
        ccmbƬ��.AddItem lobjRec("����").Value
        lobjRec.MoveNext
    Loop
    
    Exit Sub
errhandle:
    sfsub������ "�շѽ��沿��", "frm����", "sub��ʼ������", Err.Number, Err.Description, True
End Sub






Private Sub mobj����ͨ�ö���_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    Dim i As Long, j As Long
    
    On Error GoTo errhandle
    Select Case Operate
        Case "����"
            Cancel = True
            'У�����ݺϷ��ԡ�
            If Not ValidateData Then Exit Sub
            
            '�ռ�Ҫ����ķ�����Ϣ��
            Dim lstr���ܿ��ұ�� As String
            Dim lcol��¼ As Collection
            Dim lcol���� As Collection
            Dim lstr�շѱ�� As String
            
            If ccmb���ܿ���.ListIndex >= 0 Then
                lstr���ܿ��ұ�� = ccmb���ܿ���.ItemData(ccmb���ܿ���.ListIndex)
                lstr���ܿ��ұ�� = Right(lstr���ܿ��ұ��, Len(lstr���ܿ��ұ��) - 1)
            Else
                lstr���ܿ��ұ�� = um�û��������ұ��
            End If
            Set lcol���� = New Collection
            For i = 1 To cgrdDetail.Rows - 1
                Set lcol��¼ = New Collection
                For j = 0 To cgrdDetail.Cols - 1
                    lcol��¼.Add cgrdDetail.TextMatrix(i, j), cgrdDetail.TextMatrix(0, j)
                Next
                '����շ������ֶ�
                lcol��¼.Add ctxtInput(�շ�_������).Text, "������"
                lcol��¼.Add mstr���ѵ�λ���, "���ѵ�λ���"
                lcol��¼.Add ctxtInput(�շ�_���ѵ�λ).Text, "���ѵ�λ����"
                lcol��¼.Add lstr���ܿ��ұ��, "���ܿ��ұ��"
                lcol��¼.Add um�û����, "���ܿ��Ҿ�����"
                lcol��¼.Add ccmb��������.Text, "��ע1"
                lcol��¼.Add ccmbƬ��.Text, "��ע2"
                lcol����.Add lcol��¼
            Next
            
            '���滮����Ϣ��
            lstr�շѱ�� = pobj�շѹ���.func���۱���(lcol����, ctxtInput(�շ�_�շѱ��), pstrҵ�����)
            
            If cchk���.Value = 1 Then
                sub�������
            Else
                ctxtInput(�շ�_�շѱ��) = lstr�շѱ��
            End If
            pstr�շѱ�� = lstr�շѱ��
            ctxtInput(�շ�_������).SetFocus
            
        Case "ɾ��"
            If cgrdDetail.Row > 0 Then
                cgrdDetail.RemoveItem cgrdDetail.Row
                
                sub�����ܽ��
            End If
            
        Case "���"
            sub�������
            
        Case Else
    End Select
    Exit Sub
    
errhandle:
    sfsub������ "�շѽ��沿��", "frm����", "mobj����ͨ�ö���_BeforeOperate", Err.Number, Err.Description, False
    Exit Sub
    Resume
End Sub


Private Function func�����Ŀ�Ƿ���ѡ(ByVal Para�շ���Ŀ��� As String) As Boolean
On Error GoTo errhandle
    Dim i As Long
    func�����Ŀ�Ƿ���ѡ = False
    If cgrdDetail.Rows = 1 Then
        func�����Ŀ�Ƿ���ѡ = False
        Exit Function
    End If
    
    For i = 1 To cgrdDetail.Rows - 1
        If Para�շ���Ŀ��� = cgrdDetail.TextMatrix(i, �����嵥_�շ���Ŀ���) Then
            func�����Ŀ�Ƿ���ѡ = True
            Exit Function
        End If
    Next
Exit Function
errhandle:
    sfsub������ "�շѽ��沿��", "frm����", " func�����Ŀ�Ƿ���ѡ()", Err.Number, Err.Description
End Function

Private Function ValidateData() As Boolean
    On Error GoTo errhandle
    ValidateData = False
    If ctxtInput(�շ�_������).Text = vbNullString And ctxtInput(�շ�_���ѵ�λ) = vbNullString Then
        sffuncMsg """������"" �� ""���ѵ�λ"" ������������֮һ��", sf����
        Exit Function
    End If
    If cgrdDetail.Rows = 1 Then
        sffuncMsg "�޷�����Ϣ���Ա��棡", sf����
        Exit Function
    End If
    

    ValidateData = True
    Exit Function
errhandle:
    sfsub������ "�շѽ��沿��", "frm����", " ValidateData()", Err.Number, Err.Description, True
End Function


