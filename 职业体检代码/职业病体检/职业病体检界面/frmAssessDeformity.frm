VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAssessDeformity 
   Caption         =   "Form1"
   ClientHeight    =   9210
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14910
   LinkTopic       =   "Form1"
   ScaleHeight     =   9210
   ScaleWidth      =   14910
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame Frame3 
      Caption         =   "¼������"
      Height          =   8295
      Left            =   10080
      TabIndex        =   15
      Top             =   840
      Width           =   4455
      Begin VB.Frame Frame��ʷ���� 
         Caption         =   "�ϴ�����"
         Height          =   3135
         Left            =   120
         TabIndex        =   31
         Top             =   5040
         Visible         =   0   'False
         Width           =   4215
         Begin VB.TextBox Text��ʷ���� 
            Height          =   495
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   43
            Top             =   2520
            Width           =   3975
         End
         Begin VB.TextBox Text��ʷ��� 
            Height          =   735
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   41
            Top             =   1440
            Width           =   3975
         End
         Begin VB.TextBox Text��ʷʱ�� 
            Height          =   375
            Left            =   720
            TabIndex        =   35
            Top             =   720
            Width           =   3255
         End
         Begin VB.TextBox Text��ʷ�ȼ� 
            Height          =   375
            Left            =   720
            TabIndex        =   33
            Top             =   240
            Width           =   3255
         End
         Begin VB.Label Label14 
            Caption         =   "�����ۣ�"
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   2280
            Width           =   975
         End
         Begin VB.Label Label13 
            Caption         =   "�������"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label Label10 
            Caption         =   "ʱ�䣺"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label9 
            Caption         =   "�ȼ���"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   735
         End
      End
      Begin MSComCtl2.DTPicker DTPickerʱ�� 
         Height          =   375
         Left            =   840
         TabIndex        =   26
         Top             =   4560
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
         _Version        =   393216
         Format          =   59637761
         CurrentDate     =   42520
      End
      Begin VB.TextBox Text�ȼ� 
         Height          =   375
         Left            =   840
         TabIndex        =   25
         Top             =   4080
         Width           =   3495
      End
      Begin VB.TextBox Text���� 
         Height          =   855
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Top             =   2880
         Width           =   4215
      End
      Begin VB.TextBox Text���� 
         Height          =   735
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   1800
         Width           =   4215
      End
      Begin VB.TextBox Text��� 
         Height          =   975
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Top             =   480
         Width           =   4215
      End
      Begin VB.Label Label8 
         Caption         =   "ʱ�䣺"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   4680
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "�ȼ���"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "���У�"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   3840
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "�ṩ���ϣ�"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "�����ۣ�"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "�������"
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "��ѯ����"
      Height          =   6495
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   3735
      Begin VB.CheckBox Check���� 
         Caption         =   "����"
         Height          =   255
         Left            =   2160
         TabIndex        =   45
         Top             =   360
         Width           =   735
      End
      Begin VB.CheckBox Check���� 
         Caption         =   "����"
         Height          =   255
         Left            =   960
         TabIndex        =   44
         Top             =   360
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.OptionButton coptType 
         Caption         =   "�ṩ֤��"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   39
         Top             =   4440
         Width           =   1095
      End
      Begin VB.OptionButton coptType 
         Caption         =   "�Ѹ���"
         Height          =   255
         Index           =   2
         Left            =   2640
         TabIndex        =   38
         Top             =   4920
         Width           =   975
      End
      Begin VB.ComboBox Combo��� 
         Height          =   300
         Left            =   1440
         TabIndex        =   30
         Top             =   1320
         Width           =   2055
      End
      Begin VB.CheckBox Check��� 
         Caption         =   "��Ա���"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton Com��ѯ 
         Caption         =   "��ѯ"
         Height          =   375
         Left            =   1200
         TabIndex        =   14
         Top             =   5640
         Width           =   1095
      End
      Begin VB.OptionButton coptType 
         Caption         =   "����δ����"
         Height          =   195
         Index           =   1
         Left            =   1200
         TabIndex        =   13
         Top             =   4920
         Width           =   1215
      End
      Begin VB.OptionButton coptType 
         Caption         =   "������"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   4800
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.TextBox Text��λ 
         Height          =   270
         Left            =   1440
         TabIndex        =   11
         Top             =   3720
         Width           =   2055
      End
      Begin VB.CheckBox Check��λ 
         Caption         =   "��λ���ƣ�"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   3720
         Width           =   1215
      End
      Begin VB.TextBox Text��� 
         Height          =   270
         Left            =   1440
         TabIndex        =   9
         Top             =   3120
         Width           =   2055
      End
      Begin VB.CheckBox Check��� 
         Caption         =   "����ţ�"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   3120
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPicker���� 
         Height          =   255
         Left            =   1440
         TabIndex        =   6
         Top             =   2400
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         _Version        =   393216
         Format          =   59637760
         CurrentDate     =   42520
      End
      Begin MSComCtl2.DTPicker DTPicker��ʼ 
         Height          =   255
         Left            =   1440
         TabIndex        =   5
         Top             =   1920
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         _Version        =   393216
         Format          =   59637760
         CurrentDate     =   42520
      End
      Begin VB.CheckBox Check���� 
         Caption         =   "������ڣ�"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1920
         Width           =   1215
      End
      Begin VB.ComboBox Combo���� 
         Height          =   300
         Left            =   1440
         TabIndex        =   3
         Top             =   840
         Width           =   2055
      End
      Begin VB.CheckBox Check���� 
         Caption         =   "��Ա���ͣ�"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "��"
         Height          =   255
         Left            =   1080
         TabIndex        =   7
         Top             =   2400
         Width           =   375
      End
   End
   Begin MSComDlg.CommonDialog ccmdFile 
      Left            =   9600
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VSFlex8Ctl.VSFlexGrid cgrdInfo 
      Height          =   6015
      Left            =   4200
      TabIndex        =   27
      Top             =   1440
      Width           =   5775
      _cx             =   10186
      _cy             =   10610
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
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   0
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
      AutoSearchDelay =   2
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
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin MSComctlLib.Toolbar ctlb������ 
      Height          =   600
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   1058
      ButtonWidth     =   1455
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "cimg��ťͼ��"
      _Version        =   393216
      Begin MSComctlLib.ImageList cimg��ťͼ�� 
         Left            =   240
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   393216
      End
   End
   Begin VB.Label Label12 
      Caption         =   "��������"
      Height          =   255
      Left            =   4200
      TabIndex        =   37
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label11 
      Caption         =   "��������"
      Height          =   255
      Left            =   5040
      TabIndex        =   36
      Top             =   1200
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "���Ǳ���"
      Height          =   255
      Left            =   4200
      TabIndex        =   28
      Top             =   960
      Width           =   5775
   End
End
Attribute VB_Name = "frmAssessDeformity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mblnInUse As Boolean
Private WithEvents mobjGUI As cls����ͨ�ö���
Attribute mobjGUI.VB_VarHelpID = -1
Private mobjQueryResult As Object

Private Sub cgrdInfo_Click()
subClear
'Com��ѯ_Click
Dim SyNo As String
Dim obj As Object
SyNo = cgrdInfo.TextMatrix(cgrdInfo.Row, cgrdInfo.ColIndex("���"))
Set obj = dafuncGetData("select * from ְҵ�����_������Ϣ�� where ϵͳ���='" & SyNo & "'")
If obj.RecordCount > 0 Then
    If Not IsNull(obj("�ṩ������Ϣ")) Then
        Text����.Text = obj("�ṩ������Ϣ")
    End If
    If Not IsNull(obj("���еȼ�")) Then
        Text�ȼ�.Text = obj("���еȼ�")
    End If
    If Not IsNull("����ʱ��") Then
        DTPickerʱ��.Value = obj("����ʱ��")
    End If

'    '�ж��Ƿ��Ѿ�����
'    If Not IsNull(obj("�Ƿ�������")) Then
'        If obj("�Ƿ�������") = 1 And coptType(1).Value = True Then '1������ʷ�����У�0����û��������
'            Frame��ʷ����.Visible = True
'            Text��ʷ�ȼ�.Text = IIf(Not IsNull(obj("��ʷ���еȼ�")), obj("��ʷ���еȼ�"), "")
''            Text��ʷ�ȼ�.Text = obj("��ʷ���еȼ�")
'            Text��ʷʱ��.Text = IIf(Not IsNull(obj("��ʷ����ʱ��")), obj("��ʷ����ʱ��"), "")
''            Text��ʷʱ��.Text = obj("��ʷ����ʱ��")
'        End If
'    Else
'        Frame��ʷ����.Visible = False
'    End If
End If
'����ǵ��У���ʾ��������
'If Check����.Value = 1 Then
    '���֤���ж���ǰ�Ƿ��й�������ʷ
    Dim obj1 As Object
    Set obj1 = dafuncGetData("select a.���еȼ�,a.����ʱ��,a.�����,a.������ from ְҵ�����_������ϸ��ʷ��Ϣ�� a left join ְҵ�����_�����Ա������Ϣ�� b on a.���֤��=b.������ݺ��� where b.������ݺ���='" & cgrdInfo.TextMatrix(cgrdInfo.Row, cgrdInfo.ColIndex("���֤��")) & "' order by a.���ʱ�� desc")
    If obj1.RecordCount > 0 Then
        Frame��ʷ����.Visible = True
        Text��ʷ�ȼ�.Text = obj1("���еȼ�")
        Text��ʷʱ��.Text = obj1("����ʱ��")
        Text��ʷ���.Text = obj1("�����")
        Text��ʷ����.Text = obj1("������")
    Else
        Frame��ʷ����.Visible = False
    End If
'End If
Text���.Text = cgrdInfo.TextMatrix(cgrdInfo.Row, cgrdInfo.ColIndex("�����"))
Text����.Text = cgrdInfo.TextMatrix(cgrdInfo.Row, cgrdInfo.ColIndex("������"))
End Sub

Private Sub Check���_Click()
    If Check���.Value = 0 Then
        Combo���.Text = ""
    End If
End Sub

Private Sub Check����_Click()
    If Check����.Value = 0 Then
        Combo����.Text = ""
    End If
End Sub

Private Sub Com��ѯ_Click()
On Error Resume Next
Dim lsql���� As String
Dim lsql��� As String
Dim lsql��ʼ���� As Date
Dim lsql�������� As Date
Dim lsql��� As String
Dim lsql��λ As String
Dim lsqlwhere As String
Dim lsql��ѯ As String

lsqlwhere = ""
'�����������Ҫ������Ա
If Check����.Value = 1 Then
    If coptType(1).Value = True Or coptType(2).Value = True Then
        lsqlwhere = lsqlwhere + " b.��Ϻʹ������ like '%���ų�%' and a.ϵͳ���=b.ϵͳ��� "
    ElseIf coptType(0).Value = True Then
        lsqlwhere = lsqlwhere + "(b.���״̬='11' or b.��Ϻʹ������ like '%���ų�%') and a.ϵͳ���=b.ϵͳ��� "
    Else
        lsqlwhere = lsqlwhere + "(b.���״̬='6' or b.���״̬='7' or b.���״̬='8') and a.ϵͳ���=b.ϵͳ��� "
    End If
Else   '����   ��ʱû��ȡֵ��δ���
    If coptType(1).Value = True Or coptType(2).Value = True Then
        lsqlwhere = lsqlwhere + " b.��Ϻʹ������ like '%���ų�%' and a.ϵͳ���=b.ϵͳ��� "
    ElseIf coptType(0).Value = True Then
        lsqlwhere = lsqlwhere + "(b.���״̬='11' or b.��Ϻʹ������ like '%���ų�%') and a.ϵͳ���=b.ϵͳ��� "
    Else
        lsqlwhere = lsqlwhere + "(b.���״̬='6' or b.���״̬='7' or b.���״̬='8') and a.ϵͳ���=b.ϵͳ��� "
    End If
End If
'��װ��ѯ����
'1.�������
    If Check����.Value = 1 Then
        lsql���� = Combo����.Text
        lsqlwhere = lsqlwhere + "and a.��������='" & lsql���� & "' "
    Else
        Combo����.Text = ""
        lsql���� = ""
    End If
'2.������
    If Check���.Value = 1 Then
        lsql��� = Combo���.Text
        lsqlwhere = lsqlwhere + " and a.�������='" & lsql��� & "' "
    Else
        Combo���.Text = ""
        lsql��� = ""
    End If
'3.���
    If Check���.Value = 1 Then
        lsql��� = Text���.Text
        lsqlwhere = lsqlwhere + "and b.ϵͳ���='" & lsql��� & "' "
    Else
        Text���.Text = ""
        lsql��� = ""
    End If
'4.��λ
    If Check��λ.Value = 1 Then
        lsql��λ = Text��λ.Text
        lsqlwhere = lsqlwhere + "and a.��λ����='" & lsql��λ & "' "
    Else
        Text��λ.Text = ""
        lsql��λ = ""
    End If
'5.�������
    If Check����.Value = 1 Then
        lsql��ʼ���� = DTPicker��ʼ.Value
        lsql�������� = DTPicker����.Value
        lsqlwhere = lsqlwhere + " and b.������� between '" & DTPicker��ʼ.Value & " 00:00:00' and '" & DTPicker����.Value & " 23:59:59'"
    Else
        lsql��ʼ���� = ""
        lsql�������� = ""
    End If
If Check����.Value = 1 Then
    If coptType(0).Value = True Then
        ctlb������.Buttons(3).Visible = False
        ctlb������.Buttons(4).Visible = True
        ctlb������.Buttons(5).Visible = False
        ctlb������.Buttons(6).Visible = False
        ctlb������.Buttons(8).Visible = False
'        lsqlwhere = lsqlwhere + " and c.ϵͳ��� is null"
        lsqlwhere = lsqlwhere + " and (c.ϵͳ��� is null or c.����״̬='1')"
        lsql��ѯ = "select b.ϵͳ��� as ���,a.������ݺ��� as ���֤��,convert(varchar(10),b.�������,111) ���ʱ�� ,a.��λ���� as ����,a.����,a.�Ա�,a.����,b.������ as �����,b.��Ϻʹ������ as ������ from ְҵ�����_�����Ա������Ϣ�� a ,ְҵ�����_��������Ϣ�� b left join ְҵ�����_������Ϣ�� c on b.ϵͳ���=c.ϵͳ��� where " & lsqlwhere & ""
    ElseIf coptType(1).Value = True Then
        ctlb������.Buttons(3).Visible = True
        ctlb������.Buttons(5).Visible = True
        ctlb������.Buttons(6).Visible = True
        ctlb������.Buttons(8).Visible = False
        lsqlwhere = lsqlwhere + " and c.ϵͳ���=b.ϵͳ��� and ����״̬='2'"
        lsql��ѯ = "select b.ϵͳ��� as ���,a.������ݺ��� as ���֤��,convert(varchar(10),b.�������,111) ���ʱ��,a.��λ���� as ����,a.����,a.�Ա�,a.����,b.������ as �����,b.��Ϻʹ������ as ������ from ְҵ�����_�����Ա������Ϣ�� a ,ְҵ�����_��������Ϣ�� b, ְҵ�����_������Ϣ�� c where " & lsqlwhere & ""
    ElseIf coptType(2).Value = True Then
        ctlb������.Buttons(4).Visible = False
        ctlb������.Buttons(5).Visible = False
        ctlb������.Buttons(6).Visible = False
        ctlb������.Buttons(8).Visible = True
        lsqlwhere = lsqlwhere + " and c.ϵͳ���=b.ϵͳ��� and ����״̬='3'"
        lsql��ѯ = "select b.ϵͳ��� as ���,a.������ݺ��� as ���֤��,convert(varchar(10),b.�������,111) ���ʱ��,a.��λ���� as ����,a.����,a.�Ա�,a.����,b.������ as �����,b.��Ϻʹ������ as ������ from ְҵ�����_�����Ա������Ϣ�� a ,ְҵ�����_��������Ϣ�� b, ְҵ�����_������Ϣ�� c where " & lsqlwhere & ""
    End If
Else   '����  δ���
    '��������
End If
    Set mobjQueryResult = dafuncGetData(lsql��ѯ)
    Set cgrdInfo.DataSource = mobjQueryResult
    '�������֤����
    cgrdInfo.ColHidden(cgrdInfo.ColIndex("���֤��")) = True
    'ȡ����
    Label1.Caption = Format(Now, "yyyy") + "��" + IIf(Combo����.Text = "��˲���YK", "�˿󲿶�", IIf(Combo����.Text = "8023����", "ԭ8023����", Combo����.Text)) + "������Ա" + Combo���.Text + "���ų��˷���Ӱ����Աһ����"
    Label11.Caption = cgrdInfo.rows - 1
    'vsflexgrid�п�Ȱ������Զ������������ͷ����ͷ����������
    cgrdInfo.AutoSize 0, cgrdInfo.cols - 1, 0, 0
    cgrdInfo.ExplorerBar = flexExSort
    cgrdInfo.DataMode = flexDMFree
    cgrdInfo.Col = 0
    cgrdInfo.Sort = flexSortGenericDescending
    subClear
    ctlb������.Buttons(3).Visible = False
    Frame��ʷ����.Visible = False
End Sub

Private Sub coptType_Click(Index As Integer)
    Com��ѯ_Click
    If cgrdInfo.rows > 1 Then
        cgrdInfo.AutoSize 0, cgrdInfo.cols - 1, 0, 0
        cgrdInfo.ExplorerBar = flexExSort
        cgrdInfo.DataMode = flexDMFree
        cgrdInfo.Col = 0
        cgrdInfo.Sort = flexSortGenericDescending
    End If
End Sub

Private Sub Form_Load()
 On Error Resume Next
Dim lcol��������ť As New Collection '�������ϵİ�ť��ʼ�����ϡ�
 Set mobjGUI = New cls����ͨ�ö���
     '��������Ѿ���ʼ�������ٽ��г�ʼ����
    If mblnInUse Then Exit Sub
    
        '���ô�������ʹ�õı�־��
    mblnInUse = True
    
    '���ù�����������Ҫ�ĸ��ְ�ť��
    With lcol��������ť
        .Add "����Excel(&O)113"     '1
        .Add "|"
        .Add "ɾ��"            '3
        .Add "����(&S)101"     '4
        .Add "ȡ������(&E)111"     '5
        .Add "����(&F)109"     '6
        .Add "|"
        .Add "ȡ������(&E)111" '8
        .Add "|"
        .Add "�˳�"            '10
    End With
    'Ϊ��Ҫͨ������ͨ�ö�����Ƶĸ��ֿؼ�����ʼֵ��
    With mobjGUI
        Set .Form = Me
        Set .c������ = ctlb������
    End With
    mobjGUI.subInitialize lcol��������ť, ""
    DoEvents
    
    With cgrdInfo
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "���"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "���֤��"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "���ʱ��"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "����"       '���������ֵ����ʱȡ�ĵ�λ
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "����"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "�Ա�"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "����"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "�����"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "������"
        .AutoSize 0, .cols - 1, 0, 0
        .SelectionMode = flexSelectionListBox
    End With
    
    '��ʼ�������ѯ����
    listcombox
    Check����.Value = 1
    Check���.Value = 1
    Check����.Value = 1
'    coptType(0).Value = True
    DTPicker����.Value = Format(Now, "yyyy-MM-dd")
    DTPicker��ʼ.Value = Format(DateAdd("M", -5, Now()), "yyyy/MM/dd")
    DTPickerʱ��.Value = Now
    Com��ѯ_Click
End Sub

Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
 Dim lobj As Object
 On Error Resume Next
Cancel = True
Select Case Operate
    Case "����Excel"
        If cgrdInfo.rows <= 1 Then
            MsgBox "û����Ҫ�����ļ�¼��", vbOKOnly + vbExclamation, "ϵͳ��ʾ"
            Exit Sub
        End If
        
        Dim lstrFile As String
        ccmdFile.Filter = "Excel�ļ� (*.xls)|*.xls|�ı��ļ� (*.txt)|*.txt"
        ccmdFile.ShowSave
        lstrFile = ccmdFile.FileName
        If lstrFile <> "" Then
            '��Ϊ��0�У�Ϊϵͳ��š��������б���ʱΪstring
            cgrdInfo.ColDataType(cgrdInfo.ColIndex("���")) = flexDTString
            cgrdInfo.SaveGrid lstrFile, flexFileExcel, True   '����excelϵͳ���Ϊ����
            'cgrdInfo.SaveGrid lstrFile, flexFileTabText, True
        End If
        MsgBox "�������"
        
    Case "ɾ��"
        If MsgBox("��ȷ��Ҫɾ��������¼��", vbYesNo + vbQuestion + vbDefaultButton2, "ϵͳ��ʾ") = vbYes Then
            dafuncGetData ("delete from ְҵ�����_������Ϣ�� where ϵͳ���='" & cgrdInfo.TextMatrix(cgrdInfo.Row, cgrdInfo.ColIndex("���")) & "'")
            Com��ѯ_Click
            MsgBox "�ѳɹ�ɾ��������¼��"
        End If
        
    Case "����"
'        Dim lobj As Object
        Set lobj = dafuncGetData("select * from ְҵ�����_������Ϣ�� where ϵͳ���='" & cgrdInfo.TextMatrix(cgrdInfo.Row, cgrdInfo.ColIndex("���")) & "'")
        If lobj.RecordCount > 0 Then
            dafuncGetData ("update ְҵ�����_������Ϣ�� set ���֤��='" & cgrdInfo.TextMatrix(cgrdInfo.Row, cgrdInfo.ColIndex("���֤��")) & "',�������='" & cgrdInfo.TextMatrix(cgrdInfo.Row, cgrdInfo.ColIndex("���ʱ��")) & "',�����='" & Text���.Text & "',  ������='" & Text����.Text & "',�ṩ������Ϣ='" & Text����.Text & "',���еȼ�='" & Text�ȼ�.Text & "',����ʱ��='" & DTPickerʱ��.Value & "',����״̬='2' where ϵͳ���='" & cgrdInfo.TextMatrix(cgrdInfo.Row, cgrdInfo.ColIndex("���")) & "'")
            Com��ѯ_Click
            MsgBox "����ɹ���"
        Else
            dafuncGetData ("insert into ְҵ�����_������Ϣ��(ϵͳ���,���֤��,�������,�����,������,�ṩ������Ϣ,���еȼ�,����ʱ��,����״̬) values('" & cgrdInfo.TextMatrix(cgrdInfo.Row, cgrdInfo.ColIndex("���")) & "'," & cgrdInfo.TextMatrix(cgrdInfo.Row, cgrdInfo.ColIndex("���֤��")) & ",'" & cgrdInfo.TextMatrix(cgrdInfo.Row, cgrdInfo.ColIndex("���ʱ��")) & "','" & Text���.Text & "','" & Text����.Text & "','" & Text����.Text & "','" & Text�ȼ�.Text & "','" & DTPickerʱ��.Value & "','2')")
            Com��ѯ_Click
            MsgBox "����ɹ���"
        End If
        
    Case "ȡ������"
        If MsgBox("��ȷ��Ҫȡ���ñ��������", vbYesNo + vbQuestion + vbDefaultButton2, "ϵͳ��ʾ") = vbYes Then
            dafuncGetData ("update ְҵ�����_������Ϣ�� set ����״̬='1' where ϵͳ���='" & cgrdInfo.TextMatrix(cgrdInfo.Row, cgrdInfo.ColIndex("���")) & "'")
            Com��ѯ_Click
            MsgBox "��ȡ������"
        End If
'        dafuncGetData ("update ְҵ�����_������Ϣ�� set ����״̬='1' where ϵͳ���='" & cgrdInfo.TextMatrix(cgrdInfo.Row, cgrdInfo.ColIndex("���")) & "'")
    
    Case "����"
            dafuncGetData ("update ְҵ�����_������Ϣ�� set ����״̬='3',�Ƿ�������='1',��ʷ���еȼ�='" & Text�ȼ�.Text & "',��ʷ����ʱ��='" & DTPickerʱ��.Value & "',��ʷ������Ϣ='" & Text����.Text & "'where ϵͳ���='" & cgrdInfo.TextMatrix(cgrdInfo.Row, cgrdInfo.ColIndex("���")) & "'")
            Dim detail As Object
            Set detail = dafuncGetData("select * from ְҵ�����_������ϸ��ʷ��Ϣ�� where ϵͳ���='" & cgrdInfo.TextMatrix(cgrdInfo.Row, cgrdInfo.ColIndex("���")) & "'")
            If detail.RecordCount > 0 Then
                dafuncGetData ("update ְҵ�����_������ϸ��ʷ��Ϣ�� set ���֤��='" & cgrdInfo.TextMatrix(cgrdInfo.Row, cgrdInfo.ColIndex("���֤��")) & "',�������='" & cgrdInfo.TextMatrix(cgrdInfo.Row, cgrdInfo.ColIndex("���ʱ��")) & "',�����='" & Text���.Text & "',  ������='" & Text����.Text & "',�ṩ������Ϣ='" & Text����.Text & "',���еȼ�='" & Text�ȼ�.Text & "',����ʱ��='" & DTPickerʱ��.Value & "' where ϵͳ���='" & cgrdInfo.TextMatrix(cgrdInfo.Row, cgrdInfo.ColIndex("���")) & "'")
            Else
                dafuncGetData ("insert into ְҵ�����_������ϸ��ʷ��Ϣ��(���֤��,ϵͳ���,���ʱ��,�����,������,������Ϣ,���еȼ�,����ʱ��) values('" & cgrdInfo.TextMatrix(cgrdInfo.Row, cgrdInfo.ColIndex("���֤��")) & "'," & cgrdInfo.TextMatrix(cgrdInfo.Row, cgrdInfo.ColIndex("���")) & ",'" & cgrdInfo.TextMatrix(cgrdInfo.Row, cgrdInfo.ColIndex("���ʱ��")) & "','" & Text���.Text & "','" & Text����.Text & "','" & Text����.Text & "','" & Text�ȼ�.Text & "','" & DTPickerʱ��.Value & "')")
            End If
            Com��ѯ_Click
    
    Case "ȡ������"
        If MsgBox("��ȷ��Ҫȡ���ø��˼�¼��", vbYesNo + vbQuestion + vbDefaultButton2, "ϵͳ��ʾ") = vbYes Then
            dafuncGetData ("update ְҵ�����_������Ϣ�� set ����״̬='2',�Ƿ�������='0' where ϵͳ���='" & cgrdInfo.TextMatrix(cgrdInfo.Row, cgrdInfo.ColIndex("���")) & "'")
            Com��ѯ_Click
            MsgBox "��ȡ������"
        End If
        
    Case "�˳�"
        Unload Me
End Select
End Sub
'�˳�����ʱ����ղ��ֱ���
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    mblnInUse = False
    Set mobjGUI = Nothing
    Set mobjQueryResult = Nothing
End Sub
'��ʼ�������б�ֵ
Sub listcombox()
'1��Ա����
    Combo����.Text = ""
    Combo����.AddItem "8023����": Combo����.ItemData(Combo����.NewIndex) = 0
    Combo����.AddItem "��˲���": Combo����.ItemData(Combo����.NewIndex) = 1
    Combo����.AddItem "��˲���YK": Combo����.ItemData(Combo����.NewIndex) = 2
    Combo����.ListIndex = 0
 '������
    Combo���.Text = ""
    Combo���.AddItem "����": Combo���.ItemData(Combo���.NewIndex) = 0
    Combo���.AddItem "����": Combo���.ItemData(Combo���.NewIndex) = 1
    Combo���.ListIndex = 1
End Sub
Sub subClear()
    '��ո����ı���
    Text���.Text = ""
    Text����.Text = ""
    Text����.Text = ""
    Text�ȼ�.Text = ""
    DTPickerʱ��.Value = Now
End Sub
