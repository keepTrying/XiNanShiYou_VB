VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frm���Զ��� 
   Caption         =   "����"
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8865
   LinkTopic       =   "Form1"
   ScaleHeight     =   7920
   ScaleWidth      =   8865
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command7 
      Caption         =   "Sub���ݵ���"
      Height          =   435
      Left            =   5040
      TabIndex        =   12
      Top             =   6120
      Width           =   3495
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Func��ȡmdb���ݽ����ļ��е����ݷ�Χ"
      Height          =   435
      Left            =   5040
      TabIndex        =   11
      Top             =   5040
      Width           =   3495
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Func��ȡmdb�ļ��е����ݷ����嵥"
      Height          =   435
      Left            =   5040
      TabIndex        =   10
      Top             =   5640
      Width           =   3495
   End
   Begin VB.CommandButton Command8 
      Caption         =   "func�鿴����"
      Height          =   435
      Left            =   5040
      TabIndex        =   9
      Top             =   4560
      Width           =   3495
   End
   Begin VSFlex6DAOCtl.vsFlexGrid cgrdMain 
      Height          =   2055
      Left            =   120
      TabIndex        =   8
      Top             =   2400
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   3625
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
      Rows            =   50
      Cols            =   10
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
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   7320
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command6 
      Caption         =   "sub�����ļ���"
      Height          =   435
      Left            =   120
      TabIndex        =   6
      Top             =   6960
      Width           =   1995
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Sub���������Ա�Ǽ�"
      Height          =   435
      Left            =   120
      TabIndex        =   5
      Top             =   6480
      Width           =   1995
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Func�������ļ��������Ա��Ϣ"
      Height          =   435
      Left            =   2160
      TabIndex        =   4
      Top             =   4560
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "���������ʱ�־"
      Height          =   435
      Left            =   120
      TabIndex        =   3
      Top             =   5880
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "func��ȡ����������Ϣ"
      Height          =   435
      Left            =   120
      TabIndex        =   2
      Top             =   5280
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   2055
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   120
      Width           =   8655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   435
      Left            =   120
      TabIndex        =   0
      Top             =   4560
      Width           =   1755
   End
End
Attribute VB_Name = "frm���Զ���"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mobj������ӿ� As ClsManageTransmission

Private Sub Command8_Click()
    Dim lobjRec As Object
    Dim lcolRange As Collection
    Dim lcolItem As Collection
    Dim llngType As Integer
    If MsgBox("�鿴mdb������������ѡ���ǡ�����鿴mdb�⣬����鿴�������ݿ⡣", vbYesNo + vbQuestion + vbDefaultButton1, "ϵͳѯ��") = vbYes Then
        llngType = 0
    Else
        llngType = 1
    End If
    Set lcolRange = New Collection
    Set lcolItem = New Collection
    lcolItem.Add "��ʼ����", "���ݷ�Χ��"
    lcolItem.Add "2001-3-27", "���ݷ�Χֵ"
    lcolRange.Add lcolItem, lcolItem("���ݷ�Χ��")
    
    Set lcolItem = New Collection
    lcolItem.Add "��������", "���ݷ�Χ��"
    lcolItem.Add "2001-3-28", "���ݷ�Χֵ"
    lcolRange.Add lcolItem, lcolItem("���ݷ�Χ��")
    
    Set lcolItem = New Collection
    lcolItem.Add "��λ���Ƽ�", "���ݷ�Χ��"
    lcolItem.Add "asd", "���ݷ�Χֵ"
    lcolRange.Add lcolItem, lcolItem("���ݷ�Χ��")
    
    Set lcolItem = New Collection
    lcolItem.Add "ϵͳ��ŷ�Χ", "���ݷ�Χ��"
    lcolItem.Add "0103280204,0103280304", "���ݷ�Χֵ"
    lcolRange.Add lcolItem, lcolItem("���ݷ�Χ��")
    Set lobjRec = mobj������ӿ�.Func�鿴����(lcolRange, llngType)
    
    Set mobj������ӿ� = Nothing
    
    gfsubLoadGridFromRec cgrdMain, lobjRec
    
End Sub

Private Sub Command9_Click()
    Dim lcolRange As Collection
    Dim lobjRec As Object
    Set lobjRec = mobj������ӿ�.Func�鿴����(lcolRange, 0)
    
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    
    '��ʼ�����ݷ��ʶ���(����testserver)��
    dasubInitialize "Provider=MSDataShape.1;Data Provider=SQLOLEDB.1;Password=welcome;Persist Security Info=True;User ID=user26;Initial Catalog=����2001;Data Source=ychun"
    
    Set mobj������ӿ� = New ClsManageTransmission
    
    If Not umfuncУ�����("7612", "") Then
        sffuncMsg "У�����ʧ�ܡ�", sf����
    End If
    Exit Sub
errHandler:
    sfsub������ "����", "", "Form_load", Err.Number, Err.Description
    Exit Sub
    Resume
End Sub



Private Sub Command1_Click()
    Dim lobj����վ���� As Object
    On Error GoTo errHandler
    
    Text1.Text = "ҵ�����ã��Ƿ��ӡ��쵥=" & mobj������ӿ�.ҵ������("�Ƿ��ӡ��쵥") & Chr(13) & Chr(10)
    Set lobj����վ���� = mobj������ӿ�.����վ����
    Text1.Text = Text1.Text & "�������ã��ڲ������ļ�=" & lobj����վ����.�ڲ������ļ�
    
    Exit Sub
errHandler:
    sfsub������ "����", "", "Command1_Click", Err.Number, Err.Description
    Exit Sub
    Resume
End Sub

Private Sub Command2_Click()
    Dim a As New Collection
    Dim b As New Collection
    Dim e As New Collection
    Dim lstrTemp As String
    On Error GoTo errHandler
    Set a = mobj������ӿ�.Func��ȡ����ϵ������Ϣ("", "", "", "", , "��ҵ��Ա����", "����֤", "")
    For Each b In a
        lstrTemp = lstrTemp & "ϵͳ���:" & b("ϵͳ���") & ", ����:" & b("����") & ", ��λ����:" & b("��λ����") & vbCrLf
        For Each e In b("������Ŀ")
            lstrTemp = lstrTemp & "��Ŀ��" & e("��Ŀ��") & ",  "
        Next
        lstrTemp = lstrTemp & vbCrLf & vbCrLf
    Next
    Text1.Text = lstrTemp

    Exit Sub
errHandler:
    sfsub������ "����", "", "Command1_Click", Err.Number, Err.Description
    Exit Sub
    Resume
End Sub

Private Sub Command3_Click()
    On Error GoTo errHandler
    mobj������ӿ�.Sub���������ʱ�־ "111110103270106", "����֤1", "1"
    Exit Sub
errHandler:
    sfsub������ "", "", "", Err.Number, Err.Description, False
End Sub

Private Sub Command4_Click()
    Dim c As String
    Dim i As Integer
    Dim b As New Collection
    Dim a As New Collection
    
    On Error GoTo errHandler
    
    Set a = mobj������ӿ�.Func�������ļ��������Ա��Ϣ(App.Path & "\book1.xls")
    For Each b In a
        c = c & "��λ����:" & b("��λ����") & "                 "
        c = c & "��    ��:" & b("����") & vbCrLf
        c = c & "��    ��:" & b("����") & "                   "
        c = c & "��    ��:" & b("�Ա�") & vbCrLf
    Next b
    Text1.Text = c
    
    Exit Sub
errHandler:
    sfsub������ "����", "", "Command4_Click", Err.Number, Err.Description
    Exit Sub
    Resume
    
End Sub

Private Sub Command5_Click()
    Dim a As New Collection
    On Error GoTo errHandler
    
    Set a = mobj������ӿ�.Func�������ļ��������Ա��Ϣ(App.Path & "\book1.xls")
    mobj������ӿ�.Sub���������Ա�Ǽ� "������Ա���", ProgressBar1, a
    
    Exit Sub
errHandler:
    sfsub������ "����", "", "Command5_Click", Err.Number, Err.Description
    Exit Sub
    Resume
End Sub

Private Sub Command6_Click()
    On Error GoTo errHandler
    mobj������ӿ�.sub�����ļ��� App.Path & "\Form1.log"
    Exit Sub
errHandler:
    sfsub������ "����", "", "Command6_Click", Err.Number, Err.Description
    Exit Sub
    Resume
End Sub

Private Sub Command7_Click()
    Dim a As New Collection
    Dim b As New Collection
    
    On Error GoTo errHandler
    
    b.Add "��Դ,��˾", "��λ���Ƽ�"
    b.Add "00000103300001,00000103300099", "ϵͳ��ŷ�Χ"

    Exit Sub
errHandler:
    sfsub������ "����", "", "Command7_Click", Err.Number, Err.Description
    Exit Sub
    Resume
    
End Sub


