VERSION 5.00
Object = "{2327075A-FED2-4AEF-82B8-DD0C2B1AC8E1}#1.5#0"; "dyCatchPhoto.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frm���¼�� 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "����֤¼��"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   11070
   ClipControls    =   0   'False
   Icon            =   "frm���¼��.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   11070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.ComboBox ccmb��ҵ��� 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   3120
      TabIndex        =   6
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CommandButton ccmdClear 
      Caption         =   "���(&C)"
      Height          =   375
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton ccmdPrint 
      Caption         =   "��ӡ֤(&P)"
      Height          =   375
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CheckBox cchkClear 
      Caption         =   "������Զ��Զ����"
      Height          =   255
      Left            =   120
      TabIndex        =   43
      Top             =   7320
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.TextBox ctxt���� 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   1200
      MaxLength       =   10
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.CommandButton ccmd��λ 
      Caption         =   "..."
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
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "��λ��λ"
      Top             =   1440
      Width           =   495
   End
   Begin VB.ComboBox ccmbMZ 
      Height          =   300
      Left            =   1200
      TabIndex        =   38
      Top             =   2880
      Width           =   3255
   End
   Begin MSComDlg.CommonDialog ccmdFile 
      Left            =   4800
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSMask.MaskEdBox ctxt������� 
      Height          =   300
      Left            =   6600
      TabIndex        =   13
      Top             =   6000
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      _Version        =   393216
      BackColor       =   12648447
      MaxLength       =   10
      Format          =   "yyyy-mm-dd"
      Mask            =   "####-##-##"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton ccmdSaveAs 
      Caption         =   "�����Ƭ(&A)"
      Height          =   375
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton ccmdLoad 
      Caption         =   "������Ƭ(&L)"
      Height          =   375
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   4560
      Width           =   1215
   End
   Begin VB.ComboBox ccmb���� 
      BackColor       =   &H00C0FFFF&
      Height          =   300
      ItemData        =   "frm���¼��.frx":0E42
      Left            =   1200
      List            =   "frm���¼��.frx":0E4C
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   6000
      Width           =   3255
   End
   Begin VB.ComboBox ccmb��֤��λ 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      ItemData        =   "frm���¼��.frx":0E60
      Left            =   1200
      List            =   "frm���¼��.frx":0E6A
      TabIndex        =   12
      Top             =   6480
      Width           =   3255
   End
   Begin VB.ComboBox ccmb��ѵ���� 
      BackColor       =   &H00C0FFFF&
      Height          =   300
      ItemData        =   "frm���¼��.frx":0E7C
      Left            =   1200
      List            =   "frm���¼��.frx":0E86
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   5520
      Width           =   3255
   End
   Begin VB.Frame Frame2 
      Caption         =   "����"
      Height          =   4215
      Left            =   5520
      TabIndex        =   28
      Top             =   240
      Width           =   4935
      Begin dyCatchPhoto.ctlCatchPhoto ctlCatchPhoto 
         Height          =   3615
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   6376
         BackColor       =   0
         FontSize        =   11.25
      End
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   120
      TabIndex        =   27
      Top             =   6840
      Width           =   10815
   End
   Begin VB.CommandButton ccmdExit 
      Caption         =   "����(&C)"
      Height          =   375
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton ccmdSave 
      Caption         =   "����(&S)"
      Height          =   375
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7200
      Width           =   1215
   End
   Begin VB.ComboBox ccmb������ 
      BackColor       =   &H00C0FFFF&
      Height          =   300
      ItemData        =   "frm���¼��.frx":0E98
      Left            =   1200
      List            =   "frm���¼��.frx":0EA2
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   5040
      Width           =   3255
   End
   Begin VB.ListBox clstDisease 
      Height          =   1530
      ItemData        =   "frm���¼��.frx":0EB4
      Left            =   1200
      List            =   "frm���¼��.frx":0EC1
      Style           =   1  'Checkbox
      TabIndex        =   8
      Top             =   3360
      Width           =   3255
   End
   Begin VB.ComboBox ccmbOcc 
      Height          =   300
      Left            =   1200
      TabIndex        =   7
      Top             =   2400
      Width           =   3255
   End
   Begin VB.ComboBox ccmbType 
      BackColor       =   &H00C0FFFF&
      Height          =   300
      Left            =   1200
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox ctxtUnit 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   1200
      MaxLength       =   40
      TabIndex        =   4
      Top             =   1440
      Width           =   3255
   End
   Begin VB.TextBox ctxtAge 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   3600
      MaxLength       =   10
      TabIndex        =   3
      Top             =   960
      Width           =   855
   End
   Begin VB.ComboBox ccmbSex 
      BackColor       =   &H00C0FFFF&
      Height          =   300
      Left            =   1200
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox ctxtName 
      BackColor       =   &H00C0FFFF&
      Height          =   300
      Left            =   1200
      MaxLength       =   10
      TabIndex        =   1
      Top             =   480
      Width           =   3255
   End
   Begin MSMask.MaskEdBox ctxt��֤���� 
      Height          =   300
      Left            =   6600
      TabIndex        =   14
      Top             =   6480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   10
      Format          =   "yyyy-mm-dd"
      Mask            =   "####-##-##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox ctxt��Ч���� 
      Height          =   300
      Left            =   9240
      TabIndex        =   16
      Top             =   6480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   10
      Format          =   "yyyy-mm-dd"
      Mask            =   "####-##-##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox ctxt��ѵ���� 
      Height          =   300
      Left            =   9240
      TabIndex        =   15
      Top             =   6000
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      _Version        =   393216
      BackColor       =   12648447
      MaxLength       =   10
      Format          =   "yyyy-mm-dd"
      Mask            =   "####-##-##"
      PromptChar      =   "_"
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ѵ���ڣ�"
      Height          =   180
      Index           =   16
      Left            =   8280
      TabIndex        =   42
      Top             =   6120
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���ţ�"
      Height          =   180
      Index           =   15
      Left            =   120
      TabIndex        =   41
      Top             =   240
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��    �壺"
      Height          =   180
      Index           =   14
      Left            =   120
      TabIndex        =   39
      Top             =   3000
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��    �ã�"
      Height          =   180
      Index           =   13
      Left            =   120
      TabIndex        =   35
      Top             =   6120
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��Ч������"
      Height          =   180
      Index           =   12
      Left            =   8280
      TabIndex        =   34
      Top             =   6600
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��֤���ڣ�"
      Height          =   180
      Index           =   11
      Left            =   5640
      TabIndex        =   33
      Top             =   6600
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������ڣ�"
      Height          =   180
      Index           =   5
      Left            =   5640
      TabIndex        =   32
      Top             =   6120
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��쵥λ��"
      Height          =   180
      Index           =   10
      Left            =   120
      TabIndex        =   31
      Top             =   6600
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ѵ���ۣ�"
      Height          =   180
      Index           =   9
      Left            =   120
      TabIndex        =   30
      Top             =   5640
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����ۣ�"
      Height          =   180
      Index           =   8
      Left            =   120
      TabIndex        =   26
      Top             =   5160
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������֣�"
      Height          =   180
      Index           =   7
      Left            =   120
      TabIndex        =   25
      Top             =   3360
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ְ    ҵ��"
      Height          =   180
      Index           =   6
      Left            =   120
      TabIndex        =   24
      Top             =   2520
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��    �ࣺ"
      Height          =   180
      Index           =   4
      Left            =   120
      TabIndex        =   23
      Top             =   2040
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��λ���ƣ�"
      Height          =   180
      Index           =   3
      Left            =   120
      TabIndex        =   22
      Top             =   1560
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��    �䣺"
      Height          =   180
      Index           =   2
      Left            =   2520
      TabIndex        =   21
      Top             =   1080
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��    ��"
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   20
      Top             =   1080
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��    ����"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   19
      Top             =   600
      Width           =   900
   End
   Begin VB.Menu cmnuPop 
      Caption         =   "cmnuPop"
      Visible         =   0   'False
      Begin VB.Menu cmnuItemPop 
         Caption         =   "���"
         Index           =   1
      End
      Begin VB.Menu cmnuItemPop 
         Caption         =   "ɾ��"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frm���¼��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public pstrϵͳ��� As String

Private Sub ccmbType_Click()
    Dim lobjRec As Object
    On Error GoTo errhandler
    
    ccmb��ҵ���.Clear
    Set lobjRec = dafuncGetData("select * from ϵͳ����_��ҵ����ֵ���ͼ where Parent=(select InnerID from ϵͳ����_���������ֵ���ͼ where ����='" & ccmbType.Text & "'" & IIf(ccmbType.Text = "��������", " or ����='������������'", "") & ")")
    Do While Not lobjRec.EOF
        ccmb��ҵ���.AddItem lobjRec!����
        lobjRec.movenext
    Loop
    
    Exit Sub
errhandler:
    sfsub������ "����֤������", "frm���¼��", "ccmbType_Click", Err.Number, Err.Description, False
    
End Sub

Private Sub ccmb����_Click()
    On Error Resume Next
    If ccmb����.ListIndex = 0 Then
        ctxt��֤����.Text = Format(Now, "yyyy-mm-dd")
        ctxt��Ч����.Text = Format(DateAdd("y", 1, ctxt��֤����.Text), "yyyy-mm-dd")
    Else
        ctxt��֤����.Text = "____-__-__"
        ctxt��Ч����.Text = "____-__-__"
    End If
End Sub

Private Sub ccmb������_Click()
    On Error Resume Next
    If ccmb������.Text = "�ϸ�" Then
        ccmb����.ListIndex = 0
        ctxt��֤����.Text = Format(Now, "yyyy-mm-dd")
        ctxt��Ч����.Text = Format(DateAdd("y", 1, ctxt��֤����.Text), "yyyy-mm-dd")
        
    Else
        ccmb����.ListIndex = 1
        ctxt��֤����.Text = "____-__-__"
        ctxt��Ч����.Text = "____-__-__"
    End If
    
End Sub

Private Sub ccmdClear_Click()
    ctxt����.Text = ""
    ctxtName.Text = ""
    pstrϵͳ��� = ""
    ctxtAge = ""
End Sub

Private Sub ccmdExit_Click()
    On Error Resume Next
    pobj����.sub���Ǽ���ֵ "����֤¼�뱣������", cchkClear.Value
    
    Unload Me
    Call frm����֤����.subRefresh
End Sub

Private Sub ccmdLoad_Click()
    Dim lstrFile As String
    On Error GoTo errhandler
    
    ccmdFile.Filter = "BMP|*.bmp|JPG|*.jpg"
    If Dir(App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "��Ƭ", vbDirectory) <> "" Then
        ccmdFile.InitDir = App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "Photo"
    End If
    ccmdFile.FileName = pstrϵͳ���
    ccmdFile.ShowOpen
    lstrFile = ccmdFile.FileName
    If lstrFile <> "" Then
        If InStr(lstrFile, ".") > 0 Then
            Set ctlCatchPhoto.Photo = LoadPicture(lstrFile)
        End If
    End If
    Exit Sub
errhandler:
    MsgBox "������Ƭʧ�ܣ�" & Error, vbOKOnly + vbExclamation, "ϵͳ��ʾ"
End Sub

Private Sub ccmdPrint_Click()
    Dim lstrCN As String
    On Error GoTo errhandler
    Dim lstrCode As String
    
    '����ҵ�����ã��ж��Ƿ���Ҫ�Զ����ɽ���֤�š�
    lstrCode = pobj������.ҵ������("����֤������")
    lstrCode = "��" 'ʡ��ǿ��Ҫ��ʹ�ô������֤��
    If lstrCode = "��" Or pobj������.ҵ������("�ֹ����뽡��֤��") = "��" Then
    
        '�û����뽡��֤�ŵ���ʼ�š�
        lstrCN = InputBox("�����뽡��֤��", "����")
        If lstrCN = "" Then
            Exit Sub
        End If
        
        '�ж����뽡��֤���Ƿ�Ϊ���֡�
        If lstrCode = "��" Then
            Do While Not (IsNumeric(lstrCN))
                If MsgBox("������Ľ���֤�Ÿ�ʽ���ԡ��Ƿ��������룿", vbYesNo, "ϵͳ��ʾ") = vbYes Then
                    lstrCN = InputBox("�����뽡��֤����ʼ��", "����")
                Else
                    Exit Sub
                End If
            Loop
            
            '�жϿ��Ƿ�Ϸ���
            Dim lobjEncrypt As Object
            Set lobjEncrypt = CreateObject("fycarddes.clsDataEncrypt")
            If Not lobjEncrypt.funcCheckJkzCardno(lstrCN) Then
                Err.Raise 6666, , "ϵͳ�޷�ʶ�����ſ�����ȷ��������ָ���ĸ�ʽ���Ƿ����𻵣�"
            End If
            '������У��λ��
            lstrCN = lobjEncrypt.����
            Set lobjEncrypt = Nothing
        End If
    Else
        'ϵͳ�Զ����ɽ���֤��
        lstrCN = ""
    End If
    
    Dim lcolInfo As New Collection
    Dim lobj��� As cls���
    Set lobj��� = New cls���
    lobj���.ϵͳ��� = pstrϵͳ���
    
    If Not (lstrCode = "��" Or pobj������.ҵ������("�ֹ����뽡��֤��") = "��") Then
        '��Ҫϵͳ�Զ��������֤�š�
        If lobj���.״̬ = "δ��ӡ" And lobj���.����֤�� = "" Then
            lstrCN = pobj������.func���ɽ���֤��()
            lobj���.����֤�� = lstrCN
        End If
    Else
        lobj���.����֤�� = lstrCN
    End If
    
    '��������ϵͳ�������ļ�¼��û�з�֤���ںͷ�֤��λ��
    If lobj���.��֤���� = "" Then
        lobj���.��֤���� = Format(Date, "yyyy-mm-dd")
    End If
    If lobj���.��Ч���� = "" Then
        lobj���.��Ч���� = Format(DateAdd("d", -1, DateAdd("yyyy", 1, Date)), "yyyy-mm-dd")
    End If
    If lobj���.��֤��λ = "" Then
        lobj���.��֤��λ = um����վ��
    End If
                    
    lcolInfo.Add lobj���
       
    pobj������.sub��ӡ����֤ lcolInfo
    
    Exit Sub
errhandler:
    sfsub������ "����֤������", "ccmdPrint_Click", Err.Number, Err.Description, False
    
End Sub

Private Sub Ccmdsave_Click()
    Dim lobj��� As New cls���
    Dim lstr������� As String
    Dim i As Long
    On Error GoTo errhandler
    
    '�������Ϸ��ԡ�
    If ctxtName.Text = "" Then
        MsgBox "��������������������ɫ��Ҳ����¼�룡", vbOKOnly + vbExclamation, "ϵͳ��ʾ"
        ctxtName.SetFocus
        Exit Sub
    End If
    If Len(ccmbOcc.Text) > 10 Then
        MsgBox "ְҵ�������10�����֣�", vbOKOnly + vbExclamation, "ϵͳ��ʾ"
        ccmbOcc.SetFocus
        Exit Sub
    End If
    If Not IsDate(ctxt�������.Text) Then
        MsgBox "��������������ڣ�������ɫ��Ҳ����¼�룡", vbOKOnly + vbExclamation, "ϵͳ��ʾ"
        ctxt�������.SetFocus
        Exit Sub
    End If
    If Not IsDate(ctxt��ѵ����.Text) Then
        MsgBox "����������ѵ���ڣ�������ɫ��Ҳ����¼�룡", vbOKOnly + vbExclamation, "ϵͳ��ʾ"
        ctxt�������.SetFocus
        Exit Sub
    End If
    If ccmb����.ListIndex = 0 Then
        If Not IsDate(ctxt��֤����.Text) Or Not IsDate(ctxt��Ч����.Text) Then
            MsgBox "�������뷢֤���ں���Ч������������ɫ��Ҳ����¼�룡", vbOKOnly + vbExclamation, "ϵͳ��ʾ"
            ccmb��֤��λ.SetFocus
            Exit Sub
        End If
        If Len(ccmb��֤��λ.Text) > 30 Then
            MsgBox "��֤��λ�������30�����֣�", vbOKOnly + vbExclamation, "ϵͳ��ʾ"
            ccmbOcc.SetFocus
            Exit Sub
        End If
    End If
    
    '��ȡ������ء�
    lstr������� = ""
    For i = 0 To clstDisease.ListCount - 1
        If clstDisease.Selected(i) Then
            lstr������� = lstr������� & clstDisease.List(i) & ","
        End If
    Next
    If lstr������� <> "" Then lstr������� = Left(lstr�������, Len(lstr�������) - 1)
    
    '���档
    With lobj���
        .ϵͳ��� = pstrϵͳ���
        .���� = ctxt����.Text
        .���� = ctxtName.Text
        .�Ա� = ccmbSex.Text
        .���� = ctxtAge.Text
        .������ = ctxtUnit.Tag
        .��λ���� = ctxtUnit.Text
        .���� = ccmbType.Text
        .ְҵ = ccmbOcc.Text
        .���� = ccmbMZ.Text
        .������ = ccmb������.Text
        .��ѵ���� = ccmb��ѵ����.Text
        .������� = ctxt�������.FormattedText
        .��ѵ���� = ctxt��ѵ����.FormattedText
        .���� = ccmb����.Text
        
        If ctxt��֤����.FormattedText = "____-__-__" Then
            .��֤���� = ""
        Else
            .��֤���� = ctxt��֤����.FormattedText
        End If
        If ctxt��Ч����.FormattedText = "____-__-__" Then
            .��Ч���� = ""
        Else
            .��Ч���� = ctxt��Ч����.FormattedText
        End If
        .������� = lstr�������
        .��֤��λ = ccmb��֤��λ.Text
        
        If pobj������.ҵ������("�Ƿ�����") = "��" Then
            Set .��Ƭ = ctlCatchPhoto.Photo
        End If
        
        .sub����
    End With
    
    On Error Resume Next
    
    '����ö��ֵ��
    pobj����.sub��Ӽ���ֵ "��������", Trim(ccmbType.Text)
    pobj����.sub��Ӽ���ֵ "ְҵ", Trim(ccmbOcc.Text)
    pobj����.sub��Ӽ���ֵ "����", Trim(ccmbMZ.Text)
    If Trim(ccmb��֤��λ.Text) <> "" Then
        pobj����.sub��Ӽ���ֵ "��֤��λ", Trim(ccmb��֤��λ.Text)
    End If
    
    
    If pstrϵͳ��� = "" And cchkClear.Value = 1 Then
        '��ձ�¼�
        ctxtName.Text = ""
        ctxtAge.Text = ""
        ctxt����.Text = ""
        ctxt����.SetFocus
        
        '�ָ����ࡣ
        If pobj������.ҵ������("�Ƿ�����") = "��" Then
            If ctlCatchPhoto.Status = "�ָ�" Then
                ctlCatchPhoto.subת��״̬
            End If
        End If
    Else
        pstrϵͳ��� = lobj���.ϵͳ���
        ccmdPrint.Enabled = True
    End If

    
    Exit Sub
errhandler:
    sfsub������ "����֤������", "frm���¼��", "Ccmdsave_Click", Err.Number, Err.Description, False
End Sub

Private Sub ccmdSaveAs_Click()
    Dim lstrFile As String
    On Error GoTo errhandler
    
    ccmdFile.Filter = "BMP|*.bmp|JPG|*.jpg"
    If Dir(App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "��Ƭ", vbDirectory) <> "" Then
        ccmdFile.InitDir = App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "Photo"
    End If
    ccmdFile.FileName = pstrϵͳ���
    ccmdFile.ShowOpen
    lstrFile = ccmdFile.FileName
    If lstrFile <> "" Then
        SavePicture ctlCatchPhoto.Photo, lstrFile
    End If

    Exit Sub
errhandler:
    MsgBox "�����Ƭʧ�ܣ�" & Error, vbOKOnly + vbExclamation, "ϵͳ��ʾ"
End Sub

Private Sub ccmd��λ_Click()
    Dim lstrUnitName As String
    Dim lrstTmp As Object
    Dim lstrUnitNumber As String
    Dim lobj�ӿ� As Object
    
    On Error GoTo errhandler
    
    Set lobj�ӿ� = CreateObject("��λ����ҵ��.ClsUnitInterface")
    Set lrstTmp = lobj�ӿ�.func��λ�򵥶�λ(Screen.Width / 2, Screen.Height / 2)
    
    If lrstTmp Is Nothing Then
        ctxtUnit.SetFocus
        Exit Sub
    End If
    
    lstrUnitNumber = lrstTmp!������
    ctxtUnit = lrstTmp!��λ����
    ccmbType.Text = lrstTmp!��������
    ccmbType_Click
    
    ccmb��ҵ���.Text = IIf(IsNull(lrstTmp!��ҵ���), "", lrstTmp!��ҵ���)
    
    Me.ctxtUnit.Tag = lrstTmp!������ '�ڵ�λ���Ƶ�tag�м�¼��λ�����š�
    ctxtUnit.SetFocus
    Exit Sub
errhandler:
    sfsub������ "����֤������", "frm���¼��", "ccmd��λ_Click", Err.Number, Err.Description, False
End Sub

Private Sub clstDisease_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = vbRightButton Then
        If clstDisease.ListIndex < 0 Then
            cmnuItemPop(2).Enabled = False
        Else
            cmnuItemPop(2).Enabled = True
        End If
        PopupMenu cmnuPop
    End If
End Sub

Private Sub cmnuItemPop_Click(Index As Integer)
    Dim lstrItem As String
    Dim i As Long
    
    On Error Resume Next
    Select Case Index
    Case 1 '����
        '�����������֡�
        lstrItem = Trim(InputBox("�������������֣�", "ϵͳѯ��", ""))
        If lstrItem <> "" Then
            If InStr(lstrItem, "'") > 0 Then
                MsgBox "���������в��ܺ��зǷ��ַ���'�����������ţ���", vbOKOnly + vbExclamation, "ϵͳ��ʾ"
                Exit Sub
            End If
            i = 0
            For i = 0 To clstDisease.ListCount - 1
                If clstDisease.List(i) = lstrItem Then
                    Exit For
                End If
            Next
            If i = clstDisease.ListCount Then
                clstDisease.AddItem lstrItem
            End If
        End If
    Case 2 'ɾ����
        If clstDisease.ListIndex >= 0 Then
            If clstDisease.List(clstDisease.ListIndex) = "��" Then
                MsgBox "����Ŀ����ɾ����", vbOKOnly + vbExclamation, "ϵͳ��ʾ"
                Exit Sub
            End If
            clstDisease.RemoveItem clstDisease.ListIndex
        End If
    End Select
    
End Sub

Private Sub ctxt��֤����_Change()
    On Error Resume Next
    If IsDate(ctxt��֤����.Text) Then
        ctxt��Ч����.Text = Format(DateAdd("d", -1, DateAdd("yyyy", 1, ctxt��֤����.Text)), "yyyy-mm-dd")
    End If
End Sub

'���ܣ����Ʋ������뵥ӡ�ţ�����س���
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        SendKeys Chr(9)
    ElseIf KeyCode = 39 Then
        KeyCode = 0
    End If
    

End Sub



Private Sub Form_Load()
    Dim lcolInfo As Collection
    Dim i As Long
    
    On Error GoTo errhandler
    
    ccmbSex.Clear
    ccmbSex.AddItem "��"
    ccmbSex.AddItem "Ů"
    ccmbSex.ListIndex = 0
    
    '��ȡ���ࡣ
    Set lcolInfo = pobj����.������ֵ("��������", True)
    ccmbType.Clear
    For i = 1 To lcolInfo.Count
        ccmbType.AddItem lcolInfo(i)
    Next
    If ccmbType.ListCount > 0 Then ccmbType.ListIndex = 0
    
    '��ȡְҵ��
    Set lcolInfo = pobj����.������ֵ("ְҵ", True)
    ccmbOcc.Clear
    For i = 1 To lcolInfo.Count
        ccmbOcc.AddItem lcolInfo(i)
    Next
    If ccmbOcc.ListCount > 0 Then ccmbOcc.ListIndex = 0
    
    
    '��ȡ���塣
    Set lcolInfo = pobj����.������ֵ("����", True)
    ccmbMZ.Clear
    For i = 1 To lcolInfo.Count
        ccmbMZ.AddItem lcolInfo(i)
    Next
    If ccmbMZ.ListCount > 0 Then ccmbMZ.ListIndex = 0
    
    '��ȡ������֡�
    Set lcolInfo = pobj����.������ֵ("�������", True)
    clstDisease.Clear
    clstDisease.AddItem "�޴�ҵ����֢"
    For i = 1 To lcolInfo.Count
        clstDisease.AddItem lcolInfo(i)
    Next
    clstDisease.Selected(0) = True
    
    '��ȡ��֤��λ��
    Set lcolInfo = pobj����.������ֵ("��֤��λ", True)
    ccmb��֤��λ.Clear
    ccmb��֤��λ.AddItem ""
    For i = 1 To lcolInfo.Count
        ccmb��֤��λ.AddItem lcolInfo(i)
    Next
    If ccmb��֤��λ.ListCount > 1 Then
        ccmb��֤��λ.ListIndex = 1
    Else
        ccmb��֤��λ.ListIndex = 0
    End If
    
    ccmb������.ListIndex = 0
    ccmb��ѵ����.ListIndex = 0
    ccmb����.ListIndex = 0
    
    ctxt�������.Text = Format(Date, "yyyy-mm-dd")
    ctxt��֤����.Text = Format(Date, "yyyy-mm-dd")
    ctxt��ѵ����.Text = Format(Date, "yyyy-mm-dd")
    ctxt��Ч����.Text = Format(DateAdd("d", -1, DateAdd("yyyy", 1, Date)), "yyyy-mm-dd")
    
    '��ʾ�޸���Ա��Ϣ��
    Dim lobj��� As cls���
    Dim lstr������� As String
    If pstrϵͳ��� <> "" Then
        Set lobj��� = New cls���
        lobj���.ϵͳ��� = pstrϵͳ���
        
        ctxt����.Text = lobj���.����
        ctxtName.Text = lobj���.����
        If lobj���.�Ա� = "��" Then
            ccmbSex.ListIndex = 0
        Else
            ccmbSex.ListIndex = 1
        End If
        ctxtAge.Text = lobj���.����
        
        ccmbType.Text = lobj���.����
        ccmbOcc.Text = lobj���.ְҵ
        
        ctxtUnit.Text = lobj���.��λ����
        ctxtUnit.Tag = lobj���.������
        
        lstr������� = lobj���.�������
        If lstr������� <> "" Then
            lstr������� = lstr������� & ","
            For i = 0 To clstDisease.ListCount - 1
                If InStr(lstr�������, clstDisease.List(i) & ",") > 0 Then
                    clstDisease.Selected(i) = True
                End If
            Next
        End If
        
        If lobj���.������ = "�ϸ�" Then
            ccmb������.ListIndex = 0
        Else
            ccmb������.ListIndex = 1
        End If
    
        If lobj���.��ѵ���� = "�ϸ�" Then
            ccmb��ѵ����.ListIndex = 0
        Else
            ccmb��ѵ����.ListIndex = 1
        End If
        If lobj���.���� = "������֤" Then
            ccmb����.ListIndex = 0
        Else
            ccmb����.ListIndex = 1
        End If
        
        ctxt�������.Text = Format(lobj���.�������, "yyyy-mm-dd")
        
        ccmb��֤��λ.Text = lobj���.��֤��λ
        
        If lobj���.��֤���� = "" Then
            ctxt��֤����.Text = "____-__-__"
        Else
            ctxt��֤����.Text = Format(lobj���.��֤����, "yyyy-mm-dd")
        End If
        If lobj���.��Ч���� = "" Then
            ctxt��Ч����.Text = "____-__-__"
        Else
            ctxt��Ч����.Text = Format(lobj���.��Ч����, "yyyy-mm-dd")
        End If
        
        Set ctlCatchPhoto.Photo = lobj���.��Ƭ
        ccmdPrint.Enabled = True
    Else
        ccmdPrint.Enabled = False
    End If
    
    '��ʼ������ؼ���
    If pobj������.ҵ������("�Ƿ�����") = "��" Then
        Frame2.Caption = "����"
        ctlCatchPhoto.Enabled = True
        ctlCatchPhoto.funcInitVideo
        ccmdLoad.Enabled = True
        ccmdSaveAs.Enabled = True
    Else
        Frame2.Caption = "ҵ������Ϊ������"
        ctlCatchPhoto.Enabled = False
        ccmdLoad.Enabled = False
        ccmdSaveAs.Enabled = False
    End If
    If pobj����.������ֵ("����֤¼�뱣������") = "1" And pstrϵͳ��� = "" Then
        cchkClear.Value = 1
    Else
        cchkClear.Value = 0
    End If
    
    If Not umfuncУ���û�Ȩ��("����֤����_��ӡ") Then
        ccmdPrint.Visible = False
    End If
    
    Exit Sub
errhandler:
    sfsub������ "����֤������", "frm���¼��", "Form_Load", Err.Number, Err.Description, False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    ctlCatchPhoto.subDisconnect

    On Error GoTo errhandler
    Dim lstrList As String
    Dim i As Long
    
    For i = 1 To clstDisease.ListCount - 1
        lstrList = lstrList & clstDisease.List(i) & ","
    Next
    
    If lstrList <> "" Then lstrList = Left(lstrList, Len(lstrList) - 1)
    pobj����.sub���Ǽ���ֵ "�������", lstrList
    
errhandler:

End Sub


