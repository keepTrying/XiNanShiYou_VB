VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmshenghua 
   Caption         =   "�������¼��"
   ClientHeight    =   5340
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8940
   LinkTopic       =   "Form1"
   ScaleHeight     =   5340
   ScaleWidth      =   8940
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox TextPassword 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   15
      Top             =   3600
      Width           =   2655
   End
   Begin VB.TextBox Textname 
      Height          =   375
      Left            =   1560
      TabIndex        =   14
      Top             =   3000
      Width           =   2655
   End
   Begin VB.CheckBox Check�Ƿ� 
      Caption         =   "�Ƿ���ķ�����"
      Height          =   255
      Left            =   6960
      TabIndex        =   11
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton command���� 
      Caption         =   "����"
      Height          =   495
      Left            =   7560
      TabIndex        =   8
      Top             =   4800
      Width           =   1215
   End
   Begin VB.TextBox TextBase 
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   2400
      Width           =   2655
   End
   Begin VB.TextBox TextIP 
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   1800
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�˳�"
      Height          =   615
      Left            =   5280
      TabIndex        =   5
      Top             =   2520
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ȡ"
      Height          =   615
      Left            =   5280
      TabIndex        =   0
      Top             =   960
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker DTP��ʼ 
      Height          =   420
      Left            =   1080
      TabIndex        =   3
      Top             =   120
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   741
      _Version        =   393216
      CustomFormat    =   "yyyy��MM��dd��"
      Format          =   59965443
      CurrentDate     =   42370
   End
   Begin MSComCtl2.DTPicker DTP��ֹ 
      Height          =   420
      Left            =   1080
      TabIndex        =   4
      Top             =   840
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   741
      _Version        =   393216
      CustomFormat    =   "yyyy��MM��dd��"
      Format          =   59965443
      CurrentDate     =   42370
   End
   Begin VB.Label Label6 
      Caption         =   "����"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "��¼��"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Ŀ�����ݿ���"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Ŀ�������IP"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "����ʱ��"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "��ʼʱ��"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmshenghua"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'�������ݿ���ַ���
'Private Const Conn As String = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;PWD=123456;Initial Catalog=KHB2LIS;Data Source=192.168.0.155"
'hr�����ݿ�����  Catalong=���ݿ�����
Private sqlreuslt As String
Private res As Object
Private IsConnect As Boolean             '�ж����ݿ��Ƿ�����
Private cnn As ADODB.Connection       ' �������ݿ��connection����
Private rs As ADODB.Recordset           '����������recordset����
 
Private Sub Check�Ƿ�_Click()
If Check�Ƿ�.Value = 1 Then
TextIP.Enabled = True
TextBase.Enabled = True
Textname.Enabled = True
TextPassword.Enabled = True
Else
TextIP.Enabled = False
TextBase.Enabled = False
Textname.Enabled = False
TextPassword.Enabled = False
End If
End Sub

'��ѯsql2005���ݱ�������ݲ�����sql2000�������ݱ���  2016-1-21 by Ĳ��
Private Sub Command1_Click()
    Dim xmID As String
    Dim lobject As Object
    Dim ID As String
    Dim lob As Object
     Dim dtpTimeTo As Date
    Dim sql As String
    '���ݽ�������������µ����������Ϣ�����뵽���½����ı��У�  2016-2-1 by Ĳ����
    Dim SysNo As String
    command����_Click
    '�ж��Ƿ����ӳɹ���������ɹ�ֱ���˳�
    If cnn.State <> adStateOpen Then
    Exit Sub
    End If
    sql = " (TestTime between '" & Format(DTP��ʼ.Value, "yyyy-mm-dd" & " 00:00:00") & "' and '" & Format(DTP��ֹ.Value, "yyyy-mm-dd" & " 23:59:59") & "')"
    Set lob = cnn.Execute("select BarCode from labLISItemResult  where Barcode<>'' and  " & sql & " group by BarCode")
    If lob.RecordCount > 0 Then
    Dim k As Integer
    For k = 1 To lob.RecordCount
    SysNo = lob("BarCode")
    Set res = cnn.Execute("select * from labLISItemResult where Barcode='" & SysNo & "' and " & sql & "")
         dafuncGetData ("delete from ְҵ�����_�����Ϣ_�����Ƶ�����Ϣ�� where BarCode='" & SysNo & "'and " & sql & "") 'ɾ��ԭ�м�¼����������Ϊ׼
        If res.RecordCount > 0 Then
        Dim i As Long
        For i = 1 To res.RecordCount
        dafuncGetData ("insert into ְҵ�����_�����Ϣ_�����Ƶ�����Ϣ�� values('" & res("SeqNo") & "','" & res("SampleID") & "','" & res("BarCode") & "','" & res("ItemCode") & "','" & res("TestValue") & "','" & res("TestTime") & "','" & res("SampleNo") & "','" & res("Maked") & "')")
        res.MoveNext
        Next i
        End If
    lob.MoveNext
    Next k
    MsgBox "�����Ѷ�ȡ���"
    Unload Me
    Else
    MsgBox "���ݲ����ڣ���ȷ���Ƿ��͹�LIS"
    End If
    '2016-2-1 by Ĳ�� ��
End Sub


Private Sub Command2_Click()
Unload Me
End Sub


Private Sub command����_Click()
On Error GoTo erhander
Set cnn = New Connection
Dim strConnection As String
Dim sever As String
Dim base As String
Dim pawd As String
Dim username As String
sever = TextIP.Text
base = TextBase.Text
pawd = TextPassword.Text
username = Textname.Text
If cnn.State = adStateOpen Then cnn.Close
'strConnection = "Provider=SQLOLEDB.1;Password=123456;Persist Security Info=True;User ID=sa;Initial Catalog=" & base & ";Data Source=" & sever & ""
strConnection = "Provider=SQLOLEDB.1;Password=" & pawd & ";Persist Security Info=True;User ID=" & username & ";Initial Catalog=" & base & ";Data Source=" & sever & ""
cnn.ConnectionString = strConnection
cnn.CursorLocation = adUseClient
cnn.Open strConnection
Exit Sub
erhander:
Select Case Err.Number
    Case -2147467259
    MsgBox "�Ҳ���ָ��sql���ݿ�", vbExclamation, "���ݿ����"
    Case -2147217843
    MsgBox "ָ����sql server���ݿ��û������ڻ�������", vbExclamation, "���ݿ����"
    Case Else
    MsgBox "���ݻ�������ʧ�ܣ�����ϵͳ����Ա���м��", vbExclamation, "���ݿ����"
End Select
End Sub

Private Sub Form_Load()
TextIP.Enabled = False
TextBase.Enabled = False
Textname.Enabled = False
TextPassword.Enabled = False
command����.Visible = False
'Ĭ��һ��������IP�����ݿ�  2016-1-29 by Ĳ��
TextIP.Text = "192.168.0.164"
'TextBase.Text = "LabConsole"
TextBase.Text = "KHB2LIS"
Textname.Text = "sa"
TextPassword.Text = "123456"
End Sub


