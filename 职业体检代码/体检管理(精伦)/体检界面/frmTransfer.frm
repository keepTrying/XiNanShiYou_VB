VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTransfer 
   Caption         =   "�������ݵ�������"
   ClientHeight    =   3075
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3075
   ScaleWidth      =   4680
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton ccmdCancel 
      Caption         =   "��  ��"
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton ccmdOK 
      Caption         =   "ȷ  ��"
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   2400
      Width           =   855
   End
   Begin MSComctlLib.ProgressBar cprgStatus 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComCtl2.DTPicker cdtpDate 
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   52101121
      CurrentDate     =   39914
   End
   Begin VB.TextBox ctxtServerIP 
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "������ڣ�"
      Height          =   180
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "������IP��"
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   900
   End
End
Attribute VB_Name = "frmTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mobj����  As cls�û���������

Private Sub ccmdCancel_Click()
    Unload Me
End Sub

Private Sub ccmdOk_Click()
    If ctxtServerIP = "" Then
        MsgBox "���������������IP��", vbInformation, "ϵͳ��ʾ"
        ctxtServerIP.SetFocus
        Exit Sub
    End If
    
    If ctxtServerIP = sffuncGetSetting("ϵͳ����", "���ݿ�����", "��������") Then
        MsgBox "�������Լ��������ݣ�", vbInformation, "ϵͳ��ʾ"
        ctxtServerIP.SetFocus
        Exit Sub
    End If

    If MsgBox("ȷ��Ҫ�����ݴ��䵽����������", vbQuestion + vbYesNo, "ϵͳѯ��") = vbNo Then Exit Sub
    
    Dim lobjConn As New ADODB.Connection
    Dim lobjRec As Recordset, lstrSql As String
    Dim lstrDate As String, i As Integer
    
    lstrDate = Format(cdtpDate.Value, "yyyy-mm-dd")
    
    On Error GoTo errHandle
    
    cprgStatus.Visible = True
    
    lobjConn.ConnectionString = "Provider=MSDataShape.1;Data Provider=SQLOLEDB.1;Password=welcome58%*;Persist Security Info=True;User ID=jk_user;Initial Catalog=jk2006;Data Source=" & ctxtServerIP
    lobjConn.Open
    '�жϷ��������ݿ��еķ����������Ƿ��뱾����ͬ
    Set lobjRec = lobjConn.Execute("select ����վ���,���������� from ϵͳ����_ϵͳ�������ñ�")
    If lobjRec(0) = um����վ��� And lobjRec(1) = um���������� Then
        MsgBox "�������뱾���ı���λ��š�������������ȫ��ͬ�����ܴ������ݣ�", vbInformation, "ϵͳ��ʾ"
        ctxtServerIP.SetFocus
        Exit Sub
    End If
    
    '��ɾ������������Դ�ڱ����Եĸ����ڵ��������
    Dim lstrNoPre As String         '�����Եı��ǰ׺
    
    lstrNoPre = um����վ��� + um����������
'    lobjConn.Execute "delete ������_�����ʱ�־�� where ϵͳ��� in (select ϵͳ��� from ������_��������Ϣ�� where �������='" & lstrDate & "')"
'    lobjConn.Execute "delete ϵͳ����_ϵͳͼƬ����� where ��ϵͳ��='������' and ͼƬ��� in (select ϵͳ��� from ������_��������Ϣ�� where �������='" & lstrDate & "')"
'    lobjConn.Execute "delete ������_��������Ϣ�� where �������='" & lstrDate & "'"
'    lobjConn.Execute "delete ������_�����Ա������Ϣ�� where ��������='" & lstrDate & "'"
    lobjConn.Execute "delete ������_�����ʱ�־�� where ϵͳ��� in (select ϵͳ��� from ������_��������Ϣ�� where �������='" & lstrDate & "' and ϵͳ��� like '" + lstrNoPre + "%')"
    lobjConn.Execute "delete ϵͳ����_ϵͳͼƬ����� where ��ϵͳ��='������' and ͼƬ��� in (select ϵͳ��� from ������_��������Ϣ�� where �������='" & lstrDate & "' and ϵͳ��� like '" + lstrNoPre + "%')"
    lobjConn.Execute "delete ������_��������Ϣ�� where �������='" & lstrDate & "' and ϵͳ��� like '" + lstrNoPre + "%'"
    lobjConn.Execute "delete ������_�����Ա������Ϣ�� where ��������='" & lstrDate & "' and ����������� like '" + lstrNoPre + "%'"
    
    Set lobjRec = dafuncGetData("select * from ������_�����Ա������Ϣ�� where ��������='" & lstrDate & "'")
    cprgStatus.Max = lobjRec.RecordCount
    cprgStatus.Value = 0
    For i = 1 To lobjRec.RecordCount
        lobjConn.Execute "insert into ������_�����Ա������Ϣ��(�����������,������ݺ���,����,�Ա�,����,��������,��λ������,��λ����,��������,��������,Ƭ��,��ҵ���)" & _
            "values('" & lobjRec("�����������") & "','" & lobjRec("������ݺ���") & "','" & lobjRec("����") & "','" & lobjRec("�Ա�") & "','" & lobjRec("����") & "','" & lobjRec("��������") & "','" & lobjRec("��λ������") & "','" & lobjRec("��λ����") & "','" & lobjRec("��������") & "','" & lobjRec("��������") & "','" & lobjRec("Ƭ��") & "','" & lobjRec("��ҵ���") & "')"
        lobjRec.MoveNext
        cprgStatus.Value = i
        DoEvents
    Next
    Set lobjRec = dafuncGetData("select * from ������_��������Ϣ�� where �������='" & lstrDate & "'")
    cprgStatus.Max = lobjRec.RecordCount
    cprgStatus.Value = 0
    For i = 1 To lobjRec.RecordCount
        lobjConn.Execute "insert into ������_��������Ϣ��(ϵͳ���,�����������,��쵥��,�Թܱ��,��������,������,�������,�շ�����,���״̬) values('" & _
            lobjRec("ϵͳ���") & "','" & lobjRec("�����������") & "','" & lobjRec("��쵥��") & "','" & lobjRec("�Թܱ��") & "','" & lobjRec("��������") & "','" & lobjRec("������") & "','" & lobjRec("�������") & "','" & lobjRec("�շ�����") & "','" & lobjRec("���״̬") & "')"
        lobjConn.Execute "insert into ������_�����ʱ�־��(ϵͳ���) values('" & lobjRec("ϵͳ���") & "')"
        lobjRec.MoveNext
        cprgStatus.Value = i
        DoEvents
    Next
    Set lobjRec = dafuncGetData("select ͼƬ��� from ϵͳ����_ϵͳͼƬ����� where ��ϵͳ��='������' and ͼƬ��� in (select ����������� from ������_�����Ա������Ϣ�� where ��������='" & lstrDate & "')")
    cprgStatus.Max = lobjRec.RecordCount
    cprgStatus.Value = 0
    
    Dim lobjPic As StdPicture
    
    For i = 1 To lobjRec.RecordCount
        Set lobjPic = pmfunc��ȡͼƬ(lobjRec(0), "������")
        pmsub����ͼƬ lobjConn, lobjPic, lobjRec(0), "������"
        lobjRec.MoveNext
        cprgStatus.Value = i
        DoEvents
    Next
    Set lobjRec = dafuncGetData("select * from ������_��츽����Ϣ�� where ϵͳ��� in (select ϵͳ��� from ������_��������Ϣ�� where �������='" & lstrDate & "')")
    cprgStatus.Max = lobjRec.RecordCount
    cprgStatus.Value = 0
    For i = 1 To lobjRec.RecordCount
        lobjConn.Execute "insert into ������_��츽����Ϣ��(ϵͳ���,������Ŀ,��Ŀֵ,��Ŀֵ���) values('" & _
            lobjRec("ϵͳ���") & "','" & lobjRec("������Ŀ") & "','" & lobjRec("��Ŀֵ") & "','" & lobjRec("��Ŀֵ���") & "')"
        lobjRec.MoveNext
        cprgStatus.Value = i
        DoEvents
    Next
    Set lobjRec = dafuncGetData("select * from ������_�������Ϣ�� where ϵͳ��� in (select ϵͳ��� from ������_��������Ϣ�� where �������='" & lstrDate & "')")
    cprgStatus.Max = lobjRec.RecordCount
    cprgStatus.Value = 0
    For i = 1 To lobjRec.RecordCount
        lobjConn.Execute "insert into ������_�������Ϣ��(ϵͳ���,�����Ŀ,�����,���ҽʦ,��д����) values('" & _
            lobjRec("ϵͳ���") & "','" & lobjRec("�����Ŀ") & "','" & lobjRec("�����") & "'," & IIf(IsNull(lobjRec("���ҽʦ")), "null", "'" & lobjRec("���ҽʦ") & "'") & ",'" & lobjRec("��д����") & "')"
        lobjRec.MoveNext
        cprgStatus.Value = i
        DoEvents
    Next
    lobjConn.Close
    MsgBox "����ɹ���", vbInformation, "ϵͳ��ʾ"
    mobj����.sub���Ǽ���ֵ "������IP", ctxtServerIP
    cprgStatus.Visible = False
    Exit Sub
errHandle:
    sfsub������ "�����沿��", "FrmTransfer", "ccmdOK_Click", 6666, Error, False
    lobjConn.Close
    Exit Sub
    Resume
End Sub
' ���ܣ�    ����ͼƬ��
' ���룺    paraPicture����Ҫ�����ͼƬ
'           para��ʶ�ţ�����ͼƬ��Ψһ��ʶ�š�
'           para��ϵͳ���������ͼƬ����ϵͳ����
' �����    ��
' ���أ�    ��
' ע���������ñ�ʶ����ϵͳ���Ѷ�Ӧһ��ͼƬ�����滻ԭ�е�ͼƬ��
' ���ߣ�    ����
' ����ʱ�䣺2001-3-5
' �޸�˵���������������Ҫ��ͼƬ����ɸ���ϵͳ��������������ϵͳ���ơ�
' �޸��ˣ�  ����
' �޸�ʱ�䣺2001-3-9
Public Sub pmsub����ͼƬ(paraConn As Connection, ParaPicture As StdPicture, paraID As String, para��ϵͳ�� As String)
    On Error GoTo errHandler
    Dim lstrSql As String              'SQL���
    Dim lrecPicture As ADODB.Recordset           '������䷵��ͼƬ��Ϣ��RecordSet
    Dim lprbPicture As New PropertyBag '��ͼƬ��Ϣ�������л������԰�
    '��ͼƬд�����԰��������л���
    lprbPicture.WriteProperty "Picture", ParaPicture
    '���ݱ�ʶ��ȡ����Ӧ��ͼƬ��
    lstrSql = "select * from ϵͳ����_ϵͳͼƬ����� where ͼƬ���='" & paraID & "' and ��ϵͳ��='" & para��ϵͳ�� & "'"
    Set lrecPicture = New ADODB.Recordset
    lrecPicture.Open lstrSql, paraConn, adOpenKeyset, adLockOptimistic
    'Set lrecPicture = paraConn.Execute(lstrSql)
    '������ؿռ�¼����������һ����¼��
    If lrecPicture.RecordCount = 0 Then
        lrecPicture.AddNew
    End If
    '��ͼƬ��Ϣд���¼���С�
    lrecPicture("ͼƬ").AppendChunk lprbPicture.Contents
    lrecPicture("ͼƬ���") = paraID
    lrecPicture("��ϵͳ��") = para��ϵͳ��
    '�����¼�����¡�
    lrecPicture.Update
    lrecPicture.Close
errHandler:
    Set lrecPicture = Nothing
    Set lprbPicture = Nothing
    Set ParaPicture = Nothing
    If Err.Number = 0 Then Exit Sub
    Err.Raise Err.Number, , Err.Description
End Sub

Private Sub Form_Load()
    Set mobj���� = New cls�û���������
    mobj����.�û���� = "*"
    mobj����.ҵ���� = "������"
    If mobj����.������ֵ("������IP") <> "" Then ctxtServerIP = mobj����.������ֵ("������IP")
    cdtpDate.Value = Date
End Sub
