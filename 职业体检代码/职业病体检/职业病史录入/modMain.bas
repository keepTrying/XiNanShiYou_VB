Attribute VB_Name = "modMain"
'***************************************
'���ƣ�ְҵ��ʷ(�ܼ��߸�����Ϣ)����������
'������func������
'���ܣ�����������������
'���ߣ�Yunle Liu
'ʱ�䣺2012.03
'***************************************



Option Explicit

Public pobjDict As Object     'clsDictionary��

Public ���ʼǺ� As Integer   '����ְҵ��ʷ¼����޸�
Public pubϵͳ��� As String
Public pobjҵ����� As Object '������ҵ�����clsManageMedicalExam��
Public bolenProject As Boolean  '�Ƿ���ȷ�������Ŀ

Public Sub Main()
    On Error Resume Next
     '�����ֵ����
    Set pobjDict = CreateObject("�ֵ����.clsDictionary")
    Err.Clear
    '����ҵ�����
    Set pobjҵ����� = CreateObject("ְҵ������.clsManageMedicalExam")
    If Err <> 0 Then
        On Error GoTo errHandler
        Err.Raise 6666, , "�޷���������������ݶ���������ע�ᡰְҵ��ʷ¼��.dll����"
    End If
    Exit Sub
errHandler:
    sfsub������ "ְҵ��ʷ¼��", "modmain", "Main", Err.Number, Err.Description, False
End Sub

Public Function func������(ByVal paraErrNumber As Long, ByVal paraErrDes As String) As String
    Select Case paraErrNumber
    Case 6
        func������ = "�������ݹ����ѳ���ϵͳ�涨��С��"
    Case -2147217833
        func������ = "�������ݹ���������󣩣��ѳ���ϵͳ�涨���ȣ����С����"
    Case -2147217913
        func������ = "���ڸ�ʽ�Ƿ���"
    Case -2147217873 '��������ڡ�
        func������ = "ϵͳ�������������Ϊ��" & Chr(13) & Chr(10) & "(1) �����ڱ���������漰�������Ϣ�ѱ���ɾ����" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "����취��" & Chr(13) & Chr(10) & "(1) ���˳���ҵ��������棬���½��롣"
    Case 94 '��Чʹ��Null��
        func������ = "ʹ�õ��ֵ����ͨ���ֵ�������ɾ���ˣ�ϵͳ�޷��ټ���������������ϵͳ����Ա�ָ��ֵ����ݡ���ע�⣬��Ҫ���ɾ���ֵ��"
    Case 336, 337, 338, 429, 430
        func������ = "ϵͳ�������𻵣����Ѷ�ʧ����ϵͳ�޷����������С�" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "����취��" & Chr(13) & Chr(10) & "(1) ���˳�ϵͳ�������°�װϵͳ��"
    Case 440 '�ⲿ����������Զ�����
        func������ = "ϵͳ������������ֹ���С�" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "����취��" & Chr(13) & Chr(10) & "(1) ���˳�ϵͳ�����½��롣"
    Case 91 '����û�г�ʼ���ɹ���
        func������ = "��Ϊ����������ϵͳ��������ʱ�޷���������ĳ�ʼ�������˳����ܽ��棬�����½��빦�ܽ��档" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "����취��" & Chr(13) & Chr(10) & "(1) ���˳�ϵͳ�����½��롣"
    Case 5
        func������ = "��Ϊ�����жϣ�����������ϵͳ�޷��������С�" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "����취��" & Chr(13) & Chr(10) & "(1) ���˳�ϵͳ�����½��롣"
    Case Else
        func������ = paraErrDes
    End Select
End Function


