Attribute VB_Name = "modMain"
Option Explicit
'���ߣ��

Public pobjҵ����� As Object '������ҵ�����clsManageMedicalExam��
Public pobjDict As Object     '�ֵ����clsDictionary��
Public flag_database_delete As Boolean '���ݿ�ǿ��ɾ����־ 'german
Public flag_delete_any As Boolean '�Ƿ�����ɾ�� 'german

Public Sub Main()
    On Error Resume Next
    flag_database_delete = False 'Ĭ�ϲ�ǿ��ɾ�����ݿ� 'german���˱�־�ڸ�ѡ��İ�ť�¼��б��޸�
    flag_delete_any = False 'Ĭ�ϲ�����ɾ��������Ŀ 'german���˱�־�ڸ�ѡ��İ�ť�¼��б��޸�
    '����ҵ�����
    Set pobjҵ����� = CreateObject("ְҵ������.clsManageMedicalExam")
        
    Err.Clear
    
    '�����ֵ����
    Set pobjDict = CreateObject("�ֵ����.clsDictionary")
    If Err <> 0 Then
        '�ֵ���������ã������Լ�������ֵ�������
        Set pobjDict = New clsDictionary
    End If
    
End Sub

Public Function func������(ByVal paraErrNumber As Long, ByVal paraErrDes As String) As String
    Select Case paraErrNumber
    Case -2147217833
        func������ = "�������ݹ���������󣩣��ѳ���ϵͳ�涨���ȣ����С����"
    Case 6
        func������ = "�������ݹ����ѳ���ϵͳ�涨��С��"
    Case -2147217913
        func������ = "���ڸ�ʽ�Ƿ���"
    Case 94 '��Чʹ��Null��
        func������ = "ʹ�õ��ֵ����ͨ���ֵ�������ɾ���ˣ�ϵͳ�޷��ټ���������������ϵͳ����Ա�ָ��ֵ����ݡ���ע�⣬��Ҫ���ɾ���ֵ��"
    Case 336, 337, 338, 429, 430
        func������ = "ϵͳ�������𻵣����Ѷ�ʧ����ϵͳ�޷����������С����˳�ϵͳ�������°�װϵͳ��"
    Case 440 '�ⲿ����������Զ�����
        func������ = "ϵͳ������������ֹ���С����˳�ϵͳ������������ϵͳ��"
    Case 91 '����û�г�ʼ���ɹ���
        func������ = "��Ϊ����������ϵͳ��������ʱ�޷���������ĳ�ʼ�������˳����ܽ��棬�����½��빦�ܽ��档"
    Case 5
        func������ = "��Ϊ�����жϣ�����������ϵͳ�޷��������С�"
    Case Else
        func������ = paraErrDes
    End Select
End Function
