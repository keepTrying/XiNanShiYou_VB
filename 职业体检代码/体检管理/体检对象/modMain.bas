Attribute VB_Name = "modMain"
Public Const P_SUBSYSNAME = "������"   '��ϵͳ���ơ�

'�������ļ���
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
            (ByVal lpApplicationName As String, _
                ByVal lpKeyName As Any, _
                ByVal lpDefault As String, _
                ByVal lpReturnedString As String, _
                ByVal nSize As Long, _
                ByVal lpFileName As String) As Long
'Щ�����ļ���
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
            (ByVal lpApplicationName As String, _
                ByVal lpKeyName As Any, _
                ByVal lpString As Any, _
                ByVal lpFileName As String) As Long


Public pobjDict As Object         '�ֵ����

Public Sub Main()
    On Error Resume Next
    '�����ֵ����
    Set pobjDict = CreateObject("�ֵ����.clsDictionary")
    If Err <> 0 Then
        '�����Զ�����ֵ����
        Set pobjDict = New clsDictionary
    End If
    
End Sub


Public Function func������(ByVal paraErrNumber As Long, ByVal paraErrDes As String) As String
    Select Case paraErrNumber
    Case -2147217833
        func������ = "�������ݹ���������󣩣��ѳ���ϵͳ�涨���ȣ����С����"
    Case -2147217913
        func������ = "���ݸ�ʽ���������ݿ�Ҫ�󣨱��磺���ڸ�ʽ�Ƿ�������ϵͳҪ����ֵ���ͣ����������ַ����ͣ���"
    Case -2147217873
        func������ = "ϵͳ�޷�����������Ϊ��" & Chr(13) & Chr(10) & "�ôδ�����ص���Ϣ�ѱ���ɾ����"
    Case 6
        func������ = "�������ݹ����ѳ���ϵͳ�涨��С��"
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
    Case 482
        func������ = "��ӡʧ�ܣ���Ϊ�Ҳ�����ӡ���������ӡ���Ƿ�������"
    Case Else
        func������ = paraErrDes
    End Select
End Function

