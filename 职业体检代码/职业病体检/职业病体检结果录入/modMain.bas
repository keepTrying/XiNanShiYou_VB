Attribute VB_Name = "modMain"
'2012-03-01 �ڵ��
'���ӽ��¼�������к�������

Option Explicit
Public pobjҵ����� As Object
Public InputFlag As String
Public InputFlagNo As String
Public coptIndex As Integer

Private Sub Main()
    On Error GoTo errHandler
    
    '����ҵ�������ʱ��������ʾ��λ��Ϣ�ϡ�
    Set pobjҵ����� = CreateObject("ְҵ������.clsManageMedicalExam")
    
    Exit Sub
errHandler:
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

Public Sub sub��ʾ��λ����(ByVal ciptBase As Control, _
            ByVal para��λ������ As String, _
            ByVal paraGUI As cls����ͨ�ö���)
    Dim i As Long
    Dim lcolInfo As Collection
    
    If para��λ������ <> "" Then
    
        '��ȡ��λ���ԡ�
        On Error Resume Next
        '��ȡ��λ���ԡ�
        Set lcolInfo = pobjҵ�����.func��ȡ��λ����(para��λ������)
        
        ciptBase.pblnTemp = True
        
        ciptBase.Box1("��������").TrueText = ""
        ciptBase.Box1("��ҵ���").TrueText = ""
        ciptBase.Box1("Ƭ��").TrueText = ""
        ciptBase.Box1("��������").TrueText = ""
        
        ciptBase.Box1("��������").TrueText = lcolInfo("��������")
        ciptBase.Box1("��ҵ���").TrueText = lcolInfo("��ҵ���")
        ciptBase.Box1("Ƭ��").TrueText = lcolInfo("Ƭ��")
        ciptBase.Box1("��������").TrueText = lcolInfo("��������")
        
        ciptBase.Box1("��������").Text = lcolInfo("������������")
        ciptBase.Box1("��ҵ���").Text = lcolInfo("��ҵ�������")
        ciptBase.Box1("Ƭ��").Text = lcolInfo("Ƭ������")
        ciptBase.Box1("��������").Text = lcolInfo("������������")
        ciptBase.Box1("��λ��ַ").Text = lcolInfo("��λ��ַ")
        
        
        
        Dim lstrItem As String
        Dim lint�������� As Integer
        Dim lint��ҵ���  As Integer
        
        Err.Clear
        
        '�ж��Ƿ����������ࡣ
        For i = 1 To ciptBase.InfoCollection.Count
            '¼����Ŀ���ơ�
            lstrItem = ciptBase.InfoCollection(i).Title
            
            If lstrItem = "��������" Then
                lint�������� = i
            ElseIf lstrItem = "��ҵ���" Then
                lint��ҵ��� = i
            End If
            If Err <> 0 Then Exit For
        Next i
        
        '������ҵ���¼�����ֵ����ݵ�������
        Dim lstrItemTrueText As String
        Dim lobjRec As Object
        If lint�������� > 0 And lint��ҵ��� > 0 Then
            '��ȡ���������š�
            lstrItemTrueText = ciptBase.ItemTrueText(lint�������� - 1)
            '������ҵ���¼�����ֵ䡣
            If lstrItemTrueText <> "" And Not ciptBase.InfoCollection(lint��������).DictRecordSet Is Nothing Then
                Set lobjRec = ciptBase.InfoCollection(lint��������).DictRecordSet
                If Not lobjRec.EOF Then
                    paraGUI.sub��ʼ���ֵ�� lint��ҵ���, "Parent=" & lobjRec("InnerId")
                End If
            End If
        End If
        
        ciptBase.pblnTemp = False
    End If

End Sub

