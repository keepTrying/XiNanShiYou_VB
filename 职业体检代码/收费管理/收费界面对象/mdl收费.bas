Attribute VB_Name = "mdl�շ�"
'Public pint���ۿ��� As Integer '����ҵ�����õĴ��ۿ�����Ϣ
'Public pint��Ŀ���� As Integer '����ҵ�����õĿ�Ŀ����

Public pobjҵ������ As Object
Public pobj�շѹ��� As Object
Public pobj��λ��λ As Object  '��λ�����ӿ�


Public Sub Main()
    Set pobj�շѹ��� = CreateObject("�շѶ��󲿼�.cls�շѹ���")
    Set pobj��λ��λ = CreateObject("��λ����ҵ��.ClsUnitInterface")

End Sub


Public Function func¼��Ʊ�ݺ�() As String
    Dim lstrƱ�ݺ� As String
    Dim lobjRec As Object
    Set lobjRec = dafuncGetData("select ��ǰֵ from ϵͳ����_ϵͳ������ɼ�¼�� where ҵ������='�շѹ���" & um�û���� & "' and �������='�վݺ�'")
    If lobjRec.RecordCount = 0 Then
        dafuncGetData "insert into ϵͳ����_ϵͳ������ɼ�¼��(ҵ������,�������,��������,��ǰֵ,����,�Ƿ����ر�,��ǰ���) values('�շѹ���" & um�û���� & "','�վݺ�','C',0,9,'��',2008)"
        Set lobjRec = dafuncGetData("select ��ǰֵ from ϵͳ����_ϵͳ������ɼ�¼�� where ҵ������='�շѹ���" & um�û���� & "' and �������='�վݺ�'")
    End If
    lstrƱ�ݺ� = IIf(IsNull(lobjRec(0)), "", lobjRec(0))
    lLen = Len(lstrƱ�ݺ�)
    lstrƱ�ݺ� = Format(Val(lstrƱ�ݺ�) + 1, String(lLen, "0"))
    lstrƱ�ݺ� = InputBox("�����뵱ǰƱ�ݺţ�", "ϵͳѯ��", lstrƱ�ݺ�)
    
    Do While lstrƱ�ݺ� <> ""
        lLen = Len(lstrƱ�ݺ�)
        Set lobjRec = dafuncGetData("select �վݺ� from �շѹ���_������Ϣ�� where �վݺ�='" & lstrƱ�ݺ� & "'")
        If lobjRec.RecordCount Then
            MsgBox "��Ʊ�ݺ��Ѿ�ʹ���ˣ������ٴ�ʹ�á�", vbCritical, "ϵͳ��ʾ"
            lstrƱ�ݺ� = InputBox("�����뵱ǰƱ�ݺţ�", "ϵͳѯ��", lstrƱ�ݺ�)
        Else
            lstrƱ�ݺ� = Format(Val(lstrƱ�ݺ�) - 1, String(lLen, "0"))
            dafuncGetData "update ϵͳ����_ϵͳ������ɼ�¼�� set ��ǰֵ='" & lstrƱ�ݺ� & "' where ҵ������='�շѹ���" & um�û���� & "' and �������='�վݺ�'"
            Exit Do
        End If
    Loop
    
    func¼��Ʊ�ݺ� = lstrƱ�ݺ�
End Function


'����: ���������ת��Ϊ����ҵĴ�д�ַ���
'����: money       ���
'���: FuncConvertToCapsStr     ת���Ĵ�д�ַ���
'����޸�ʱ��: 96.6.11
'--------------------------------------------------
Public Function FuncConvertToCapsStr(Money As Currency) As String
On Error GoTo errhandle
    Const digit_str = "��Ҽ��������½��ƾ�"
    Const unit_str = "Ǫ��ʰ��Ǫ��ʰԪ�Ƿ�"
    Dim money_str As String
    
    If Money > 99999999.99 Then
        FuncConvertToCapsStr = ""
    ElseIf Money = 0 Then
        FuncConvertToCapsStr = "��Ԫ��"
    Else
        Dim temp_str As String
        Dim i, j As Integer
        
        If Money < 0 Then
            money_str = "��"
            Money = -Money
        Else
            money_str = ""
        End If
        
        temp_str = Format(Money, "00000000.00")
        
        'ת����������
        For i = 1 To 8
            If Mid(temp_str, i, 1) <> "0" Then Exit For
        Next
        For i = i To 8
            j = CInt(Mid(temp_str, i, 1))
            If j > 0 Then
                money_str = money_str & Mid(digit_str, j + 1, 1) & Mid(unit_str, i, 1)
            Else
                If i = 4 Then
                    money_str = money_str & "��"
                ElseIf i = 8 Then
                    money_str = money_str & "Ԫ"
                ElseIf Mid(temp_str, i + 1, 1) <> "0" Then
                    money_str = money_str & Mid(digit_str, j + 1, 1)
                End If
            End If
        Next
        
        'ת��С������
        If Right(temp_str, 2) = "00" Then
            money_str = money_str & "��"
        Else
            'ת����
            j = CInt(Mid(temp_str, 10, 1))
            If j > 0 Then
                money_str = money_str & Mid(digit_str, j + 1, 1) & "��"
            Else
                money_str = money_str & "��"
            End If
            'ת����
            j = CInt(Mid(temp_str, 11, 1))
            If j > 0 Then
                money_str = money_str & Mid(digit_str, j + 1, 1) & "��"
            Else
                money_str = money_str & "��"
            End If
        End If
        
        FuncConvertToCapsStr = money_str
    End If
    Exit Function
errhandle:
    sfsub������ "�շѽ��沿��", "mdl�շ�", " FuncConvertToCapsStr()", Err.Number, Err.Description, True
End Function



Public Function gfuncKeyNum(paraKey As Integer) As Integer
    On Error Resume Next
    Select Case paraKey
        Case 8, 13, 46, 48 To 57
        ' �����˸��(ASCII��Ϊ8)���س���(ASCII��Ϊ13)��С��������ּ�(ASCII��Ϊ48��57)
            gfuncKeyNum = paraKey

        Case Else
            gfuncKeyNum = 0 ' �����������գ���ASCII��Ϊ0���ַ���ʾ������������Ϣ
    End Select
End Function
