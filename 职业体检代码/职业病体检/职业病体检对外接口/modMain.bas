Attribute VB_Name = "modMain"
Option Explicit

'���ܣ���paraRange[���ݷ�Χ�������ݷ�Χֵ] ���ϵ�ֵ����������򵼳�������,��ƴ��һ��SQL���������.
'���룺paraRange As Collection [���ݷ�Χ�������ݷ�Χֵ]�����ݷ�Χ������ʼ���ڣ���ʼ���ڣ���λ���Ƽ�����ϵͳ���,��ϵͳ��š�
'       paraInput (true ����/false ����)
'���أ���and ��ʼ��string.  ��"and ������� >= convert(datetime,'2001-3-1') and ��λ���� in ('��λA','��λb')
'���ߣ����ơ�
Public Function funcFilter(ByVal paraRange As Collection, Optional paraInput As Boolean = True) As String
    Dim lstrSQL As String
    Dim lstrUnit As String      '��ŵ�λ��ż��ִ�,����ƴSQL���
    Dim lintUnit As Integer     '��Ƕ������ִ��е�λ��,����ƴSQL���ĵ�λ���(ÿ���������ŵı���ö��ŷָ�)
    
    On Err GoTo errHandler
    
    If Not paraRange Is Nothing Then
        '��ʼ���ڡ�
        If sffunc�жϼ��ϼ�ֵ�Ƿ����(paraRange, "��ʼ����") Then
            If paraRange("��ʼ����")("���ݷ�Χֵ") <> "" Then
                If paraInput Then
                    lstrSQL = lstrSQL & " and ������_���������ݿ�.������� >= #" & paraRange("��ʼ����")("���ݷ�Χֵ") & "#"
                Else
                    lstrSQL = lstrSQL & " and ������_���������ݿ�.������� >= '" & paraRange("��ʼ����")("���ݷ�Χֵ") & "'"
                End If
            End If
        End If
        '�������ڡ�
        If sffunc�жϼ��ϼ�ֵ�Ƿ����(paraRange, "��������") Then
            If paraRange("��������")("���ݷ�Χֵ") <> "" Then
                If paraInput Then
                    lstrSQL = lstrSQL & " and ������_���������ݿ�.������� <= #" & paraRange("��������")("���ݷ�Χֵ") & "#"
                Else
                    lstrSQL = lstrSQL & " and ������_���������ݿ�.������� <= '" & paraRange("��������")("���ݷ�Χֵ") & "'"
                End If
            End If
        End If
        '��λ���Ƽ���
        If sffunc�жϼ��ϼ�ֵ�Ƿ����(paraRange, "��λ���Ƽ�") Then
            lstrUnit = Trim(paraRange("��λ���Ƽ�")("���ݷ�Χֵ"))
            If lstrUnit <> "" Then
                lstrSQL = lstrSQL & " and ������_���������ݿ�.��λ���� in ("
                lstrUnit = lstrUnit & IIf(Right(lstrUnit, 1) <> ",", ",", "")
                While Len(lstrUnit) > 0
                    lintUnit = InStr(1, lstrUnit, ",")
                    lstrSQL = lstrSQL & "'" & Left(lstrUnit, lintUnit - 1) & "',"
                    lstrUnit = Right(lstrUnit, Len(lstrUnit) - lintUnit)
                Wend
                lstrSQL = Left(lstrSQL, Len(lstrSQL) - 1) & ")"
            End If
        End If
        'ϵͳ��ŷ�Χ��
        If sffunc�жϼ��ϼ�ֵ�Ƿ����(paraRange, "��ϵͳ���") Then
            If paraRange("��ϵͳ���")("���ݷ�Χֵ") <> "" Then
                lstrSQL = lstrSQL & " and ������_���������ݿ�.ϵͳ��� >= '" & paraRange("��ϵͳ���")("���ݷ�Χֵ") & "'"
            End If
        End If
        If sffunc�жϼ��ϼ�ֵ�Ƿ����(paraRange, "��ϵͳ���") Then
            If paraRange("��ϵͳ���")("���ݷ�Χֵ") <> "" Then
                lstrSQL = lstrSQL & " and ������_���������ݿ�.ϵͳ��� <= '" & paraRange("��ϵͳ���")("���ݷ�Χֵ") & "'"
            End If
        End If
        '������(��������)
        If sffunc�жϼ��ϼ�ֵ�Ƿ����(paraRange, "������") Then
            lstrSQL = lstrSQL & " and ������_���������ݿ�.�������� = '" & paraRange("������")("���ݷ�Χֵ") & "'"
        End If
        '��������
        If sffunc�жϼ��ϼ�ֵ�Ƿ����(paraRange, "��������") Then
            If paraRange("��������")("���ݷ�Χֵ") Then
                lstrSQL = lstrSQL & " and ������_���������ݿ�.���״̬ = " & P_ENDED_STATUS
            End If
        End If
        
    End If
    funcFilter = lstrSQL
    Exit Function
errHandler:
    sfsub������ "������ӿڲ���", "ClsManngeTransmission", "funcFilter", Err.Number, Err.Description, True
End Function

