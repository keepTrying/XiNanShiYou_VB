Attribute VB_Name = "modCtrInput"
Option Explicit

'�ж��Ƿ��������뷨��
Private Declare Function ImmIsIME Lib "imm32.dll" (ByVal hkl As Long) As Long
'�������뷨��
Private Declare Function ActivateKeyboardLayout Lib "user32" (ByVal hkl As Long, ByVal flags As Long) As Long
'ȡ�����ڵ����뷨��
Private Declare Function GetKeyboardLayout Lib "user32" (ByVal dwLayout As Long) As Long
'ȡ���������뷨��
Private Declare Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, lpList As Long) As Long

Public Sub sub�ָ��������뷨()
    Dim hCurKBDLayout As Long
    Dim Buff As String
    Dim llngCount As Long
    Dim hKB(24) As Long, BuffLen As Long
    Dim hEngKB As Long
    Dim i As Long
    
    On Error GoTo erHandler
    
    hCurKBDLayout = GetKeyboardLayout(0) 'ȡ��Ŀǰ�����뷨
    If ImmIsIME(hCurKBDLayout) = 1 Then  '�������뷨
        Buff = String(255, 0)
        llngCount = GetKeyboardLayoutList(25, hKB(0)) 'ȡ���������뷨
        For i = 1 To llngCount
            If ImmIsIME(hKB(i - 1)) = 0 Then 'Ӣ�����뷨
                hEngKB = hKB(i - 1)
            End If
        Next
        ActivateKeyboardLayout hEngKB, 0         '���л���Ӣ�ġ�
        ActivateKeyboardLayout hCurKBDLayout, 0  '���л��������뷨
    End If
    Exit Sub
erHandler:
End Sub
