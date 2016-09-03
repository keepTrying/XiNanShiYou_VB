Attribute VB_Name = "modCtrInput"
Option Explicit

'判断是否中文输入法。
Private Declare Function ImmIsIME Lib "imm32.dll" (ByVal hkl As Long) As Long
'设置输入法。
Private Declare Function ActivateKeyboardLayout Lib "user32" (ByVal hkl As Long, ByVal flags As Long) As Long
'取得现在的输入法。
Private Declare Function GetKeyboardLayout Lib "user32" (ByVal dwLayout As Long) As Long
'取得所有输入法。
Private Declare Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, lpList As Long) As Long

Public Sub sub恢复中文输入法()
    Dim hCurKBDLayout As Long
    Dim Buff As String
    Dim llngCount As Long
    Dim hKB(24) As Long, BuffLen As Long
    Dim hEngKB As Long
    Dim i As Long
    
    On Error GoTo erHandler
    
    hCurKBDLayout = GetKeyboardLayout(0) '取得目前的输入法
    If ImmIsIME(hCurKBDLayout) = 1 Then  '中文输入法
        Buff = String(255, 0)
        llngCount = GetKeyboardLayoutList(25, hKB(0)) '取得所有输入法
        For i = 1 To llngCount
            If ImmIsIME(hKB(i - 1)) = 0 Then '英文输入法
                hEngKB = hKB(i - 1)
            End If
        Next
        ActivateKeyboardLayout hEngKB, 0         '先切换到英文。
        ActivateKeyboardLayout hCurKBDLayout, 0  '再切换中文输入法
    End If
    Exit Sub
erHandler:
End Sub
