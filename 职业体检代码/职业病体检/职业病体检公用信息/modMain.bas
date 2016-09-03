Attribute VB_Name = "modMain"
'读配置文件。
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
            (ByVal lpApplicationName As String, _
                ByVal lpKeyName As Any, _
                ByVal lpDefault As String, _
                ByVal lpReturnedString As String, _
                ByVal nSize As Long, _
                ByVal lpFileName As String) As Long
'些配置文件。
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
            (ByVal lpApplicationName As String, _
                ByVal lpKeyName As Any, _
                ByVal lpString As Any, _
                ByVal lpFileName As String) As Long


Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const CB_SHOWDROPDOWN = &H14F


Public Sub Main()
    On Error Resume Next
    '调用“cls系统函数”的初始化接口，以便创建其隐形实例。
    'sfsubInit
End Sub
