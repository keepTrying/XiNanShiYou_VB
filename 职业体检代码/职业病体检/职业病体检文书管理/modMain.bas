Attribute VB_Name = "modMain"
Option Explicit

' Modified code from an article in Visual Studio Magazine by Hank Marquis
' http://www.fawcette.com/vsm/2002_05/magazine/columns/desktopdeveloper/default_pf.asp

'Private Const VER_PLATFORM_WIN32s As Long = 0
Private Const VER_PLATFORM_WIN32_WINDOWS As Long = 1
Private Const VER_PLATFORM_WIN32_NT As Long = 2

Private Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformID As Long
  szCSDVersion As String * 128
End Type

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Function WinVer() As String
  ' Returns the Operating System
  Dim tOVI As OSVERSIONINFO

  tOVI.dwOSVersionInfoSize = Len(tOVI)      ' Should be 148 : (128 + (4 * 5))  - Long data type is 4 bytes

  If GetVersionEx(tOVI) = 1 Then
    If tOVI.dwPlatformID = VER_PLATFORM_WIN32_NT And tOVI.dwMajorVersion = 5 And tOVI.dwMinorVersion = 1 Then
      WinVer = "XP"
    ElseIf tOVI.dwPlatformID = VER_PLATFORM_WIN32_NT And tOVI.dwMajorVersion = 5 And tOVI.dwMinorVersion = 0 Then
      WinVer = "00"
    ElseIf tOVI.dwPlatformID = VER_PLATFORM_WIN32_NT And tOVI.dwMajorVersion = 4 Then
      WinVer = "NT"
    ElseIf tOVI.dwPlatformID = VER_PLATFORM_WIN32_WINDOWS And tOVI.dwMajorVersion = 4 And tOVI.dwMinorVersion = 90 Then
      WinVer = "ME"
    ElseIf (tOVI.dwPlatformID = VER_PLATFORM_WIN32_WINDOWS) And (tOVI.dwMajorVersion > 4) Or (tOVI.dwMajorVersion = 4 And tOVI.dwMinorVersion > 0) Then
      WinVer = "98"
    ElseIf tOVI.dwPlatformID = VER_PLATFORM_WIN32_WINDOWS And tOVI.dwMajorVersion = 4 And tOVI.dwMinorVersion = 0 Then
      WinVer = "95"
    End If
  End If

  If Len(Trim$(WinVer)) = 0 Then
    ' OS not recognized, probably need to update this module for a new OS
    ' or possible that GetVersionEX API call didn't work properly
    WinVer = "UN"     ' Unknown
  End If
End Function


