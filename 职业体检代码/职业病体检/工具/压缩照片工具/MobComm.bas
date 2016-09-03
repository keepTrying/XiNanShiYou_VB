Attribute VB_Name = "MobComm"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3CF4937E0290"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Module"

Option Explicit
      
Public Plng数据期限 As String
Public pbln清空 As Boolean
Public PStr照片目录 As String

'公共变量
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const BIF_RETURNONLYFSDIRS = &H1

Public Type BROWSEINFO
        hOwner As Long
        pidlRoot As Long
        pszDisplayName As String
        lpszTitle As String
        ulFlags As Long
        lpfn As Long
        lParam As Long
        iImage As Long
End Type
      
Public Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long



'功能: 初始化数据访问对象,初始化程序,是整个程序的入口
Sub Main()
    Dim lstrServer As String
    Dim lstrData  As String
    Dim llngCount As Long
    Dim i As Long
    
    On Error Resume Next
    '判断该系统是否已经运行
    If App.PrevInstance = True Then
        Dim lstrTitle As String 'AppTitle
        lstrTitle = App.Title
        App.Title = ""
        AppActivate lstrTitle
        End
    End If
    
    On Error GoTo errHandler
    
    lstrServer = sffuncGetSetting("系统管理", "数据库配置", "服务器名")
    lstrData = sffuncGetSetting("系统管理", "数据库配置", "数据库名")
    
    '初始化数据访问对象。
Connect:    dasubInitialize "Provider=MSDataShape.1;Data Provider=SQLOLEDB.1;Password=welcome;Persist Security Info=True;User ID=user26;Initial Catalog=" & lstrData & ";Data Source=" & lstrServer
    
    frm导照片.Show

    Exit Sub
errHandler:
    MsgBox "导照片工具程序初始化失败！解决办法：" & Chr(13) & Chr(10) & "请检查数据库服务是否处于正常运行状态！" & Chr(13) & Chr(10) & "防疫2001库是否存在？"
End Sub

Public Sub subExit()


End Sub


