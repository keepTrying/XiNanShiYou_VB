Attribute VB_Name = "MobComm"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3CF4937E0290"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Module"

Option Explicit
      
Public Plng�������� As String
Public pbln��� As Boolean
Public PStr��ƬĿ¼ As String

'��������
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



'����: ��ʼ�����ݷ��ʶ���,��ʼ������,��������������
Sub Main()
    Dim lstrServer As String
    Dim lstrData  As String
    Dim llngCount As Long
    Dim i As Long
    
    On Error Resume Next
    '�жϸ�ϵͳ�Ƿ��Ѿ�����
    If App.PrevInstance = True Then
        Dim lstrTitle As String 'AppTitle
        lstrTitle = App.Title
        App.Title = ""
        AppActivate lstrTitle
        End
    End If
    
    On Error GoTo errHandler
    
    lstrServer = sffuncGetSetting("ϵͳ����", "���ݿ�����", "��������")
    lstrData = sffuncGetSetting("ϵͳ����", "���ݿ�����", "���ݿ���")
    
    '��ʼ�����ݷ��ʶ���
Connect:    dasubInitialize "Provider=MSDataShape.1;Data Provider=SQLOLEDB.1;Password=welcome;Persist Security Info=True;User ID=user26;Initial Catalog=" & lstrData & ";Data Source=" & lstrServer
    
    frm����Ƭ.Show

    Exit Sub
errHandler:
    MsgBox "����Ƭ���߳����ʼ��ʧ�ܣ�����취��" & Chr(13) & Chr(10) & "�������ݿ�����Ƿ�����������״̬��" & Chr(13) & Chr(10) & "����2001���Ƿ���ڣ�"
End Sub

Public Sub subExit()


End Sub


