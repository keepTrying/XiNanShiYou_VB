Attribute VB_Name = "mdl������"
Option Explicit


Public pobjƽ̨�ṹ As Object  '����ƽ̨�ṹ

Public pblnCancel As Boolean            '�Ƿ�ȷ���˳�
Public pblnExit As Boolean              '���˳���ע��
Public pblnע�� As Boolean

Public pcol�ֵ伯 As New Collection


Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Integer, ByVal Y As Integer, ByVal CX As Integer, ByVal CY As Integer, ByVal wFlags As Integer)
Public Declare Function GetTopWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Long) As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long


Private Const HKEY_CURRENT_USER = &H80000001
Private Const REG_SZ = 1

Public Const GWL_STYLE = (-16)
Public Const WS_BORDER = &H800000
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOMOVE = &H2
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_DRAWFRAME = &H20
Public Const GWL_WNDPROC = (-4)
Public Const WM_ACTIVATE = &H6
Public Const WM_DESTROY = &H2
Public Const SWP_NOOWNERZORDER = &H200

Public Const HWND_TOP = 0
Public Const HWND_BOTTOM = 1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Public pcolWndProc As New Collection       '�Ӵ���Proc
Public pcol�������� As New Collection      '�Ӵ��������
Public pcolҵ������ As New Collection      '�Ӵ�������ҵ������
Public pcol�Ӵ����� As New Collection    '�Ӵ�����
Public plngMainHwnd As Long                '��������
Public pstrMainCaption As String           '������Caption

Public pstrSysName As String

Public pstr��ϵͳ��� As String        '�޸ģ�2003-7-9��������ܹ��ϵ���ϵͳ��ɡ�
Public pstr�汾���� As String
Public pbln���� As Boolean

Public pstr�û���� As String           '�û���Ψһ���

Public pstrServer As String
Public pobj��ʹ�ͻ��� As Object
'���ܣ�ϵͳ���дӴ�ģ�����
'���룺��
'�������
'���أ���
'ע�������
'���ߣ�������
'����ʱ�䣺2001-3-6
Sub Main()
    Dim i As Long
    
    '�������д���
    On Error Resume Next
    pbln���� = True
        
    '�жϸ�ϵͳ�Ƿ��Ѿ�����
    If App.PrevInstance = True Then
        Dim lstrTitle As String 'AppTitle
        lstrTitle = App.Title
        App.Title = ""
        AppActivate lstrTitle
        End
    End If
    
    '�ж��Զ��������������Ƿ��Ѿ����������ǣ�����������
    Dim lstrDestination  As String
    Dim lstrSource As String
    
    lstrSource = App.Path & "\AutoUpgradeFile\AutoUpgrade.exe"
    If Dir(lstrSource) <> "" Then
        MsgBox "ϵͳ��Ҫ���½���������������ȷ������ť����������", vbInformation, "ϵͳ��ʾ"
        lstrDestination = App.Path & "\AutoUpgrade.exe"
        SetAttr lstrDestination, vbNormal
        FileCopy lstrSource, lstrDestination
        Kill lstrSource
        Shell lstrDestination, vbNormalFocus
        End
    End If
    
    '��ȡ�����в�����ϵͳ���ơ�
    Dim lngCount As Long            '����������
    Dim varCom As Variant           '�������顣
    lngCount = 10
    varCom = funcGetCommandLine(lngCount)
    If lngCount >= 1 Then
        pstrSysName = varCom(1)
        If lngCount > 1 Then
            pstr�汾���� = varCom(2)
        End If
    
    Else
        pstrSysName = "�������߹�����Ϣϵͳ"
        pstr�汾���� = "S" '��׼�档
    End If
    
    '�޸ģ�2003-9-29����������Ͼ��û�������Ҫ��1̨����վ����������ϵͳ��
    '����ϵͳ���ƻ�ȡ����·����
    Dim lstrSubSec As String
    lstrSubSec = "ϵͳ����"
    If pstrSysName Like "��������*" Then
        lstrSubSec = "����ϵͳ"
    ElseIf pstrSysName Like "�����ල*" Then
        lstrSubSec = "�ලϵͳ"
    End If
        
    '��ȡϵͳ���á�
    Dim lstrServer As String       '��������
    Dim lstrDatabase As String     '���ݿ���
    Dim lstrDogServer As String    '����������������
    lstrServer = sffuncGetSetting(lstrSubSec, "���ݿ�����", "��������")
    lstrDatabase = sffuncGetSetting(lstrSubSec, "���ݿ�����", "���ݿ���")
    lstrDogServer = sffuncGetSetting(lstrSubSec, "���ݿ�����", "��������������")
    
    '����������ȫϵͳ�����ݸ���ϵͳ�����޸ġ�ϵͳ�����µ����á�
    If pstrSysName <> "�������߹�����Ϣϵͳ" Then
        If lstrServer <> "" Then
            '����odbc����Դ��
            Dim strAttributes As String
            
            strAttributes = "Database=" & lstrDatabase & _
                vbCr & "Description=" & "" & _
                vbCr & "OemToAnsi=No" & _
                vbCr & "Server=" & lstrServer
            DBEngine.RegisterDatabase "WSFY2001", "SQL Server", True, strAttributes
        
            sfsubSaveSetting "ϵͳ����", "���ݿ�����", "��������", lstrServer
            sfsubSaveSetting "ϵͳ����", "���ݿ�����", "���ݿ���", lstrDatabase
            sfsubSaveSetting "ϵͳ����", "���ݿ�����", "��������������", lstrDogServer
        Else
            lstrServer = sffuncGetSetting("ϵͳ����", "���ݿ�����", "��������")
            lstrDatabase = sffuncGetSetting("ϵͳ����", "���ݿ�����", "���ݿ���")
            lstrDogServer = sffuncGetSetting("ϵͳ����", "���ݿ�����", "��������������")
        End If
    End If
    
    '�޸ģ�2002-11-7�����ǿ�ƽ����������ã���֤���ڸ�ʽ��ȷ����win2000������Ч��win98��Ҫ����������
    sub��������
       
    On Error Resume Next
'
'    '��ʼ�����ݷ��ʶ���
    dasubInitialize ("Provider=MSDataShape.1;Data Provider=SQLOLEDB.1;Password=welcome;Persist Security Info=True;User ID=user26;Initial Catalog=" & lstrDatabase & ";Data Source=" & lstrServer)
    If Err.Number <> 0 Then
        '��windows��ȫģʽ��½��
        Err.Clear
        On Error GoTo errHandle
        dasubInitialize ("driver={SQL Server};Database=" & lstrDatabase & ";Server=" & lstrServer)
    End If
    
    dasubInitialize lstrServer
    
    '��ʼ�����ݷ��ʶ���(ʹ������ SQL Server �� OLE DB �ṩ��)
    If Err.Number <> 0 Then
        '��windows��ȫģʽ��½��
        Err.Clear
        dasubInitialize ("Provider=MSDataShape.1;Data Provider=SQLOLEDB.1;Password=welcome58%*;Persist Security Info=True;User ID=jk_user;Initial Catalog=" & lstrDatabase & ";Data Source=" & lstrServer)
    End If
    
    If Err.Number <> 0 Then
        '��windows��ȫģʽ��½��
        Err.Clear
        dasubInitialize "Provider=sqloledb;Data Source=" & lstrServer & ";Initial Catalog=" & lstrDatabase & _
                        "Integrated Security=SSPI"
        'dasubInitialize ("driver={SQL Server};Database=" & lstrDatabase & ";Server=" & lstrServer)
    End If
    
    'ʹ������ ODBC �� OLE DB �ṩ�ߣ���ʹ�� ODBC ����Դ����
    If Err.Number <> 0 Then
        Err.Clear
        dasubInitialize "Driver={SQL Server};" & _
                        "Server=" & lstrServer & ";Database=" & lstrDatabase & ";" & _
                        "Uid=jk_user;Pwd=welcome58%*"
    End If
    
    'ʹ������ ODBC �� OLE DB �ṩ��(windows��ȫģʽ)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo errHandle
        dasubInitialize "Driver={SQL Server};Server=" & lstrServer & ";Database=" & lstrDatabase & ";Trusted_Connection=yes"
    End If
    
    '���ݷ�����ʱ��ˢ�±��ع���վ���ڡ�
    Dim lstrDate As String
    lstrDate = (dafuncGetData("select getdate()").Fields(0))
    Date = CDate(lstrDate)
    Time = CDate(lstrDate)
    
    On Error GoTo errHandle
    
    Set pcol�ֵ伯 = New Collection
    Set pcolҵ������ = New Collection
    Set pcol�������� = New Collection
    Set pcolWndProc = New Collection
'    '�޸ģ�2002-8-26������ж��ϴδ������ڵ����ڼ���Ƿ񳬹����죬�����Ƿ�������δ���䡣
'    Dim lstrError As String
'    lstrError = func�жϹ������ݴ����Ƿ�ʱ()
'    If lstrError <> "" Then
'        MsgBox lstrError & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "����ϵͳ����Ա�������ʹ�á��������ݴ��乤�ߡ�������δ����Ľ���֤�������Ϣ�����ƻ����߿���Ϣ���䵽ʡ�����������ġ�", vbOKOnly + vbExclamation, "ϵͳ��ʾ"
'        End
'    End If
    
    '����ƽ̨�ṹ��һ��ʵ����
    Set pobjƽ̨�ṹ = CreateObject("ͨ�ö���.clsƽ̨�ṹ")
    frmSplash.Show       '��ʾϵͳ��Ϣ
    
    Dim lstrTime As String
    
    lstrTime = Now
    Do While DateDiff("s", lstrTime, Now) < 2
        DoEvents
    Loop
    
    '�ر�splash,��ʾ��¼����
    FrmLogin.clblSysName = "�������ʺźͿ����Խ���" & pstrSysName & "��"
    FrmLogin.Show vbModal
    Unload frmSplash
       
    dlsub�������
errHandle:
    If Err.Number = 0 Then Exit Sub
    Call sfsub������("������", "mdl������", "sub Main", Err.Number, Err.Description, False)
    Unload frmSplash
End Sub

'���ܣ���������Ϣ�Ĵ��ں������ص���������
'ע��������������������趨�ϵ���Գ���
'���ߣ�����
'����ʱ�䣺2001-4-17
Public Function funcClassing(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    Dim lobj As cls������
    Set lobj = New cls������
    funcClassing = lobj.funcClassing(hWnd, Msg, wParam, lParam)
    Exit Function
    Static lblnTerminate As Boolean
End Function

'Public Function SetWindowLong(ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'    SetWindowLong = 0
'End Function

'���ܣ��жϽ���֤�����֤�����ƻ����߿������Ƿ�ʱ��
'���أ�����ʱ�Ĵ�����Ϣ��
'�޸ģ�2003-3-28����������Ϊ7�졣
Private Function func�жϹ������ݴ����Ƿ�ʱ() As String
    Dim lstr������� As String
    Dim llngCount As Long
    Dim lobjRec As Object
    Dim lstrResult As String
    
    On Error GoTo errHandler
    
    '�жϽ���֤�Ƿ�ʱ��
    Set lobjRec = dafuncGetData("select max(��������) from ����֤_�����¼��")
    lstr������� = IIf(IsNull(lobjRec(0)), "", lobjRec(0))
    If lstr������� = "" Then
        '��δ�������ȡ���罡��֤��ӡ���ڡ�
        Set lobjRec = dafuncGetData("select min(��ȡ����) from ����֤_����֤��Ϣ�� where isnull(��ȡ����,'1945-01-01')>'1945-01-01'")
        lstr������� = IIf(IsNull(lobjRec(0)), "", lobjRec(0))
    End If
    '��ȡδ����Ľ���֤������
    Set lobjRec = dafuncGetData("select count(*) from ����֤_����֤��Ϣ�� a  INNER JOIN ����֤_����֤״̬�ֵ���ͼ c ON a.����֤״̬ = c.InnerID and c.���� = '�ѷ���' and ϵͳ��� not in (select ϵͳ��� from ����֤_�����¼��)")
    llngCount = IIf(IsNull(lobjRec(0)), 0, lobjRec(0))
    If llngCount > 0 Then
        '�޸ģ�2003-3-28����������Ϊ7�졣
        If DateDiff("d", lstr�������, Now) > 7 Then
            '�ѳ���7��δ����������ˡ�
            lstrResult = "�ѳ�������δ���佡��֤���ݵ�ȫʡ�������ݿ⣡"
        ElseIf llngCount >= 1000 Then
            lstrResult = "�����ۼƳ���1000������֤����δ���䵽ȫʡ�������ݿ⣡"
        End If
    End If
    
    '�ж����֤���Ƿ�ʱ��
    Set lobjRec = dafuncGetData("select max(��������) from ���֤_�����¼��")
    lstr������� = IIf(IsNull(lobjRec(0)), "", lobjRec(0))
    If lstr������� = "" Then
        '��δ�������ȡ�������֤���������ڡ�
        Set lobjRec = dafuncGetData("select min(��������) from ���֤_��λ�����ż�¼��")
        lstr������� = IIf(IsNull(lobjRec(0)), "", lobjRec(0))
    End If
    '��ȡδ��������֤��������
    Set lobjRec = dafuncGetData("select count(*) from ���֤_��λ�����ż�¼�� where ���� not in (select ���� from ���֤_�����¼��)")
    llngCount = IIf(IsNull(lobjRec(0)), 0, lobjRec(0))
    If llngCount > 0 Then
        '�޸ģ�2003-3-28����������Ϊ7�졣
        If DateDiff("d", lstr�������, Now) > 7 Then
            '�ѳ�������δ����������ˡ�
            lstrResult = IIf(lstrResult = "", "", lstrResult & Chr(13) & Chr(10)) & "�ѳ�������δ���������Ϣ�����ݵ�ȫʡ�������ݿ⣡"
        ElseIf llngCount >= 1000 Then
            lstrResult = IIf(lstrResult = "", "", lstrResult & Chr(13) & Chr(10)) & "�����ۼƳ���1000�������Ϣ������δ���䵽ȫʡ�������ݿ⣡"
        End If
    End If
    
    '�жϼƻ����߿��Ƿ�ʱ��
    Set lobjRec = dafuncGetData("select max(��������) from �ƻ�����_�����¼��")
    lstr������� = IIf(IsNull(lobjRec(0)), "", lobjRec(0))
    If lstr������� = "" Then
        '��δ�������ȡ�����ͯ�Ǽ����ڡ�
        Set lobjRec = dafuncGetData("select min(�Ǽ�����) from �ƻ�����_��ͯ������Ϣ�� where isnull(����,'')<>'' and ��ͯ״̬<>'���'")
        lstr������� = IIf(IsNull(lobjRec(0)), "", lobjRec(0))
    End If
    '��ȡδ����ļƻ����߿�������
    Set lobjRec = dafuncGetData("select count(*) from �ƻ�����_��ͯ������Ϣ�� where isnull(����,'')<>'' and ��ͯ״̬<>'���' and ���� not in (select ���� from �ƻ�����_�����¼��)")
    llngCount = IIf(IsNull(lobjRec(0)), 0, lobjRec(0))
    If llngCount > 0 Then
        '�޸ģ�2003-3-28����������Ϊ7�졣
        If DateDiff("d", lstr�������, Now) >= 7 Then
            '�ѳ�������δ����������ˡ�
            lstrResult = IIf(lstrResult = "", "", lstrResult & Chr(13) & Chr(10)) & "�ѳ�������δ����ƻ����߿����ݵ�ȫʡ�������ݿ⣡"
        ElseIf llngCount >= 1000 Then
            lstrResult = IIf(lstrResult = "", "", lstrResult & Chr(13) & Chr(10)) & "�����ۼƳ���1000���ƻ����߿�����δ���䵽ȫʡ�������ݿ⣡"
        End If
    End If
    
    func�жϹ������ݴ����Ƿ�ʱ = lstrResult
    Exit Function
errHandler:
End Function

'���ܣ���ȡ��������
'������2001-11-16
'���ߣ����
Public Function funcGetLocalName() As String
    Dim lstrLocal As String * 255 '��������
    Dim llngLen As Long
    
    On Error Resume Next
    
    llngLen = 60
    Call GetComputerName(lstrLocal, llngLen)
    funcGetLocalName = Trim(lstrLocal)
    'ȥ���ַ��������0��
    Do While Asc(Right(funcGetLocalName, 1)) = 0
        funcGetLocalName = Left(funcGetLocalName, Len(funcGetLocalName) - 1)
    Loop

End Function

'���ܣ����ñ��������ڸ�ʽΪyyyy/mm/dd��ʱ���ʽΪhh:mm:ss��
'������2002-11-7�������
Public Sub sub��������()
    On Error Resume Next
    Dim llngKey As Long
    Dim lstrValue As String
    
    RegCreateKey HKEY_CURRENT_USER, "Control Panel\International", llngKey
        
    lstrValue = "/"
    RegSetValueEx llngKey, "sDate", 0, REG_SZ, ByVal lstrValue, LenB(lstrValue)
    
    lstrValue = "yyyy/MM/dd"
    RegSetValueEx llngKey, "sShortDate", 0, REG_SZ, ByVal lstrValue, LenB(lstrValue)
    
    lstrValue = ":"
    RegSetValueEx llngKey, "sTime", 0, REG_SZ, ByVal lstrValue, LenB(lstrValue)
    
    lstrValue = "HH:mm:ss"
    RegSetValueEx llngKey, "sTimeFormat", 0, REG_SZ, ByVal lstrValue, LenB(lstrValue)
    RegCloseKey llngKey
    
End Sub


'���ܣ���ȡ�����в��������������������������Զ��Ÿ�����
Private Function funcGetCommandLine(paraMaxArgs As Long)
    Dim c, strCmdLine, intCmdLnLen, i, intArgsNum
    Dim lblnBeginQuato As Boolean      '�Ƿ��ѿ�ʼ�����ڡ�
    ReDim arrArgs(1 To paraMaxArgs)    '��ȡ�Ĳ�����
    
    
    strCmdLine = Command() 'ȡ�������в�����
    intCmdLnLen = Len(strCmdLine)
    lblnBeginQuato = False
    intArgsNum = 0
    
    '��һ��һ���ַ��ķ�ʽȡ�������в�����
    For i = 1 To intCmdLnLen
        c = Mid(strCmdLine, i, 1)
        
        If c = "'" Then
            lblnBeginQuato = Not lblnBeginQuato
        End If
        If lblnBeginQuato Then
            If c = "'" Then
                '�µĲ�����
                '�������Ƿ���ࡣ
                If intArgsNum = paraMaxArgs Then Exit For
                intArgsNum = intArgsNum + 1
            End If
            '���ַ��ӵ���ǰ�����С�
            If c <> "'" Then
                arrArgs(intArgsNum) = arrArgs(intArgsNum) & c
            End If
        End If
        
    Next i
    
    '����ʵ�ʵĲ�������
    paraMaxArgs = intArgsNum
    If intArgsNum > 0 Then
        '���������Сʹ��պ÷��ϲ���������
    ReDim Preserve arrArgs(1 To intArgsNum)
    End If
    For i = 1 To paraMaxArgs
        arrArgs(i) = Trim(arrArgs(i))
    Next i
    
    '�����鷵�ء�
    funcGetCommandLine = arrArgs()
End Function


Public Sub sub��¼��ʹ����()
    '��ȡ�������ơ�
    Dim lstrLocalName As String
    Dim i As Long
    
    lstrLocalName = funcGetLocalName()
    
    '�޸ģ�2002-8-5�������¼��ʹ��������
    '�޸ģ�2002-8-30������������ϲ��ܵ�¼��ʹ����
    On Error Resume Next
    If UCase(Trim(pstrServer)) <> UCase(Trim(lstrLocalName)) Then
        If pblnע�� Then
            Set pobj��ʹ�ͻ��� = Nothing
            For i = 1 To 30000
                DoEvents
            Next
        End If
        
        Set pobj��ʹ�ͻ��� = CreateObject("��ʹ�ͻ���.cls��ʹ����ͻ���")
        pobj��ʹ�ͻ���.sub��¼��ʹ���� um�û���, um�û���������
        Err.Clear
    End If
End Sub

Public Sub sub�˳���ʹ����()
    On Error Resume Next
    Call pobj��ʹ�ͻ���.sub�ر�����
    Set pobj��ʹ�ͻ��� = Nothing
End Sub
