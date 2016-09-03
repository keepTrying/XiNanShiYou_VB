Attribute VB_Name = "mdl主程序"
Option Explicit


Public pobj平台结构 As Object  '定义平台结构

Public pblnCancel As Boolean            '是否确认退出
Public pblnExit As Boolean              '是退出或注销
Public pbln注销 As Boolean

Public pcol字典集 As New Collection


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

Public pcolWndProc As New Collection       '子窗体Proc
Public pcol操作名称 As New Collection      '子窗体操作名
Public pcol业务名称 As New Collection      '子窗体所属业务名称
Public pcol子窗体句柄 As New Collection    '子窗体句柄
Public plngMainHwnd As Long                '主窗体句柄
Public pstrMainCaption As String           '主窗体Caption

Public pstrSysName As String

Public pstr子系统许可 As String        '修改：2003-7-9（杨春）加密狗上的子系统许可。
Public pstr版本代号 As String
Public pbln试用 As Boolean

Public pstr用户编号 As String           '用户的唯一编号

Public pstrServer As String
Public pobj信使客户端 As Object
'功能：系统运行从此模块进入
'输入：无
'输出：无
'返回：无
'注意事项：无
'作者：王晓华
'创建时间：2001-3-6
Sub Main()
    Dim i As Long
    
    '忽略所有错误。
    On Error Resume Next
    pbln试用 = True
        
    '判断该系统是否已经运行
    If App.PrevInstance = True Then
        Dim lstrTitle As String 'AppTitle
        lstrTitle = App.Title
        App.Title = ""
        AppActivate lstrTitle
        End
    End If
    
    '判断自动升级程序自身是否已经升级，若是，重新运行它
    Dim lstrDestination  As String
    Dim lstrSource As String
    
    lstrSource = App.Path & "\AutoUpgradeFile\AutoUpgrade.exe"
    If Dir(lstrSource) <> "" Then
        MsgBox "系统需要重新进行升级，请点击“确定”按钮进行升级！", vbInformation, "系统提示"
        lstrDestination = App.Path & "\AutoUpgrade.exe"
        SetAttr lstrDestination, vbNormal
        FileCopy lstrSource, lstrDestination
        Kill lstrSource
        Shell lstrDestination, vbNormalFocus
        End
    End If
    
    '获取命令行参数：系统名称。
    Dim lngCount As Long            '参数个数。
    Dim varCom As Variant           '参数数组。
    lngCount = 10
    varCom = funcGetCommandLine(lngCount)
    If lngCount >= 1 Then
        pstrSysName = varCom(1)
        If lngCount > 1 Then
            pstr版本代号 = varCom(2)
        End If
    
    Else
        pstrSysName = "卫生防疫管理信息系统"
        pstr版本代号 = "S" '标准版。
    End If
    
    '修改：2003-9-29（杨春）根据南京用户需求，需要在1台工作站上连接两个系统。
    '根据系统名称或取配置路径。
    Dim lstrSubSec As String
    lstrSubSec = "系统管理"
    If pstrSysName Like "疾病控制*" Then
        lstrSubSec = "疾控系统"
    ElseIf pstrSysName Like "卫生监督*" Then
        lstrSubSec = "监督系统"
    End If
        
    '获取系统配置。
    Dim lstrServer As String       '服务器名
    Dim lstrDatabase As String     '数据库名
    Dim lstrDogServer As String    '网络锁服务器名。
    lstrServer = sffuncGetSetting(lstrSubSec, "数据库配置", "服务器名")
    lstrDatabase = sffuncGetSetting(lstrSubSec, "数据库配置", "数据库名")
    lstrDogServer = sffuncGetSetting(lstrSubSec, "数据库配置", "网络锁服务器名")
    
    '若不是运行全系统，根据各子系统配置修改“系统管理”下的配置。
    If pstrSysName <> "卫生防疫管理信息系统" Then
        If lstrServer <> "" Then
            '更新odbc数据源。
            Dim strAttributes As String
            
            strAttributes = "Database=" & lstrDatabase & _
                vbCr & "Description=" & "" & _
                vbCr & "OemToAnsi=No" & _
                vbCr & "Server=" & lstrServer
            DBEngine.RegisterDatabase "WSFY2001", "SQL Server", True, strAttributes
        
            sfsubSaveSetting "系统管理", "数据库配置", "服务器名", lstrServer
            sfsubSaveSetting "系统管理", "数据库配置", "数据库名", lstrDatabase
            sfsubSaveSetting "系统管理", "数据库配置", "网络锁服务器名", lstrDogServer
        Else
            lstrServer = sffuncGetSetting("系统管理", "数据库配置", "服务器名")
            lstrDatabase = sffuncGetSetting("系统管理", "数据库配置", "数据库名")
            lstrDogServer = sffuncGetSetting("系统管理", "数据库配置", "网络锁服务器名")
        End If
    End If
    
    '修改：2002-11-7（杨春）强制进行区域设置，保证日期格式正确。对win2000立刻生效，win98需要重新启动。
    sub区域设置
       
    On Error Resume Next
'
'    '初始化数据访问对象
    dasubInitialize ("Provider=MSDataShape.1;Data Provider=SQLOLEDB.1;Password=welcome;Persist Security Info=True;User ID=user26;Initial Catalog=" & lstrDatabase & ";Data Source=" & lstrServer)
    If Err.Number <> 0 Then
        '换windows安全模式登陆。
        Err.Clear
        On Error GoTo errHandle
        dasubInitialize ("driver={SQL Server};Database=" & lstrDatabase & ";Server=" & lstrServer)
    End If
    
    dasubInitialize lstrServer
    
    '初始化数据访问对象(使用用于 SQL Server 的 OLE DB 提供者)
    If Err.Number <> 0 Then
        '换windows安全模式登陆。
        Err.Clear
        dasubInitialize ("Provider=MSDataShape.1;Data Provider=SQLOLEDB.1;Password=welcome58%*;Persist Security Info=True;User ID=jk_user;Initial Catalog=" & lstrDatabase & ";Data Source=" & lstrServer)
    End If
    
    If Err.Number <> 0 Then
        '换windows安全模式登陆。
        Err.Clear
        dasubInitialize "Provider=sqloledb;Data Source=" & lstrServer & ";Initial Catalog=" & lstrDatabase & _
                        "Integrated Security=SSPI"
        'dasubInitialize ("driver={SQL Server};Database=" & lstrDatabase & ";Server=" & lstrServer)
    End If
    
    '使用用于 ODBC 的 OLE DB 提供者（不使用 ODBC 数据源）：
    If Err.Number <> 0 Then
        Err.Clear
        dasubInitialize "Driver={SQL Server};" & _
                        "Server=" & lstrServer & ";Database=" & lstrDatabase & ";" & _
                        "Uid=jk_user;Pwd=welcome58%*"
    End If
    
    '使用用于 ODBC 的 OLE DB 提供者(windows安全模式)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo errHandle
        dasubInitialize "Driver={SQL Server};Server=" & lstrServer & ";Database=" & lstrDatabase & ";Trusted_Connection=yes"
    End If
    
    '根据服务器时间刷新本地工作站日期。
    Dim lstrDate As String
    lstrDate = (dafuncGetData("select getdate()").Fields(0))
    Date = CDate(lstrDate)
    Time = CDate(lstrDate)
    
    On Error GoTo errHandle
    
    Set pcol字典集 = New Collection
    Set pcol业务名称 = New Collection
    Set pcol操作名称 = New Collection
    Set pcolWndProc = New Collection
'    '修改：2002-8-26（杨春）判断上次传输日期到现在间隔是否超过三天，并且是否有数据未传输。
'    Dim lstrError As String
'    lstrError = func判断共享数据传输是否及时()
'    If lstrError <> "" Then
'        MsgBox lstrError & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "请找系统管理员解决：先使用“共享数据传输工具”把所有未传输的健康证、许可信息卡、计划免疫卡信息传输到省共享数据中心。", vbOKOnly + vbExclamation, "系统提示"
'        End
'    End If
    
    '创建平台结构的一个实例。
    Set pobj平台结构 = CreateObject("通用对象.cls平台结构")
    frmSplash.Show       '显示系统信息
    
    Dim lstrTime As String
    
    lstrTime = Now
    Do While DateDiff("s", lstrTime, Now) < 2
        DoEvents
    Loop
    
    '关闭splash,显示登录窗体
    FrmLogin.clblSysName = "请输入帐号和口令以进入" & pstrSysName & "："
    FrmLogin.Show vbModal
    Unload frmSplash
       
    dlsub解除死锁
errHandle:
    If Err.Number = 0 Then Exit Sub
    Call sfsub错误处理("主程序", "mdl主程序", "sub Main", Err.Number, Err.Description, False)
    Unload frmSplash
End Sub

'功能：捕获窗体消息的窗口函数（回调函数）。
'注意事项：请勿在主程序中设定断点调试程序。
'作者：罗庆
'创建时间：2001-4-17
Public Function funcClassing(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    Dim lobj As cls主程序
    Set lobj = New cls主程序
    funcClassing = lobj.funcClassing(hWnd, Msg, wParam, lParam)
    Exit Function
    Static lblnTerminate As Boolean
End Function

'Public Function SetWindowLong(ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'    SetWindowLong = 0
'End Function

'功能：判断健康证、许可证卡、计划免疫卡传输是否及时。
'返回：不及时的错误信息。
'修改：2003-3-28（杨春）间隔改为7天。
Private Function func判断共享数据传输是否及时() As String
    Dim lstr最后日期 As String
    Dim llngCount As Long
    Dim lobjRec As Object
    Dim lstrResult As String
    
    On Error GoTo errHandler
    
    '判断健康证是否及时。
    Set lobjRec = dafuncGetData("select max(传输日期) from 健康证_传输记录表")
    lstr最后日期 = IIf(IsNull(lobjRec(0)), "", lobjRec(0))
    If lstr最后日期 = "" Then
        '从未传输过，取最早健康证打印日期。
        Set lobjRec = dafuncGetData("select min(领取日期) from 健康证_健康证信息表 where isnull(领取日期,'1945-01-01')>'1945-01-01'")
        lstr最后日期 = IIf(IsNull(lobjRec(0)), "", lobjRec(0))
    End If
    '获取未传输的健康证数量。
    Set lobjRec = dafuncGetData("select count(*) from 健康证_健康证信息表 a  INNER JOIN 健康证_健康证状态字典视图 c ON a.健康证状态 = c.InnerID and c.名称 = '已发放' and 系统编号 not in (select 系统编号 from 健康证_传输记录表)")
    llngCount = IIf(IsNull(lobjRec(0)), 0, lobjRec(0))
    If llngCount > 0 Then
        '修改：2003-3-28（杨春）间隔改为7天。
        If DateDiff("d", lstr最后日期, Now) > 7 Then
            '已超过7天未传输过数据了。
            lstrResult = "已超过七天未传输健康证数据到全省共享数据库！"
        ElseIf llngCount >= 1000 Then
            lstrResult = "已有累计超过1000条健康证数据未传输到全省共享数据库！"
        End If
    End If
    
    '判断许可证卡是否及时。
    Set lobjRec = dafuncGetData("select max(传输日期) from 许可证_传输记录表")
    lstr最后日期 = IIf(IsNull(lobjRec(0)), "", lobjRec(0))
    If lstr最后日期 = "" Then
        '从未传输过，取最早许可证卡发放日期。
        Set lobjRec = dafuncGetData("select min(发放日期) from 许可证_单位卡发放记录表")
        lstr最后日期 = IIf(IsNull(lobjRec(0)), "", lobjRec(0))
    End If
    '获取未传输的许可证卡数量。
    Set lobjRec = dafuncGetData("select count(*) from 许可证_单位卡发放记录表 where 卡号 not in (select 卡号 from 许可证_传输记录表)")
    llngCount = IIf(IsNull(lobjRec(0)), 0, lobjRec(0))
    If llngCount > 0 Then
        '修改：2003-3-28（杨春）间隔改为7天。
        If DateDiff("d", lstr最后日期, Now) > 7 Then
            '已超过七天未传输过数据了。
            lstrResult = IIf(lstrResult = "", "", lstrResult & Chr(13) & Chr(10)) & "已超过七天未传输许可信息卡数据到全省共享数据库！"
        ElseIf llngCount >= 1000 Then
            lstrResult = IIf(lstrResult = "", "", lstrResult & Chr(13) & Chr(10)) & "已有累计超过1000条许可信息卡数据未传输到全省共享数据库！"
        End If
    End If
    
    '判断计划免疫卡是否及时。
    Set lobjRec = dafuncGetData("select max(传输日期) from 计划免疫_传输记录表")
    lstr最后日期 = IIf(IsNull(lobjRec(0)), "", lobjRec(0))
    If lstr最后日期 = "" Then
        '从未传输过，取最早儿童登记日期。
        Set lobjRec = dafuncGetData("select min(登记日期) from 计划免疫_儿童基本信息表 where isnull(卡号,'')<>'' and 儿童状态<>'异地'")
        lstr最后日期 = IIf(IsNull(lobjRec(0)), "", lobjRec(0))
    End If
    '获取未传输的计划免疫卡数量。
    Set lobjRec = dafuncGetData("select count(*) from 计划免疫_儿童基本信息表 where isnull(卡号,'')<>'' and 儿童状态<>'异地' and 卡号 not in (select 卡号 from 计划免疫_传输记录表)")
    llngCount = IIf(IsNull(lobjRec(0)), 0, lobjRec(0))
    If llngCount > 0 Then
        '修改：2003-3-28（杨春）间隔改为7天。
        If DateDiff("d", lstr最后日期, Now) >= 7 Then
            '已超过七天未传输过数据了。
            lstrResult = IIf(lstrResult = "", "", lstrResult & Chr(13) & Chr(10)) & "已超过七天未传输计划免疫卡数据到全省共享数据库！"
        ElseIf llngCount >= 1000 Then
            lstrResult = IIf(lstrResult = "", "", lstrResult & Chr(13) & Chr(10)) & "已有累计超过1000条计划免疫卡数据未传输到全省共享数据库！"
        End If
    End If
    
    func判断共享数据传输是否及时 = lstrResult
    Exit Function
errHandler:
End Function

'功能：获取本机名。
'创建：2001-11-16
'作者：杨春。
Public Function funcGetLocalName() As String
    Dim lstrLocal As String * 255 '本机名。
    Dim llngLen As Long
    
    On Error Resume Next
    
    llngLen = 60
    Call GetComputerName(lstrLocal, llngLen)
    funcGetLocalName = Trim(lstrLocal)
    '去掉字符串后面的0。
    Do While Asc(Right(funcGetLocalName, 1)) = 0
        funcGetLocalName = Left(funcGetLocalName, Len(funcGetLocalName) - 1)
    Loop

End Function

'功能：设置本机的日期格式为yyyy/mm/dd，时间格式为hh:mm:ss。
'创建：2002-11-7（杨春）。
Public Sub sub区域设置()
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


'功能：获取命令行参数，各参数必须用引号引起。以逗号隔开。
Private Function funcGetCommandLine(paraMaxArgs As Long)
    Dim c, strCmdLine, intCmdLnLen, i, intArgsNum
    Dim lblnBeginQuato As Boolean      '是否已开始引号内。
    ReDim arrArgs(1 To paraMaxArgs)    '读取的参数。
    
    
    strCmdLine = Command() '取得命令行参数。
    intCmdLnLen = Len(strCmdLine)
    lblnBeginQuato = False
    intArgsNum = 0
    
    '以一次一个字符的方式取出命令行参数。
    For i = 1 To intCmdLnLen
        c = Mid(strCmdLine, i, 1)
        
        If c = "'" Then
            lblnBeginQuato = Not lblnBeginQuato
        End If
        If lblnBeginQuato Then
            If c = "'" Then
                '新的参数。
                '检测参数是否过多。
                If intArgsNum = paraMaxArgs Then Exit For
                intArgsNum = intArgsNum + 1
            End If
            '将字符加到当前参数中。
            If c <> "'" Then
                arrArgs(intArgsNum) = arrArgs(intArgsNum) & c
            End If
        End If
        
    Next i
    
    '返回实际的参数个数
    paraMaxArgs = intArgsNum
    If intArgsNum > 0 Then
        '调整数组大小使其刚好符合参数个数。
    ReDim Preserve arrArgs(1 To intArgsNum)
    End If
    For i = 1 To paraMaxArgs
        arrArgs(i) = Trim(arrArgs(i))
    Next i
    
    '将数组返回。
    funcGetCommandLine = arrArgs()
End Function


Public Sub sub登录信使服务()
    '获取本机名称。
    Dim lstrLocalName As String
    Dim i As Long
    
    lstrLocalName = funcGetLocalName()
    
    '修改：2002-8-5（杨春）登录信使服务器。
    '修改：2002-8-30（杨春）服务器上不能登录信使服务。
    On Error Resume Next
    If UCase(Trim(pstrServer)) <> UCase(Trim(lstrLocalName)) Then
        If pbln注销 Then
            Set pobj信使客户端 = Nothing
            For i = 1 To 30000
                DoEvents
            Next
        End If
        
        Set pobj信使客户端 = CreateObject("信使客户端.cls信使服务客户端")
        pobj信使客户端.sub登录信使服务 um用户名, um用户所属科室
        Err.Clear
    End If
End Sub

Public Sub sub退出信使服务()
    On Error Resume Next
    Call pobj信使客户端.sub关闭连结
    Set pobj信使客户端 = Nothing
End Sub
