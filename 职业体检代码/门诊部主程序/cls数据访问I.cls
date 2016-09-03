VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls数据访问I"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Const Mint连接超时时间 = 6
Const Mint查询超时时间 = 6

Private WithEvents mcls消息传递 As cls消息传递   '与其它对象间进行消息传递的类
Attribute mcls消息传递.VB_VarHelpID = -1
Private mblnNotFreePcn As Boolean
Public pcn As ADODB.Connection                  'ADO数据连接对象

Public pblnFirstTry As Boolean                  '第一次重试连接
Public pblnNoTryReConnect As Boolean            '不再重试连接
Public pintTranCount As Integer                 '当前TranCount
Public pstrDSN As String                        '与服务器数据库进行连接的字符串
Public pblnConnect As Boolean                   '表示是否处于连接状态

Public pintInitializeTimes As Integer           '数据访问对象被初始化的次数。
Private Type dataField                           '返回的数据列的定义
    name As String
    type As Integer
    len As Long
End Type


Public Property Get dapstrDSN() As String
    dapstrDSN = pstrDSN
End Property

Public Property Let dapstrDSN(ByVal NewValue As String)
    '
End Property

'在类初始化时获取消息传递的唯一实例。
Private Sub Class_Initialize()
    On Error GoTo errHandler
    '初始化次数自增长。
    pintInitializeTimes = pintInitializeTimes + 1
    Set mcls消息传递 = sf消息传递
    
    Exit Sub
errHandler:
    Set mcls消息传递 = Nothing
End Sub

'在类释放时关闭数据库连接。
Private Sub Class_Terminate()
    On Error Resume Next
    '初始化次数递减。
    pintInitializeTimes = pintInitializeTimes - 1
    '如果所有对数据访问对象的引用全部释放则释放模块中定义的公用变量。
    If pintInitializeTimes = 0 Then
        pcn.Close
        Set pcn = Nothing
        Set mcls消息传递 = Nothing
    End If
End Sub

'初始化数据库连接。
Public Sub dasubInitialize(paraDSN As String)
    If paraDSN <> "" Then
        pstrDSN = paraDSN
        pblnConnect = True
    End If
    
    '设置变量缺省值。
    pblnFirstTry = False
    pblnNoTryReConnect = False
End Sub

Public Sub subDisconnect()
    pblnConnect = False
End Sub
Public Sub dasubDisconnect()
    pblnConnect = False
End Sub

' 功能：    执行查询语句。
' 输入：    paraSQL：SQL语句
' 输出：    无
' 返回：    保存返回值的RecordSet对象
' 注意事项：对于不需返回值的SQL语句如UPDATE语句也使用此方法执行，返回值不可用。
' 作者：    刘洪
' 创建时间：2004-3-25
Public Function dafuncGetData(ByVal paraSQL As String) As ADODB.Recordset
    Dim lstrErr As String
    Dim lrecTemp As ADODB.Recordset   'SQL语句返回的RecordSet
    Dim lstrResult As String, lobjDoc As Object
    Dim lobjHttp As Object, lstrWebserviceURL As String
    Dim lstrRequest As String, lintStatus As Integer
    Dim lstrProcName As String
    Dim lstrSql As String
    
    If pstrDSN = "" Then
        Err.Raise 6666, , "没有设置中心服务器的域名！请先调用dasubInitialize进行设置。"
        Exit Function
    End If
    
    On Error GoTo errHandler
    
    Set lobjHttp = CreateObject("MSXML2.XMLHTTP")
    Set lobjDoc = CreateObject("MSXML.DOMDocument")
    lstrWebserviceURL = "http://" & pstrDSN & "/DataAccessService/clsDataAccess.asmx/funcGetRecordset"
    '设置参数及其值
    lstrRequest = "pstrSql=" & funcConvert(paraSQL)
'    lstrRequest = "pstrSql=" & paraSQL ' & "&pstrErrMsg="
    lobjHttp.Open "POST", lstrWebserviceURL, False
    '设置这个Content-Type很重要
    lobjHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    lobjHttp.send (lstrRequest)
    If lobjHttp.Status <> 200 Then
        Err.Raise lobjHttp.Status, , lobjHttp.statusText & vbCrLf & "SQL语句：" & paraSQL
        Exit Function
    End If
    lintStatus = lobjDoc.Load(lobjHttp.responseXML)
    lstrResult = lobjDoc.xml
'    lstrResult = Replace(lstrResult, "&lt;", "<", 1, -1, 1)
'    lstrResult = Replace(lstrResult, "&gt;", ">", 1, -1, 1)
'    lstrResult = Replace(lstrResult, "&amp;", "&", 1, -1, 1)
        
    lobjDoc.loadXML lstrResult
    
    '把返回的数据构建成Recordset
    Dim node As IXMLDOMNodeList
    '取出结构描述信息，并转换成字段描述信息
    Set node = lobjDoc.selectNodes("//xs:sequence")
    If node.length = 0 Then Exit Function       '没有结构信息，说明没有返回数据
    
    Dim lintFieldCount As Integer, lobjFields() As dataField
    Dim i As Integer
    
    lintFieldCount = node.Item(0).childNodes.length     '字段个数
    
    ReDim lobjFields(1 To lintFieldCount)
    
    For i = 1 To lintFieldCount
        lobjFields(i).name = node.Item(0).childNodes(i - 1).Attributes(0).nodeValue   '字段名
        Select Case Mid(node.Item(0).childNodes(i - 1).Attributes(1).nodeValue, 4)    '字段类型
            Case "boolean"
                lobjFields(i).type = 11         '"adBoolean"
            Case "long"
                lobjFields(i).type = 20         '"adBigint"
            Case "double"
                lobjFields(i).type = 5          '"adDouble"
            Case "string"
                lobjFields(i).type = 200        '"adVarchar"
                lobjFields(i).len = 1           '初始为1，在构建数据插入语句时将其调整为记录集中对应数据的最大长度
            Case "base64Binary"
                lobjFields(i).type = 128        '"adBinary"
                lobjFields(i).len = 1           '初始为1，在构建数据插入语句时将其调整为记录集中对应数据的最大长度
            Case "dateTime"
                lobjFields(i).type = 7          '"adDate"
            Case "int"
                lobjFields(i).type = 3          '"adInteger"
            Case "float", "decimal"
                lobjFields(i).type = 4         '"adDecimal(16)"
                lobjFields(i).len = 4
            Case "short"
                lobjFields(i).type = 2          '"adSmallInt"
            Case "unsignedByte"
                lobjFields(i).type = 17         '"adUnsignedTinyInt"
            Case Else
                If Mid(node.Item(0).childNodes(i - 1).Attributes(1).nodeValue, 8, 4) = "Guid" Then
'                    lobjFields(i).type = 72         '"adGUID"
                    lobjFields(i).type = 200        '作为adVarchar对待
                    lobjFields(i).len = 36
                Else
                    lobjFields(i).type = "ad" & Mid(node.Item(0).childNodes(i - 1).Attributes(1).nodeValue, 4)
                End If
        End Select
    Next
    '对字符串类型字段的长度进行调整
    Dim lintRecordCount As Integer, lstrInsert As String, j As Integer
    Dim curNode As IXMLDOMNode, curNode1 As IXMLDOMNode, lintLen As Long
    Dim lstrRecord() As Variant
    
    Set lrecTemp = New Recordset
    
    Set node = lobjDoc.selectNodes("//Table")
    If node.length = 0 Then        '没有记录信息，说明是一个空记录集
        '构建记录集的结构
        With lrecTemp
            For i = 1 To lintFieldCount
                If lobjFields(i).len > 0 Then
                    .Fields.Append lobjFields(i).name, lobjFields(i).type, lobjFields(i).len, adFldIsNullable
                Else
                    .Fields.Append lobjFields(i).name, lobjFields(i).type, , adFldIsNullable
                End If
            Next
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .Open       ' 不需要连接对象。
        End With
        Set dafuncGetData = lrecTemp
        Set lrecTemp = Nothing
        Exit Function
    End If
    
    lintRecordCount = node.length     '记录个数
    ReDim lstrRecord(1 To lintRecordCount, 1 To lintFieldCount)
    Dim lobjBase64 As New clsBase64
    
    For i = 1 To lintRecordCount
        Set curNode = node.Item(i - 1)
        For j = 1 To lintFieldCount
            Set curNode1 = curNode.selectSingleNode(lobjFields(j).name)
            If Not curNode1 Is Nothing Then              '有这个字段对应的值
                If lobjFields(j).type = 200 Then
                    lintLen = LenB(curNode1.nodeTypedValue)
                    '修改字符串字段的长度
                    If lobjFields(j).len < lintLen Then lobjFields(j).len = lintLen
                    lstrRecord(i, j) = Replace(curNode1.nodeTypedValue, Chr(10), Chr(13) & Chr(10), , , vbTextCompare)
                ElseIf lobjFields(j).type = 128 Then
                    lstrRecord(i, j) = lobjBase64.DecodeB(curNode1.nodeTypedValue)
                    lobjFields(j).len = UBound(lstrRecord(i, j)) - LBound(lstrRecord(i, j))
                Else
                    lstrRecord(i, j) = curNode1.nodeTypedValue
                End If
            Else
                lstrRecord(i, j) = Null
            End If
        Next
    Next
    '构建记录集的结构
    With lrecTemp
        For i = 1 To lintFieldCount
            If lobjFields(i).len > 0 Then
                If lobjFields(i).type = 128 Then
                    .Fields.Append lobjFields(i).name, lobjFields(i).type, lobjFields(i).len, adFldIsNullable + adFldLong
                Else
                    .Fields.Append lobjFields(i).name, lobjFields(i).type, lobjFields(i).len, adFldIsNullable
                End If
            Else
                .Fields.Append lobjFields(i).name, lobjFields(i).type, , adFldIsNullable
            End If
        Next
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open       ' 不需要连接对象。
    End With
    
    '构建记录集的数据
    For i = 1 To lintRecordCount
        lrecTemp.AddNew
        For j = 1 To lintFieldCount
            If lrecTemp(lobjFields(j).name).type = 7 Then
                lrecTemp(lobjFields(j).name) = IIf(IsNull(lstrRecord(i, j)), Null, Left(lstrRecord(i, j), 10) & " " & Mid(lstrRecord(i, j), 12, 8))
            Else
                lrecTemp(lobjFields(j).name) = lstrRecord(i, j)
            End If
        Next
        lrecTemp.Update
    Next
    
    '返回RecordSet。
    Set dafuncGetData = lrecTemp
    dafuncGetData.MoveFirst
    Set lrecTemp = Nothing
    Exit Function
errHandler:
    Dim llngErrNumber As Long
    Dim lstrErrDesc As String
    
    llngErrNumber = Err.Number
    lstrErrDesc = Err.Description
    If Err.Number = 6666 Then
        Err.Raise llngErrNumber, , lstrErrDesc
    Else
        Err.Raise llngErrNumber, , lstrErrDesc & vbCrLf & "SQL语句：" & paraSQL
    End If
    Exit Function
    Resume
End Function

' 功能：    开始一个事务。
' 输入：    无
' 输出：    无
' 返回：    无
' 注意事项：无
' 作者：    罗庆
' 创建时间：2001-3-5
Public Sub dasubBeginTran()
End Sub

' 功能：    开始一个事务。
' 输入：    无
' 输出：    无
' 返回：    无
' 注意事项：无
' 作者：    罗庆
' 创建时间：2001-3-5
Public Sub dasubCommitTran()
End Sub

' 功能：    开始一个事务。
' 输入：    无
' 输出：    无
' 返回：    无
' 注意事项：无
' 作者：    罗庆
' 创建时间：2001-3-5
Public Sub dasubRollBack()
End Sub

'功能：    设定数据库操作操时时间。
'输入：    paraQueryTimeOut       '操时的时间秒数
'输出：    无
'返回：    无
'注意事项：无
'作者：    罗庆
'创建时间：2001-3-5
Public Sub dasubSetQueryTimeout(paraQueryTimeOut As Integer)
End Sub

'程序产生错误时控制事务回滚。
Private Sub mcls消息传递_回滚事务(Cancel As Boolean)
End Sub

'功能：把枚举值集合构造成一个字典表的记录。
'输入：paraInfo as string/collection'枚举值列表/单值集合。
'返回：Recordset[编号，名称]。
'创建：2001-12-25。
Public Function dafunc构造字典记录(ByVal paraInfo As Variant) As Recordset
    Dim lcolInfo As Collection
    Dim recTemp As Recordset
    
    Dim strShape  As String
    Dim lstrTemp As String
    Dim lstrItem As String
    Dim i As Long
    
    On Error GoTo errHandler
    '若参数是枚举值串，转化为集合。
    If VarType(paraInfo) = vbString Then
        lstrTemp = paraInfo
        Set lcolInfo = New Collection
        If lstrTemp <> "" Then
            If Right(lstrTemp, 1) <> "," Then
                lstrTemp = lstrTemp & ","
            End If
            Do While InStr(lstrTemp, ",") > 0
                lstrItem = Left(lstrTemp, InStr(lstrTemp, ",") - 1)
                lcolInfo.Add lstrItem
                
                lstrTemp = Right(lstrTemp, Len(lstrTemp) - InStr(lstrTemp, ","))
            Loop
        End If
    Else
        Set lcolInfo = paraInfo
    End If
    
    Set recTemp = New Recordset
    strShape = "SHAPE APPEND New adVarChar(50) As 编号,New adVarChar(50) As 名称"
    
    recTemp.ActiveConnection = "provider=msdatashape;data provider=none;"
    recTemp.Open strShape, , adOpenStatic, adLockOptimistic
    
    '填充数据
    For i = 1 To lcolInfo.Count
        recTemp.AddNew Array("编号", "名称"), Array(Format(i, "00"), lcolInfo(i))
    Next
    If recTemp.RecordCount > 0 Then
        recTemp.MoveFirst
    End If
    Set dafunc构造字典记录 = recTemp
    Exit Function
errHandler:
    Err.Raise Err.Number, , Err.Description
    Exit Function
    Resume
End Function
'对命令串进行简单加密
Public Function funcConvert(pstrSource As String) As String
    Dim lintLen As Long, i As Long
    Dim lstrChr As String
    Dim j As Long
    Dim lstrDest() As String, llngDestNum As Long
    
    lintLen = Len(pstrSource)
    llngDestNum = lintLen / 1024
    If lintLen Mod 1024 <> 0 Then llngDestNum = llngDestNum + 1
    ReDim lstrDest(1 To llngDestNum)

    j = 1
    For i = 1 To lintLen
        lstrChr = Hex(Asc(Mid(pstrSource, i, 1)))
        lstrDest(j) = lstrDest(j) & IIf(Len(lstrChr) > 2, lstrChr, "00" & lstrChr)
        If i Mod 1024 = 0 Then j = j + 1
    Next
    For i = 1 To llngDestNum
        funcConvert = funcConvert + lstrDest(i)
    Next
    lstrChr = funcConvert
    funcConvert = ""
    For i = 1 To Len(lstrChr) / 2
        funcConvert = funcConvert & Hex(25 Xor CInt("&H" & Mid(lstrChr, 2 * (i - 1) + 1, 2)))
    Next
End Function

