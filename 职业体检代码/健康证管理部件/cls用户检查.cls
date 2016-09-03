VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Type dataField                           '返回的数据列的定义
    name As String
    type As Integer
    len As Long
End Type

'在类初始化时获取消息传递的唯一实例。
Private Sub Class_Initialize()
End Sub

'在类释放时关闭数据库连接。
Private Sub Class_Terminate()
End Sub

' 功能：    执行查询语句。
' 输入：    paraSQL：SQL语句
' 输出：    无
' 返回：    保存返回值的RecordSet对象
' 注意事项：对于不需返回值的SQL语句如UPDATE语句也使用此方法执行，返回值不可用。
' 作者：    刘洪
' 创建时间：2004-3-25
Function dafuncGetDataNet(ByVal paraSQL As String) As ADODB.Recordset
    Dim lstrErr As String
    Dim lrecTemp As ADODB.Recordset   'SQL语句返回的RecordSet
    Dim lstrResult As String, lobjDoc As Object
    Dim lobjHttp As Object, lstrWebserviceURL As String
    Dim lstrRequest As String, lintStatus As Integer
    Dim lstrProcName As String
    Dim lstrSql As String
    
'    If pstrDSN = "" Then
'        Err.Raise 6666, , "没有设置中心服务器的域名！请先调用dasubInitialize进行设置。"
'        Exit Function
'    End If
    
    On Error GoTo errhandler
    
    Set lobjHttp = CreateObject("MSXML2.XMLHTTP")
    Set lobjDoc = CreateObject("MSXML.DOMDocument")
    lstrWebserviceURL = "http://www.fcstech.com.cn/UserDataAccess/clsDataAccess.asmx/funcGetRecordset"
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
    
    Set lrecTemp = New ADODB.Recordset
    
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
        Set dafuncGetDataNet = lrecTemp
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
    Set dafuncGetDataNet = lrecTemp
    dafuncGetDataNet.MoveFirst
    Set lrecTemp = Nothing
    Exit Function
errhandler:
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
'检查当前是否已经超过老版卡号最后使用期限，若是，返回true，否则返回false
Public Function funcCheckExpireDate() As Boolean
    Dim lobjRec As Recordset
    Dim lstrExpireDate As String
    Dim lstrNo As String
    
    On Error GoTo errHandle
    
    funcCheckExpireDate = True
    
    '“已存在键值”为用户编号，“可用键值1”为老版卡号最后使用期限
    Set lobjRec = dafuncGetData("select 已存在键值,可用键值1 from 系统管理_树形控件已存在键值表")
    If lobjRec.RecordCount = 0 Then     '没有认证记录
        funcCheckExpireDate = True
        Exit Function
    End If
    lstrNo = func解密(lobjRec(0))
    If Not IsNull(lobjRec(1)) Then
        lstrExpireDate = func解密(lobjRec(1))
        If Date > CDate(lstrExpireDate) Then
            funcCheckExpireDate = True
        Else
            funcCheckExpireDate = False
            Exit Function
        End If
    End If
    '没有记录最后期限或已经过期，从公司服务器获取最新的期限信息
    Set lobjRec = dafuncGetDataNet("exec getJkzMaxNo " + lstrNo)
    If IsNull(lobjRec("ExpireDate")) Then       '未记录期限，视为不加控制
        funcCheckExpireDate = False
    Else
        dafuncGetData "update 系统管理_树形控件已存在键值表 set 可用键值1=N'" & func加密(Format(lobjRec("ExpireDate"), "yyyy-mm-dd")) & "'"
        If Date > lobjRec("ExpireDate") Then
            funcCheckExpireDate = True
        Else
            funcCheckExpireDate = False
        End If
    End If
    Exit Function
errHandle:
    Err.Raise Err, , "获取用户认证信息错误，请检查本电脑是否能访问互联网。"
End Function
'检查指定卡号是否是该用户的卡，并且是否超过公司的打印范围，若没有超过，返回true，否则返回false
Public Function funcCheckMaxNo(paraCardno As String) As Boolean
    Dim lobjRec As Recordset
    Dim lstrMaxNo As String, lstrNo As String
    
    On Error GoTo errHandle
    
    '“已存在键值”为用户编号，“可用键值2”为最大卡号
    Set lobjRec = dafuncGetData("select 已存在键值,可用键值2 from 系统管理_树形控件已存在键值表")
    If lobjRec.RecordCount = 0 Then     '没有认证记录
        funcCheckMaxNo = False
        Exit Function
    End If
    lstrNo = func解密(lobjRec(0))
    If Left(paraCardno, 4) <> CInt(lstrNo) Then       '卡号前缀与用户编号不符
        funcCheckMaxNo = False
        Exit Function
    End If
    If Not IsNull(lobjRec(1)) Then
        lstrMaxNo = func解密(lobjRec(1))
        If CInt(Mid(paraCardno, 5, 6)) <= CInt(lstrMaxNo) Then
            funcCheckMaxNo = True
            Exit Function
        End If
    End If
    '没有记录最大卡号或已经超过，从公司服务器获取最新的期限信息
    Set lobjRec = dafuncGetDataNet("exec getJkzMaxNo " + lstrNo)
    If IsNull(lobjRec("MaxNo")) Then       '未记录最大卡号，视为不加控制
        funcCheckMaxNo = True
    Else
        dafuncGetData "update 系统管理_树形控件已存在键值表 set 可用键值2=N'" & func加密(CStr(lobjRec("MaxNo"))) & "'"
        If CInt(Mid(paraCardno, 5, 6)) <= lobjRec("MaxNo") Then
            funcCheckMaxNo = True
        Else
            funcCheckMaxNo = False
        End If
    End If
    Exit Function
errHandle:
    Err.Raise Err, , "获取用户认证信息错误，请检查本电脑是否能访问互联网。"
End Function
Public Sub subSetCurNo(paraNo As String)
    Dim lobjRec As Recordset
    Dim lstrNo As String
    
    On Error Resume Next
    
    '“已存在键值”为用户编号
    Set lobjRec = dafuncGetData("select 已存在键值 from 系统管理_树形控件已存在键值表")
    If lobjRec.RecordCount = 0 Then     '没有认证记录
        Exit Sub
    End If
    lstrNo = func解密(lobjRec(0))
    dafuncGetDataNet "exec setJkzCurNo " & lstrNo & "," & CInt(Right(paraNo, 6))
End Sub
'功能：将用户的密码加密
'输入：用户输入的密码
'输出：无
'返回：加密后的密码
'注意事项：无
'作者：王晓华
'创建时间：2001-03-22
Private Function func加密(ByVal para原来密码 As String) As String
    On Error GoTo errHandle
    Dim lintTemp As Integer
    Dim lint取余 As Integer
    Dim i As Integer
    Dim j As Integer
    Dim lstrTemp As String
    lintTemp = Len(para原来密码)   '密码长度
    i = lintTemp \ 4              '密码以四位加密
    lint取余 = lintTemp - i * 4   '不足四位补a
    Select Case lint取余
        Case 1
        para原来密码 = para原来密码 & "aaa"
        i = i + 1
        Case 2
        para原来密码 = para原来密码 & "aa"
        i = i + 1
        Case 3
        para原来密码 = para原来密码 & "a"
        i = i + 1
    End Select
    Dim lobjEncrypt As New clsDataEncrypt
    For j = 1 To i
        If lobjEncrypt.funcDESEncrypt(Mid(para原来密码, 4 * (j - 1) + 1, 4), lstrTemp) = False Then Exit Function
        func加密 = func加密 & lstrTemp
    Next j
    Exit Function
errHandle:
    Call sfsub错误处理("通用对象", "cls用户管理", "func加密", Err.Number, Err.Description, True)
End Function
Private Function func解密(ByVal para原来密码 As String) As String
    On Error GoTo errHandle
    Dim lintTemp As Integer
    Dim lint取余 As Integer
    Dim i As Integer
    Dim j As Integer, k As Integer
    Dim lstrTemp As String
    lintTemp = Len(para原来密码)   '密码长度
    If lintTemp Mod 4 > 0 Then      '如果长度不为4的倍数，退出
        func解密 = ""
        Exit Function
    End If
    
    i = lintTemp \ 4              '密码以四位加密
    
    Dim lobjEncrypt As New clsDataEncrypt
    For j = 1 To i
        If lobjEncrypt.funcDESDecrypt(Mid(para原来密码, 4 * (j - 1) + 1, 4), lstrTemp) = False Then Exit Function
        k = InStr(lstrTemp, "a")
        If k = 0 Then k = 5
        func解密 = func解密 & Left(lstrTemp, k - 1)
    Next j
    Exit Function
errHandle:
    Call sfsub错误处理("通用对象", "cls用户管理", "func加密", Err.Number, Err.Description, True)
End Function
'对命令串进行简单加密
Function funcConvert(pstrSource As String) As String
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
