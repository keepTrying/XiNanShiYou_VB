VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsManageTransmission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'属性变量。


Public Function func获取体检人员名单(ByVal paraStartDate As String, ByVal paraEndDate As String, ByVal paraUnit As String, _
                ByVal paraSheet As String, ByVal paraStartCode As String, ByVal paraEndCode As String) As Object
    Set func获取体检人员名单 = dafuncGetData("exec 体检管理_获取导出体检人员信息 '" & paraStartCode & "','" & paraEndCode & "','" _
                        & paraUnit & "','" & paraSheet & "','" & paraStartCode & "','" & paraEndCode & "'")
End Function


Public Sub sub导出体检人员名单(ByVal paraFileName As String, ByVal paraID As String)
    Dim lobjFso As Object
    Dim lobjFile As Object
    Dim lobjBase As Object
    Dim lobjDetail As Object
    Dim llngSourceFile As Long '原文件号。
    Dim lstrLine As String
    Dim i As Long
    On Error GoTo errHandler
    
    Set lobjFso = CreateObject("Scripting.FileSystemObject")
    
    Set lobjBase = dafuncGetData("select * from temp_体检基本信息 where ID='" & paraID & "' order by 系统编号")
    
    '设置目的文件的属性为可读可写。
    If Dir(paraFileName) <> "" Then
        Set lobjFile = lobjFso.GetFile(paraFileName)
        lobjFile.Attributes = 0
    End If
    llngSourceFile = FreeFile
    Open paraFileName For Output As #llngSourceFile
    
    Do While Not lobjBase.EOF
        '基本信息
        lstrLine = "D1"
        For i = 0 To lobjBase.fields.Count - 2
            lstrLine = lstrLine & "," & IIf(IsNull(lobjBase(i).Value), "", lobjBase(i).Value)
        Next
        lstrLine = lstrLine & ",,,,,,," '另外加7个逗号
        Print #llngSourceFile, lstrLine
        
        '结果信息。
        Set lobjDetail = dafuncGetData("select * FROM temp_体检结果信息 where ID='" & paraID & "' and 系统编号='" & lobjBase("系统编号") & "'")
        lstrLine = "D2," & IIf(IsNull(lobjBase("试管编号")), "", lobjBase("试管编号"))
        Do While Not lobjDetail.EOF
            lstrLine = lstrLine & "," & lobjDetail("体检项目名称") & "," & IIf(IsNull(lobjDetail("体检结果")), "", lobjDetail("体检结果"))
            lobjDetail.movenext
        Loop
        lstrLine = lstrLine & ","
        Print #llngSourceFile, lstrLine
        
        lobjBase.movenext
    Loop
    
    Close #llngSourceFile
    Exit Sub
errHandler:
    Close #llngSourceFile
    sfsub错误处理 "体检数据导入导出", "clsManageTransmission", "sub导出体检人员名单", Err.Number, Err.Description, True
    Exit Sub
    Resume
End Sub

'读入体检结果到临时表 temp_体检结果信息 中。
'返回：temp_体检结果信息 表中的ID。
Public Function func读入体检结果(ByVal paraFileName As String) As String
    Dim llngSourceFile As Long '原文件号。
    Dim lstrLine1 As String
    Dim lstrLine2 As String
    Dim lobjRec As Object
    Dim i As Long
    Dim lstrID As String
    Dim lstrSysNo As String
    Dim larrValue
    Dim lstrItem As String
    Dim lstrValue As String
    Dim lstrFields As String
    Dim lintNum As Integer
    On Error GoTo errHandler
    
    func读入体检结果 = ""
    
    '获取ID。
    Set lobjRec = dafuncGetData("select convert(varchar(30),getdate(),126)")
    lstrID = lobjRec(0)
    
    '获取临时表 temp_体检基本信息 的结构
    lstrFields = ""
    Set lobjRec = dafuncGetData("select * from temp_体检基本信息 where 1=0")
    For i = 0 To lobjRec.fields.Count - 1
        lstrFields = lstrFields & lobjRec(i).Name & ","
    Next
    lstrFields = Left(lstrFields, Len(lstrFields) - 1)
    lintNum = lobjRec.fields.Count
    
    llngSourceFile = FreeFile
    
    Open paraFileName For Input As #llngSourceFile
    Do While Not EOF(llngSourceFile)
        Line Input #llngSourceFile, lstrLine1
        lstrSysNo = ""
        If lstrLine1 <> "" And InStr(lstrLine1, ",") > 0 Then
            '去掉D1,
            lstrLine1 = Right(lstrLine1, Len(lstrLine1) - InStr(lstrLine1, ","))
            
            '去掉后面7个,
            lstrLine1 = Left(lstrLine1, Len(lstrLine1) - 7)
            
            '把所有字段放入一个数组。
            larrValue = Split(lstrLine1, ",")
            '最后一项是系统编号。
            lstrSysNo = larrValue(UBound(larrValue))
            
            '保存基本信息。
            lstrLine1 = "insert into temp_体检基本信息(" & lstrFields & ") values("
            For i = 0 To UBound(larrValue)
                lstrLine1 = lstrLine1 & "'" & larrValue(i) & "',"
            Next
            lstrLine1 = lstrLine1 & "'" & lstrID & "')"
            dafuncGetData lstrLine1
        End If
        If EOF(llngSourceFile) Then
            Err.Raise 6666, , "文件格式不正确，总行数不是偶数！"
        End If
        Line Input #llngSourceFile, lstrLine2
        
        '插入体检结果到临时表
        If lstrLine2 <> "" And lstrSysNo <> "" Then
            '去掉D2,试管编号
            lstrLine2 = Right(lstrLine2, Len(lstrLine2) - InStr(lstrLine2, ","))
            lstrLine2 = Right(lstrLine2, Len(lstrLine2) - InStr(lstrLine2, ","))
            '去掉最后那个,
            lstrLine2 = Left(lstrLine2, Len(lstrLine2) - 1)
            
            Do While lstrLine2 <> "" And InStr(lstrLine2, ",") > 0
                lstrItem = Left(lstrLine2, InStr(lstrLine2, ",") - 1)
                lstrLine2 = Right(lstrLine2, Len(lstrLine2) - InStr(lstrLine2, ","))
                If InStr(lstrLine2, ",") > 0 Then
                    lstrValue = Left(lstrLine2, InStr(lstrLine2, ",") - 1)
                    lstrLine2 = Right(lstrLine2, Len(lstrLine2) - InStr(lstrLine2, ","))
                Else
                    lstrValue = lstrLine2
                End If
                '插入。
                If lstrItem <> "" Then
                    dafuncGetData "insert into temp_体检结果信息(ID,系统编号,体检项目名称,体检结果) values('" & lstrID & "','" & lstrSysNo & "','" & lstrItem & "','" & lstrValue & "')"
                End If
            Loop
            
        End If
    Loop
    Close #llngSourceFile
    func读入体检结果 = lstrID
    Exit Function
errHandler:
    If llngSourceFile <> 0 Then
        Close #llngSourceFile
    End If
    sfsub错误处理 "体检数据导入导出", "clsManageTransmission", "func读入体检结果", Err.Number, Err.Description, True
    Exit Function
    Resume
End Function
