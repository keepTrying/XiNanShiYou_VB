VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLocalWorkBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private mcol单位名称 As Collection '记录体检登记当日已录入的单位名称。
Private mcol单位编号 As Collection '单位申请编号。

Private mstr单位名称列表 As String

Private mstr工作簿文件 As String

'当日体检工作记忆簿文件格式：
'       [基本信息]
'            工作日期=
'       [体检登记]
'            单位名称=
'           …
'作者：杨春
Private Sub Class_Initialize()
    Dim lstrList As String '所有单位名称列表。
    Dim lstrItem As String
    Dim lstrName As String '单位名称。
    Dim lstrNo As String   '单位申请编号。
    Dim lstr  As String
    Dim i As Long
    
    On Error Resume Next
    Set mcol单位名称 = New Collection
    Set mcol单位编号 = New Collection
    
    mstr单位名称列表 = ""
    
    '判断工作记忆簿是否存在，若不存在，创建它。。
    If Dir("c:\temp", vbDirectory) = "" Then
        MkDir "c:\temp"
    End If
    mstr工作簿文件 = "c:\temp\当日体检工作记忆簿.ini"
    
    On Error GoTo errHandler
    If Dir(mstr工作簿文件) <> "" Then
        '修改：2002-1-28（判断“当日体检工作记忆簿.ini”的日期，若<当天，删除该文件）。
        Dim fs As FileSystemObject
        Dim lobjFile As File
        On Error Resume Next
        Set fs = CreateObject("Scripting.FileSystemObject")
        Set lobjFile = fs.GetFile(mstr工作簿文件)
        If Format(lobjFile.DateLastModified, "yyyy-mm-dd") < Format(Date, "yyyy-mm-dd") Then
            lobjFile.Delete True
            Set lobjFile = Nothing
            Set fs = Nothing
            Exit Sub
        Else
            Set lobjFile = Nothing
            Set fs = Nothing
        End If
        On Error GoTo errHandler
        
        '从本地临时路径：c:\temp\当日体检工作记忆簿.ini”中获取当前工作站的当日需要记忆的信息。
        '目前只有体检登记时录入过的单位名称。
            
        '获取体检登记时录入的单位名称。
        lstr = Space(600)
        i = GetPrivateProfileString("体检登记", "单位名称", "", lstr, 600, mstr工作簿文件)
        mstr单位名称列表 = Trim(lstr)
        If mstr单位名称列表 <> "" Then
            If Asc(Right(mstr单位名称列表, 1)) = 0 Then
                mstr单位名称列表 = Left(mstr单位名称列表, Len(mstr单位名称列表) - 1)
            End If
        End If
        '把单位串分解到集合中。
        lstrList = mstr单位名称列表
        If lstrList <> "" And Right(lstrList, 1) <> "," Then lstrList = lstrList & ","
        Do While InStr(1, lstrList, ",") > 0
            lstrItem = Trim(Left(lstrList, InStr(1, lstrList, ",") - 1))
            If lstrItem <> "" Then
                If InStr(lstrItem, "|") > 0 Then
                    lstrNo = Left(lstrItem, InStr(lstrItem, "|") - 1)
                    lstrName = Right(lstrItem, Len(lstrItem) - InStr(lstrItem, "|"))
                Else
                    lstrNo = ""
                    lstrName = lstrItem
                End If
                mcol单位名称.Add lstrName, lstrName
                mcol单位编号.Add lstrNo, lstrName
            End If
            lstrList = Right(lstrList, Len(lstrList) - InStr(1, lstrList, ","))
        Loop
        
    End If
    
    Exit Sub
errHandler:
    '忽略错误。
    Exit Sub
    Resume
End Sub

'功能：获取本地工作记忆簿中所有的单位名称。
'返回：单位名称集 as collection。
'作者：杨春
Public Property Get 单位名称集() As Collection
    Dim lstr单位名称 As Variant
    On Error GoTo errHandler
    
    Set 单位名称集 = New Collection
    For Each lstr单位名称 In mcol单位名称
        单位名称集.Add lstr单位名称, lstr单位名称
    Next
    
    Exit Property
errHandler:
    'sfsub错误处理 "体检对象部件", "clsLocalWorkBook", "Property Get 单位名称集", Err.Number, Err.Description, True
    
End Property
'返回：返回指定单位名称的单位编号。
'作者：杨春
Public Property Get 单位编号(ByVal paraName As String) As String
    On Error GoTo errHandler
    单位编号 = mcol单位编号(paraName)
    Exit Property
errHandler:
    单位编号 = ""
    
End Property

'功能：向本地工作记忆簿添加一个单位名称。
'输入：单位申请编号|单位名称
'作者：杨春
Public Sub sub增加单位名称(ByVal para单位名称 As String)
    Dim lstrOld As String
    Dim lstrName As String
    Dim lstrNo As String
    Dim i As Long
    
    On Error GoTo errHandler
    para单位名称 = Trim(para单位名称)
    
    If para单位名称 <> "" Then
        If InStr(para单位名称, "|") > 0 Then
            lstrNo = Trim(Left(para单位名称, InStr(para单位名称, "|") - 1))
            lstrName = Trim(Right(para单位名称, Len(para单位名称) - InStr(para单位名称, "|")))
        Else
            lstrName = para单位名称
            lstrNo = ""
        End If
        
        '首先判断该单位名称是否已在“mcol单位名称”中存在。
        If Not sffunc判断集合键值是否存在(mcol单位名称, lstrName) Then
            lstrOld = ""
            '若该单位名称不存在，加入集合“mcol单位名称”中。
            mcol单位名称.Add lstrName, lstrName
            mcol单位编号.Add lstrNo, lstrName
        Else
            '已存在，判断单位申请编号是否相同。
            lstrOld = Trim(mcol单位编号(lstrName))
            If lstrOld <> lstrNo Then
                '修改了单位申请编号。
                mcol单位编号.Remove lstrName
                mcol单位编号.Add lstrNo, lstrName
            End If
        End If
        
        '只保存最近的15个单位。
        Do While mcol单位名称.Count > 15
            mcol单位名称.Remove 1
            mcol单位编号.Remove 1
        Loop
        mstr单位名称列表 = ""
        For i = 1 To mcol单位名称.Count
            mstr单位名称列表 = Trim(mstr单位名称列表) & mcol单位编号(mcol单位名称(i)) & "|" & mcol单位名称(i) & ","
        Next
        If mstr单位名称列表 <> "" Then mstr单位名称列表 = Left(mstr单位名称列表, Len(mstr单位名称列表) - 1)
        
        '重写本地工作簿文件。
        WritePrivateProfileString "体检登记", "单位名称", mstr单位名称列表, mstr工作簿文件
    End If
    Exit Sub
errHandler:
    'sfsub错误处理 "体检对象部件", "clsLocalWorkBook", "sub增加单位名称", Err.Number, Err.Description, True
End Sub

