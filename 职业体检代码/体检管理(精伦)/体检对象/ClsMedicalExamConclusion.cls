VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsMedicalExamConclusion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private mlngID As Long     '体检结论id。
Private mstr名称 As String '体检结论名称。


Public Property Get ID() As Long
    ID = mlngID
End Property

'功能：给本对象的关键属性体检结论ID赋值，从而通过字典对象获取体检结论名称。
'作者：杨春
Public Property Let ID(ByVal vNewValue As Long)
    Dim lobjRec As Object
    
    On Error GoTo errHandler
    mlngID = vNewValue
    '获取体检结论名称。
    Set lobjRec = dafuncGetData("select 体检结论名称 from 体检管理_体检结论判断条件数据库 where 体检结论ID=" & mlngID)
    If lobjRec.recordcount > 0 Then
        mstr名称 = IIf(IsNull(lobjRec("体检结论名称")), "", lobjRec("体检结论名称").Value)
    Else
        mstr名称 = ""
    End If
    
    Exit Property
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "ClsMedicalExamConclusion", "Property Let ID", 6666, lstrError, True
    Exit Property
    Resume
End Property
'功能：获取当前体检结论的所有判断条件。
'作者：杨春
Property Get 所有判断条件() As Collection
    Dim lobjRec As Object
    Dim lobjConclusionFilter As ClsConclusionFilter
    Dim lcolInfo As Collection
    
    On Error GoTo errHandler
    Set lcolInfo = New Collection
    
    '获取库中当前体检结论id所有判断条件分组。
    Set lobjRec = dafuncGetData("select distinct 编号 from 体检管理_体检结论判断条件表 where 体检结论=" & mlngID)
    Do While Not lobjRec.EOF
        '创建条件分组对象。
        Set lobjConclusionFilter = New ClsConclusionFilter
        
        '设置对象的属性。
        lobjConclusionFilter.ID = mlngID
        lobjConclusionFilter.编号 = lobjRec("编号")
        lcolInfo.Add lobjConclusionFilter, Trim(Str(lobjRec("编号")))
        
        lobjRec.MoveNext
    Loop
    Set 所有判断条件 = lcolInfo
    Exit Property
errHandler:
    sfsub错误处理 "体检对象部件", "ClsMedicalExamConclusion", "Property Get 所有判断条件", Err.Number, Err.Description, True
    Exit Property
    Resume
End Property
'功能：获取当前体检结论的指定分组编号的判断条件。
'作者：杨春
Property Get 判断条件(ByVal para编号 As String) As ClsConclusionFilter
    Dim lobjConclusionFilter As ClsConclusionFilter
    
    On Error GoTo errHandler
    '创建条件分组对象。
    Set lobjConclusionFilter = New ClsConclusionFilter
    
    '设置返回对象的属性。
    lobjConclusionFilter.ID = mlngID
    lobjConclusionFilter.编号 = para编号
    
    '返回。
    Set 判断条件 = lobjConclusionFilter
    
    Exit Property
errHandler:
    sfsub错误处理 "体检对象部件", "ClsMedicalExamConclusion", "Property Get 判断条件", Err.Number, Err.Description, True
End Property

'功能：从库中删除当前体检结论的指定分组编号的所有判断子条件。
'作者：杨春
Public Sub Sub删除条件分组(ByVal para编号 As String)
    On Error GoTo errHandler

    dafuncGetData "delete 体检管理_体检结论判断条件表 where 体检结论=" & mlngID & " and 编号='" & para编号 & "'"
    
    Exit Sub
errHandler:
    sfsub错误处理 "体检对象部件", "ClsMedicalExamConclusion", "Sub删除条件分组", Err.Number, Err.Description, True
End Sub

'功能：判断指定系统编号的体检记录是否可下当前结论。
'作者：杨春
Public Function Func判断是否可下本结论(ByVal para系统编号 As String) As Boolean
    Dim lobjRec As Object
    Dim lobjConclusionFilter As ClsConclusionFilter
    
    On Error GoTo errHandler
    Func判断是否可下本结论 = False
    
    '获取库中当前体检结论id所有判断条件分组。
    Set lobjRec = dafuncGetData("select distinct 编号,体检结论 from 体检管理_体检结论判断条件表 where 体检结论=" & mlngID)
    Do While Not lobjRec.EOF
        '创建条件分组对象。
        Set lobjConclusionFilter = New ClsConclusionFilter
        
        '设置对象的属性。
        lobjConclusionFilter.ID = lobjRec("体检结论")
        lobjConclusionFilter.编号 = lobjRec("编号")
        
        If lobjConclusionFilter.Func判断是否满足条件(para系统编号) Then
            Func判断是否可下本结论 = True
            Exit Do
        End If
        lobjRec.MoveNext
    Loop
    Exit Function
errHandler:
    sfsub错误处理 "体检对象部件", "ClsMedicalExamConclusion", "Func判断是否可下本结论", Err.Number, Err.Description, True
    Exit Function
    Resume
End Function
