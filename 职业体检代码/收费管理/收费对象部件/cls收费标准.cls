VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls收费标准"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'定义与属性相对应的变量

Private mstr收费标准名称 As String
Public 允许修改数量 As Integer
Public 过滤条件 As Integer

Private mcol收费项目 As Collection '收费项目编号,收费项目名称,单价,数量,key:收费项目编号

Public Property Get 收费标准名称() As String
    收费标准名称 = mstr收费标准名称
End Property

'设置属性编号的值
Public Property Let 收费标准名称(para收费标准名称 As String)
    Dim lobjRec As Object
    Dim lcolItem As Collection
    mstr收费标准名称 = para收费标准名称
    Set mcol收费项目 = New Collection
    Set lobjRec = dafuncGetData("select 编号,收费标准名称,收费项目编号 ,收费项目名称,单价 ,数量,允许修改数量,过滤条件 from 收费管理_收费标准信息视图 where 收费标准名称='" & 收费标准名称 & "'")
    If lobjRec.recordcount > 0 Then
        mstr收费标准名称 = lobjRec!收费标准名称
        允许修改数量 = lobjRec!允许修改数量
        过滤条件 = lobjRec!过滤条件
        
        Do While Not lobjRec.EOF
            Set lcolItem = New Collection
            lcolItem.Add lobjRec("收费项目编号").Value, "收费项目编号"
            lcolItem.Add lobjRec("收费项目名称").Value, "收费项目名称"
            lcolItem.Add lobjRec("单价").Value, "单价"
            lcolItem.Add lobjRec("数量").Value, "数量"
            mcol收费项目.Add lcolItem, lcolItem("收费项目编号")
            
            lobjRec.movenext
        Loop
        
    End If
    
End Property

Public Property Get 收费项目() As Collection
    Set 收费项目 = mcol收费项目
End Property


Public Sub sub添加项目(ByVal para收费项目编号 As String, ByVal para收费项目名称 As String, ByVal para单价 As Currency, ByVal para数量 As Integer)
    Dim lcolItem As Collection
    If sffunc判断集合键值是否存在(mcol收费项目, para收费项目编号) Then
        mcol收费项目.Remove para收费项目编号
    End If
    Set lcolItem = New Collection
    lcolItem.Add para收费项目编号, "收费项目编号"
    lcolItem.Add para收费项目名称, "收费项目名称"
    lcolItem.Add para单价, "单价"
    lcolItem.Add para数量, "数量"
    mcol收费项目.Add lcolItem, lcolItem("收费项目编号")
    
End Sub

Public Sub sub保存(ByVal para新标准名称 As String)
    Dim i As Long
    Dim lobjRec As Object
    On Error GoTo errHandler
    
    If mstr收费标准名称 <> "" Then
        dafuncGetData "delete 收费管理_收费标准信息表 where 收费标准名称='" & mstr收费标准名称 & "'"
       
    End If
    mstr收费标准名称 = para新标准名称
    For i = 1 To mcol收费项目.Count
        dafuncGetData "insert into 收费管理_收费标准信息表(收费标准名称,过滤条件,允许修改数量,收费项目编号,单价,数量) values(" _
            & "'" & mstr收费标准名称 & "'," & 过滤条件 & "," & 允许修改数量 & "," _
            & "'" & mcol收费项目(i)("收费项目编号") & "'," & mcol收费项目(i)("单价") & "," & mcol收费项目(i)("数量") & ")"
    Next
    
    Exit Sub
errHandler:
    sfsub错误处理 "收费对象部件", "cls收费标准", "sub保存", Err.Number, Err.Description, True
    Exit Sub
    Resume
End Sub


Public Sub sub删除标准()
    dafuncGetData "delete 收费管理_收费标准信息表 where 收费标准名称='" & mstr收费标准名称 & "'"
End Sub

Private Sub Class_Initialize()
    Set mcol收费项目 = New Collection
End Sub
