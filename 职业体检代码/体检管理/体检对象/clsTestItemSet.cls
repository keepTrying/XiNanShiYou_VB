VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTestItemSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public 体检项目编码 As String
Public 体检大类 As Long
Public 属性 As String '常规/化验。

Private Sub Class_Initialize()
    体检大类 = -1
    体检项目编码 = ""
    属性 = ""
End Sub

'功能：获取所有体检项目(或指定体检大类、体检项目编码）的信息
'输入：无。
'返回：Recordset[编码，名称，默认值，可选值，属性]。
'作者：杨春
Public Property Get 体检项目() As Object
    Dim lstrFilter As String
    
    On Error GoTo errHandler
    '拼条件。
    If 体检项目编码 <> "" Then lstrFilter = "编码='" & 体检项目编码 & "'"
    If 体检大类 <> -1 Then lstrFilter = lstrFilter & IIf(lstrFilter = "", "", " and ") & "体检大类=" & 体检大类
    If 属性 <> "" Then lstrFilter = lstrFilter & IIf(lstrFilter = "", "", " and ") & "属性='" & 属性 & "'"
    
    '执行sql语句。
    Set 体检项目 = dafuncGetData("select * from 体检管理_体检项目设置表" & IIf(lstrFilter = "", "", " where " & lstrFilter))
    
    Exit Property
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检对象部件", "ClsTestItemSet", "Property Get 体检项目", 6666, lstrError, False
    
End Property

