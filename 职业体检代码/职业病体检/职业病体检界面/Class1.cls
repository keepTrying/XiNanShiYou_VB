VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private mstr单位申请编号 As String

Public Property Let 单位信息申请(danwei As String)
    mstr单位申请编号 = danwei
End Property
Public Property Get 单位信息申请() As Object
    Dim lstrSQL语句 As String
    Dim lstrTmp As String
    On Error GoTo errHandler
    
    lstrSQL语句 = "select * from 单位档案_单位基本信息表 where 申请编号='" & mstr单位申请编号 & "'"
    Set 单位信息编号 = dafuncGetData(lstrSQL语句)
    Exit Property
errHandler:
    sfsub错误处理 "体检对象部件", "ClsMedicalExamTemplateSet", "Get 经济性质", Err.Number, Err.Description, True
End Property
