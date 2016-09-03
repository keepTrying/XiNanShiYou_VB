VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMedicalProjt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function subDelParent(ByVal paraProj As String)
    Dim lstrsql As String
    Dim lstr科室 As String
    Select Case paraProj
        Case "01"
            lstr科室 = "五官科"
        Case "02"
            lstr科室 = "内科"
        Case "03"
            lstr科室 = "外科"
        Case "04"
            lstr科室 = "妇科"
        Case "05"
            lstr科室 = "检验科"
        Case ""
            Exit Function
    End Select
    lstrsql = "delete 职业病体检_结果信息_" & lstr科室 & " where left(体检项目,2)=" & lstr科室
    dafuncGetData lstrsql
    Exit Function
errHandler:
   sfsub错误处理 "职业病史录入", "clsmedicalprojt", "public sub save体检项目", Err.Number, Err.Description, False
End Function

Public Function subAddParent(ByVal paraProj As String)
    Dim lstrsql As String
    Dim lstr科室 As String
    Select Case paraProj
        Case "01"
            lstr科室 = "五官科"
        Case "02"
            lstr科室 = "内科"
        Case "03"
            lstr科室 = "外科"
        Case "04"
            lstr科室 = "妇科"
        Case "05"
            lstr科室 = "检验科"
        Case ""
            Exit Function
    End Select
    lstrsql = "delete 职业病体检_结果信息_" & lstr科室 & " where left(体检项目,2)=" & lstr科室
    dafuncGetData lstrsql
    Exit Function
errHandler:
   sfsub错误处理 "职业病史录入", "clsmedicalprojt", "public sub save体检项目", Err.Number, Err.Description, False
End Function

