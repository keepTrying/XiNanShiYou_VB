VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls收费项目"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'&  ============================| cls收费项目 |====================================


Option Explicit

Private mstr收费项目编号 As String
Public 收费项目名称 As String
Public 计量单位 As String
Public 助记符 As String
Public 单价 As Currency
Public 最小单价 As Currency
Public 最大单价 As Currency
Public 票据类型编号 As Integer

Public Property Let 收费项目编号(ByVal vNewValue As String)
    mstr收费项目编号 = vNewValue
    Dim lobjRec As Object
    If mstr收费项目编号 <> "" Then
        Set lobjRec = dafuncGetData("select * from 收费管理_收费项目字典表 where 收费项目编号='" & vNewValue & "'")
        If lobjRec.recordcount > 0 Then
            收费项目名称 = lobjRec!收费项目名称
            计量单位 = IIf(IsNull(lobjRec!计量单位), "", lobjRec!计量单位)
            助记符 = IIf(IsNull(lobjRec!助记符), "", lobjRec!助记符)
            单价 = IIf(IsNull(lobjRec!单价), 0, lobjRec!单价)
            最小单价 = IIf(IsNull(lobjRec!最小单价), 0, lobjRec!最小单价)
            最大单价 = IIf(IsNull(lobjRec!最大单价), 0, lobjRec!最大单价)
            票据类型编号 = IIf(IsNull(lobjRec!票据类型编号), 0, lobjRec!票据类型编号)
        
        End If
    End If
End Property
Public Property Get 收费项目编号() As String
    收费项目编号 = mstr收费项目编号
End Property

Public Sub sub保存(ByVal para上级编号 As String)
    Dim lstrNo As String
    
    If mstr收费项目编号 = "" Then
        '新增项目。生成项目编号。
        Dim lobjRec As Object
        Set lobjRec = dafuncGetData("select max(right(收费项目编号,3)) from 收费管理_收费项目字典表 where 收费项目编号 like '" & para上级编号 & "%' and len(收费项目编号)=" & Len(para上级编号) + 3)
        lstrNo = Format(Val(IIf(IsNull(lobjRec(0)), "0", lobjRec(0))) + 1, "000")
        mstr收费项目编号 = para上级编号 & lstrNo
        
        '插入。
        dafuncGetData "insert into 收费管理_收费项目字典表(收费项目编号,收费项目名称,计量单位,助记符,单价,最小单价,最大单价,票据类型编号) values(" _
                        & "'" & mstr收费项目编号 & "','" & 收费项目名称 & "','" & 计量单位 & "','" & 助记符 & "'," & 单价 & "," & 最小单价 & "," & 最大单价 & "," & 票据类型编号 & ")"
    Else
        '修改。
        dafuncGetData "update 收费管理_收费项目字典表 set 收费项目名称='" & 收费项目名称 & "',计量单位='" & 计量单位 _
                        & "',助记符='" & 助记符 & "',单价=" & 单价 & ",最小单价=" & 最小单价 & ",最大单价=" & 最大单价 & ",票据类型编号=" & 票据类型编号 _
                        & " where 收费项目编号='" & mstr收费项目编号 & "'"
    End If
End Sub
