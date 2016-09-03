VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsShowPhoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Property Let 系统编号(ByVal vNewValue As String)
    mstr系统编号 = vNewValue
End Property
Public Property Get 像片() As StdPicture
    On Error Resume Next
    Dim mpicPhoto As Object
    '获取照片
'    If mpicPhoto Is Nothing Then
        Set mpicPhoto = pmfunc获取图片资料(mstr系统编号, 1)
'    End If
    
    Set 像片 = mpicPhoto
End Property
Public Function pmfunc获取图片资料(ByVal 系统编号 As String, ByVal 次数 As Integer) As Object
    Dim strTemp As String
    Dim re As Object
    Set re = dafuncGetData("select 图片 from 职业病体检_复查资料信息表 where 系统编号='" & 系统编号 & "' and 上传次数='" & 次数 & "'")
    re.Type = adTypeBinary
'    Set re = dafuncGetData("select 图片 from 系统管理_系统图片管理表 where 图片编号='" & 系统编号 & "' and 子系统名='" & 系统 & "'")
    pmfunc获取图片资料 = re
End Function

