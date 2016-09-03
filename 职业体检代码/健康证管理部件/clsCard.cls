VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
   
Public Col As Long

Private mstrSN As String '系统编号。
Public CN As String '健康证号。
Public Name As String '姓名。
Public Sex As String '性别。
Public Age As String
Public EmCertCatalog As String '卫生种类。
Public ExaminationDate As String  '发证日期。
Public ExaminUnit As String   '体检单位。
Public ExaminResult As String '体检结论。

Public MZ As String '民族。
Public UnitName As String '单位名称。

Public Photo As StdPicture '照片。


'系统编号。
Public Property Get SN() As String
    SN = mstrSN
End Property

Public Property Let SN(ByVal vNewValue As String)
    Dim lobjRec As Object
    Dim lobjBase As Object
    mstrSN = vNewValue
    '获取健康证其他信息。
    Set lobjBase = dafuncGetData("select * from 健康证管理_办证申请信息表 where 系统编号='" & vNewValue & "'")
    If lobjBase.RecordCount > 0 Then
        Name = IIf(IsNull(lobjBase!姓名), "", lobjBase!姓名)
        Sex = lobjBase!性别
        Age = IIf(IsNull(lobjBase!年龄), "", lobjBase!年龄)
        '卫生种类。
        EmCertCatalog = lobjBase!种类
        ExaminationDate = Format(lobjBase!体检日期, "yyyy-mm-dd")
        ExaminUnit = IIf(IsNull(lobjBase!体检单位), "", lobjBase!体检单位)
        ExaminResult = IIf(IsNull(lobjBase!体检结论), "", lobjBase!体检结论)
        MZ = IIf(IsNull(lobjBase!民族), "", lobjBase!民族)
        
        UnitName = IIf(IsNull(lobjBase!单位名称), "", lobjBase!单位名称)
        
        '获取照片。
        If IIf(IsNull(lobjBase!照片编号), "", lobjBase!照片编号) <> "" Then
            Set Photo = pmfunc获取图片(lobjBase!照片编号, "健康证管理")
        Else
            Set Photo = LoadPicture()
        End If
        
    End If
End Property
