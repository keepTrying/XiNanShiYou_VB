VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLocalConfigure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private mlng工作模式 As PENUM_工作模式
Private mstrExcel文件 As String
Private mstr内部导入文件 As String
Private mstr内部导出文件 As String

'本地工作站配置文件格式:
'        [工作模式]
'            工作模式 = 0
'        [数据导入导出]
'            Excel文件 = "c:\temp\a.xls"
'            内部导入文件 = "c:\temp\b.dat"
'            内部导出文件 = "c:\temp\c.dat"

Private Sub Class_Initialize()
    Dim lstrFile As String '本地工作站配置文件文件名（带路径和后缀）。
    Dim lstr As String
    
    Set mcol单位名称 = New Collection
    
    On Error Resume Next
    '判断本地工作站配置文件文件名是否存在，若不存在，创建它。。
    If Dir("c:\temp") = "" Then
        MkDir "c:\temp"
    End If
    lstrFile = "c:\temp\本地工作站体检配置文件.ini"
    
    If Dir(lstrFile) = "" Then
        '创建工作记忆簿文件。
    End If
    
    '从本地工作站体检配置文件中获取当前工作站的信息。
    '工作模式、Excel文件、内部导入文件、内部导出文件。
    On Error GoTo err_handler
    
    Exit Sub
    
err_handler:
End Sub

Public Property Get Excel文件() As String

End Property

Public Property Let Excel文件(ByVal vNewValue As String)

End Property

Public Property Get 内部导入文件() As String

End Property

Public Property Let 内部导入文件(ByVal vNewValue As String)

End Property

Public Property Get 内部导出文件() As String

End Property

Public Property Let 内部导出文件(ByVal vNewValue As String)

End Property

