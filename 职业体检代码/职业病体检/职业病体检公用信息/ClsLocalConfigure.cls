VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsLocalConfigure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'设计者：杨春

Private mlng工作模式 As PENUM_工作模式
Private mstrExcel文件 As String
Private mstr内部导入文件 As String
Private mstr内部导出文件 As String

Private mstr配置文件名 As String  '配置文件名及其路径、后缀。

'本地工作站配置文件格式:
'        [工作模式]
'            工作模式 = 0
'        [数据导入导出]
'            Excel文件 = "c:\temp\a.xls"
'            内部导入文件 = "c:\temp\b.dat"
'            内部导出文件 = "c:\temp\c.dat"
'作者：杨春
Private Sub Class_Initialize()
    Dim lstr As String
    
    On Error Resume Next
    '判断本地工作站配置文件文件名是否存在，若不存在，创建它。。
    If Dir("c:\temp", vbDirectory) = "" Then
        MkDir "c:\temp"
    End If
    mstr配置文件名 = "c:\temp\本地工作站体检配置文件.ini"
    
    '读配置文件。
    subRefresh
    
End Sub
Public Property Get 工作模式() As Long
    工作模式 = mlng工作模式
End Property

Public Property Let 工作模式(ByVal vNewValue As Long)
    On Error GoTo errHandler
    mlng工作模式 = vNewValue
    
    '写入本地配置文件中。
    sub写本地配置文件
    Exit Property
errHandler:
    sfsub错误处理 "体检公用信息部件", "ClsLocalConfigure", "sub写本地配置文件", Err.Number, Err.Description, True
End Property

Public Property Get Excel文件() As String
    Excel文件 = mstrExcel文件
End Property

Public Property Let Excel文件(ByVal vNewValue As String)
    On Error GoTo errHandler
    mstrExcel文件 = vNewValue
    
    '写入本地配置文件中。
    sub写本地配置文件
    Exit Property
    
errHandler:
    sfsub错误处理 "体检公用信息部件", "ClsLocalConfigure", "sub写本地配置文件", Err.Number, Err.Description, True
End Property

Public Property Get 内部导入文件() As String
    内部导入文件 = mstr内部导入文件
End Property

Public Property Let 内部导入文件(ByVal vNewValue As String)
    On Error GoTo errHandler
    mstr内部导入文件 = vNewValue

    '写入本地配置文件中。
    sub写本地配置文件
    Exit Property
errHandler:
    sfsub错误处理 "体检公用信息部件", "ClsLocalConfigure", "sub写本地配置文件", Err.Number, Err.Description, True
End Property

Public Property Get 内部导出文件() As String
    内部导出文件 = mstr内部导出文件
End Property

Public Property Let 内部导出文件(ByVal vNewValue As String)
    On Error GoTo errHandler
    mstr内部导出文件 = vNewValue
    
    '写入本地配置文件中。
    sub写本地配置文件
    Exit Property
errHandler:
    sfsub错误处理 "体检公用信息部件", "ClsLocalConfigure", "sub写本地配置文件", Err.Number, Err.Description, True
End Property

'功能：重新读配置文件。
'作者：杨春
Public Sub subRefresh()
    Dim lstr As String
    
    On Error GoTo errHandler
    If Dir(mstr配置文件名) <> "" Then
        '从本地工作站体检配置文件中获取当前工作站的信息。
        '工作模式、Excel文件、内部导入文件、内部导出文件。
        lstr = Space(2)
        Call GetPrivateProfileString("工作模式", "工作模式", "", lstr, 10, mstr配置文件名)
        mlng工作模式 = Val(lstr)
        
        lstr = Space(100)
        Call GetPrivateProfileString("数据导入导出", "Excel文件", "", lstr, 100, mstr配置文件名)
        mstrExcel文件 = Trim(lstr)
        
        lstr = Space(100)
        Call GetPrivateProfileString("数据导入导出", "内部导入文件", "", lstr, 100, mstr配置文件名)
        mstr内部导入文件 = Trim(lstr)
        
        lstr = Space(100)
        Call GetPrivateProfileString("数据导入导出", "内部导出文件", "", lstr, 100, mstr配置文件名)
        mstr内部导出文件 = Trim(lstr)
    Else
        Err.Raise 6666, , "配置文件不存在！"
    End If
    
    Exit Sub
errHandler:
    
End Sub

'功能：把本对象属性写入本地配置文件中。
'作者：杨春
Private Sub sub写本地配置文件()
    On Error GoTo errHandler

    '若配置文件不存在，创建。
    WritePrivateProfileString "工作模式", "工作模式", Str$(mlng工作模式), mstr配置文件名
    WritePrivateProfileString "数据导入导出", "Excel文件", mstrExcel文件, mstr配置文件名
    WritePrivateProfileString "数据导入导出", "内部导入文件", mstr内部导入文件, mstr配置文件名
    WritePrivateProfileString "数据导入导出", "内部导出文件", mstr内部导出文件, mstr配置文件名
    
    Exit Sub
errHandler:
    sfsub错误处理 "体检公用信息部件", "ClsLocalConfigure", "sub写本地配置文件", Err.Number, Err.Description, True
End Sub

