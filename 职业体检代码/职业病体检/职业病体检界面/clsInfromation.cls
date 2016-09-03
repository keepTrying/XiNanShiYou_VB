VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInfromation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'查看信息的内容是直接从数据库里取出的，不需要用到下面的内容，否则会出现编译错误  2015-10-27

''个人信息栏
'Private mstr系统编号 As String
'Private mstr姓名 As String
'Private mstr性别 As String
'Private mstr年龄 As String
'Private mstr单位 As String
'Private mstr现工种 As String
'Private mstr职务 As String
'Private mstr单位名称 As String
'Private mstr危害因素 As String
'
''功能：表明当前窗体是否已加载，以便主导航界面判断当前窗体是否已执行过Form_Load。
'Public Property Get pblnInUse() As Boolean
'    On Error GoTo errHandler
'    pblnInUse = mblnInUse
'    Exit Property
'End Property
'Public Property Get 年龄() As String
'    年龄 = mstr年龄
'End Property
'
'Public Property Get 性别() As String
'    性别 = mstr性别
'End Property
'Public Property Get 姓名() As String
'    姓名 = mstr姓名
'End Property
'Public Property Get 单位名称() As String
'    单位名称 = mstr单位名称
'End Property
'Public Property Get 职务() As String
'    职务 = mstr职务
'End Property
'Public Property Get 现工种() As String
'    现工种 = mstr现工种
'End Property
'Public Property Get 危害因素() As String
'   危害因素 = mstr危害因素
'End Property
'Public Property Let 系统编号(ByVal vNewValue As String)
'    Dim lobjRec As Object
'    On Error GoTo errHandler
'    mstr系统编号 = vNewValue
'     tmp已登记 = True
'    If Len(mstr系统编号) > 5 Then
'        dasubSetQueryTimeout 6000
'        Set lobjRec = dafuncGetData("select * from 职业病体检_体检人员基本信息表 where 系统编号='" & mstr系统编号 & "'")
'        If lobjRec.RecordCount > 0 Then
'            mstr姓名 = IIf(IsNull(lobjRec("姓名")), "", lobjRec("姓名"))
'            mstr性别 = IIf(IsNull(lobjRec("性别")), "", lobjRec("性别"))
'            mstr年龄 = IIf(IsNull(lobjRec("年龄")), "", lobjRec("年龄"))
'            mstr单位名称 = IIf(IsNull(lobjRec("单位名称")), "", lobjRec("单位名称"))
'            mstr现工种 = IIf(IsNull(lobjRec("现工种")), "", lobjRec("现工种"))
'            mstr职务 = IIf(IsNull(lobjRec("职务或职称")), "", lobjRec("职务或职称"))
'            mstr危害因素 = IIf(IsNull(lobjRec("危害因素")), "", lobjRec("危害因素"))
'        End If
'        lobjRec.Close
'    End If
'    Exit Property
'errHandler:
'    sfsub错误处理 "职业病史录入", "clslifehstregt", "Property Let 系统编号", Err.Number, Err.Description, True
'End Property

