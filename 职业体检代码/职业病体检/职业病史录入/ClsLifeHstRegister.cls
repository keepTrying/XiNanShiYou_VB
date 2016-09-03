VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLifeHstRegt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'****************************************
'名称：生活史获取类模块
'函数：
'功能：职业病史(受检者个人信息)录入 生活史 类的封装
'作者：Yunle Liu
'时间：2012.03
'****************************************

Option Explicit

'Private pblnInUse As Boolean
'个人信息栏
Private mstr系统编号 As String
Private mstr姓名 As String
Private mstr性别 As String
Private mstr年龄 As String
Private mstr单位 As String
Private mstr现工种 As String
Private mstr职务 As String
Private mstr单位名称 As String
Private mstr危害因素 As String

'个人生活栏
Public mstr婚否 As String
Public mstrmatehelh As String
Public mstrmarrydate As String
Public mstrmatejob As String
Public mstrmateradioac As String
Public mstr异位妊娠 As String
Public mstr异常胎 As String
Public mstr孕次 As String
Public mstr活产 As String
Public mstr早产 As String
Public mstr死产 As String
Public mstr现有子女 As String
Public mstr流产 As String
Public mstr畸胎 As String
Public mstr多胎 As String
Public mstr子女健康 As String
Public mstr不孕不育 As String
Public mstr饮酒 As String
Public mstr吸烟 As String
Public mstr酒龄 As String
Public mstr烟龄 As String
Public mstr戒烟时长 As String
Public mstr吸烟量 As String
Public mstr饮酒量 As String
Public mstrMore As String
Public mstr过敏史 As String
Public mstr家族史 As String
Public mstr初潮 As String
Public mstr经期 As String
Public mstr周期 As String
Public mstr末次月经 As String
Public mstr停经 As String
Public mstrOther As String
'增加女孩个数、男孩出生日期、女孩出生日期 2015-7-1 by lanchao
Public mstr现有女孩 As String
Public mstr男孩出生日期 As String
Public mstr女孩出生日期 As String
'增加女孩个数1、男孩出生日期1、女孩出生日期1 2015-9-28
Public mstr现有女孩1 As String
Public mstr男孩出生日期1 As String
Public mstr女孩出生日期1 As String
Public tmp已登记 As Boolean
Public mblninuse As Boolean

'功能：表明当前窗体是否已加载，以便主导航界面判断当前窗体是否已执行过Form_Load。
Public Property Get pblnInUse() As Boolean
    On Error GoTo errHandler
    pblnInUse = mblninuse
    Exit Property
errHandler:
    sfsub错误处理 "职业病史录入", "frmlifehstregt", "Property Get pblnInUse", Err.Number, Err.Description, True
End Property

Public Property Get 年龄() As String
    年龄 = mstr年龄
End Property

Public Property Get 性别() As String
    性别 = mstr性别
End Property
Public Property Get 姓名() As String
    姓名 = mstr姓名
End Property
Public Property Get 单位名称() As String
    单位名称 = mstr单位名称
End Property
Public Property Get 职务() As String
    职务 = mstr职务
End Property
Public Property Get 现工种() As String
    现工种 = mstr现工种
End Property
Public Property Get 危害因素() As String
   危害因素 = mstr危害因素
End Property

Public Property Let 系统编号(ByVal vNewValue As String)
    Dim lobjRec As Object
    On Error GoTo errHandler
    mstr系统编号 = vNewValue
     tmp已登记 = True
    If Len(mstr系统编号) > 5 Then
        dasubSetQueryTimeout 6000
        Set lobjRec = dafuncGetData("select * from 职业病体检_体检人员基本信息表 where 系统编号='" & mstr系统编号 & "'")
        If lobjRec.RecordCount > 0 Then
            mstr姓名 = IIf(IsNull(lobjRec("姓名")), "", lobjRec("姓名"))
            mstr性别 = IIf(IsNull(lobjRec("性别")), "", lobjRec("性别"))
            mstr年龄 = IIf(IsNull(lobjRec("年龄")), "", lobjRec("年龄"))
            mstr单位名称 = IIf(IsNull(lobjRec("单位名称")), "", lobjRec("单位名称"))
            mstr现工种 = IIf(IsNull(lobjRec("现工种")), "", lobjRec("现工种"))
            mstr职务 = IIf(IsNull(lobjRec("职务或职称")), "", lobjRec("职务或职称"))
            mstr危害因素 = IIf(IsNull(lobjRec("危害因素")), "", lobjRec("危害因素"))
        Else
            MsgBox "没找到该条码对应的信息，请检查是否已登记！"
            tmp已登记 = False
        End If
        lobjRec.Close
    End If
    Exit Property
errHandler:
    sfsub错误处理 "职业病史录入", "clslifehstregt", "Property Let 系统编号", Err.Number, Err.Description, True
End Property

'保存个人生活史
Public Sub subSaveLifeHst()
    Dim mbln是否已存在 As Boolean
    Dim lstrsql As String
    Dim lobjRec As Object
    On Error GoTo errHandler
    mbln是否已存在 = False
    If Len(Trim(mstr系统编号)) > 5 Then
        Set lobjRec = dafuncGetData("select * from 职业病体检_个人生活史表 where 系统编号='" & mstr系统编号 & "'")
        If lobjRec.RecordCount > 0 Then
            mbln是否已存在 = True
        End If
    Else
        MsgBox "系统编号错误，请检查！", vbInformation, "系统提示"
        Exit Sub
    End If
     '保存体检基本信息。
    If mbln是否已存在 Then
        '已存在，修改记录。
        'lstrsql = "update 职业病体检_个人生活史表 set 试管编号='" & mstr试管编号 & "',体检表编号='" & mobj体检表.体检表名 & "'"
        '        & ",体检日期='" & mstr体检日期 & " ',收费批号='" & mstr收费批号 & "',体检类别='" & mlng体检人类别 & "',体检类型='" & mlng体检人类型 & "'" _
         '       & " where 系统编号='" & mstr系统编号 & "'"
    Else
        '新增，插入记录。
        lstrsql = "insert into 职业病体检_个人生活史表(系统编号,初潮,周期,经期,末次月经,停经年龄,是否结婚,结婚日期,配偶接触放射,配偶职业,配偶健康状况,孕次,活产,早产,死产,自然流产,畸胎,多胎,异位妊娠,异常胎,不孕不育原因,现有子女数目,子女健康状况,过敏史,吸烟程度,饮酒程度,烟龄,酒龄,戒烟时长,吸烟量,饮酒量,生活更多,家族史,现有女孩,男孩出生日期,女孩出生日期,其他) values(" _
                & "'" & mstr系统编号 & "','" & mstr初潮 & "','" & mstr周期 & "','" & mstr经期 & "','" & mstr末次月经 & "','" & mstr停经 & "','" & mstr婚否 & "','" & mstrmarrydate & "','" & mstrmateradioac & "','" & mstrmatejob & "','" & mstrmatehelh & "','" & mstr孕次 & "','" _
                & mstr活产 & "','" & mstr早产 & "','" & mstr死产 & "','" & mstr流产 & "','" & mstr畸胎 & "','" & mstr多胎 & "','" & mstr异位妊娠 & "','" & mstr异常胎 & "','" & mstr不孕不育 & "','" & mstr现有子女 & "','" & mstr子女健康 & "','" & mstr过敏史 & "','" & mstr吸烟 & "','" & mstr饮酒 & "','" _
                & mstr烟龄 & "','" & mstr酒龄 & "','" & mstr戒烟时长 & "','" & mstr吸烟量 & "','" & mstr饮酒量 & "','" & mstrMore & "','" & mstr家族史 & "','" & mstr现有女孩 & "','" & mstr男孩出生日期 & "','" & mstr女孩出生日期 & "','" & mstrOther & "'" & ")"
    dafuncGetData lstrsql
    End If
    'dafuncGetData lstrsql
    Exit Sub
errHandler:
    sfsub错误处理 "职业病史录入", "frmlifehstregt", "public sub subsavelifehst", Err.Number, Err.Description, True
End Sub

'删除  个人生活史
Public Sub subDelLifeHst()
    Dim lstrsql As String
    On Error GoTo errHandler
    lstrsql = "delete from 职业病体检_个人生活史表 where 系统编号='" & mstr系统编号 & "'"
    dafuncGetData lstrsql
    Exit Sub
errHandler:
    sfsub错误处理 "职业病史录入", "frmlifehstregt", "public sub subdellifehst", Err.Number, Err.Description, True
End Sub
