VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMedicalExamEnum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'作者：杨春

Public Enum PENUM_工作模式
    P_NET_STYLE = 0    '网络模式。
    P_LOCAL_STYLE      '单机模式。
End Enum

Public Enum PENUM_体检类型
    P_EXAM_FIRST = 0    '初检
    P_EXAM_AGAIN = 1    '复检
    P_EXAM_ANNUAL = 2   '年检
End Enum
Public Enum PENUM_体检状态
    P_LOGIN_STATUS = 0  '登记状态
    P_EXAMING_STATUS    '正在检验状态
    P_CONCLUED_STATUS   '已下结论状态
    P_ENDED_STATUS      '体检结论已确定状态
End Enum
Public Enum PENUM_化验单状态
    P_BIG_NOT_CHECKED_STATUS = 0 '待复核
    P_BIG_WRONG_STATUS           '复核有错误待修改
    P_BIG_OK_STATUS              '复核通过
End Enum

'错误码。
Public Enum PENUM_错误码
    P_ERRORCODE_VIOLATE = -2147217873 '违背外键关联原则。
    P_ERRORCODE_CREATEOBJECT_FAILED = 429 '创建activex部件的对象失败。
End Enum
