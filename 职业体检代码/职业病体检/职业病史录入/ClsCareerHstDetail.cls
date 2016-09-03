VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsCareerHstDetl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************
'名称：职业病史(受检者个人信息)录入详细类模块
'函数：
'功能：职业病史(受检者个人信息)录入 职业史详细参数类的封装
'作者：Yunle Liu
'时间：2012.03
'***************************************


Option Explicit
Public mstr编号 As String
Public mstr单位 As String
Public mstr部门 As String
Public mstr工种 As String
Public mstr备注 As String
Public mstr起始时间 As String
Public mstr结束时间 As String
Public mstr危害种类 As String
Public mstr措施 As String
Public mstr接触时间 As String
Public mstr放射种类 As String
Public mstr工作量 As String
Public mstr照射量 As String
Public mstr过量照射史 As String
Public mstr是否放射性 As String
