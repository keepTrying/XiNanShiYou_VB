VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSymptomDetl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'*****************************************
'名称：自觉症状录入详细类模块
'函数：
'功能：职业病史(受检者个人信息)录入 自觉症状 详细参数类的封装
'作者：Yunle Liu
'时间：2012.03
'*****************************************

Option Explicit

Public mstr编号 As String
Public mstr症状 As String
Public mstr出现时间 As String
Public mstr程度 As String

