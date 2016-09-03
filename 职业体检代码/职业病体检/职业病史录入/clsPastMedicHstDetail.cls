VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPastMedcHstDetl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'****************************************
'名称：职业病史(受检者个人信息)详细类模块
'函数：
'功能：职业病史(受检者个人信息)录入 职业病史详细参数类的封装
'作者：Yunle Liu
'时间：2012.03
'****************************************


Option Explicit

Public mstr编号 As String
Public mstr疾病名称 As String
Public mstr诊断单位 As String
Public mstr诊断日期 As String
Public mstr转归 As String
Public mstr治疗经过 As String
