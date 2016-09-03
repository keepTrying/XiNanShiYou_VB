VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mfrmProgress As frmProgress

Private mstrTitle As String
Private mstrMsg As String

Private mlngMaxValue As Long
Private mlngMinValue As Long
Private mlngValue As Long

Private Sub Class_Initialize()
    Set mfrmProgress = New frmProgress
    mlngMaxValue = 100
    mlngMinValue = 0
    mlngValue = 0
    mfrmProgress.cpbMain.Min = 0
    mfrmProgress.cpbMain.Max = 100
    mfrmProgress.cpbMain.Value = 0
End Sub

Public Property Get FormTitle() As String
    FormTitle = mstrTitle
End Property
Public Property Let FormTitle(ByVal vNewValue As String)
    mstrTitle = vNewValue
    mfrmProgress.Caption = mstrTitle
End Property

Public Property Get Message() As String
    Message = mstrMsg
End Property
Public Property Let Message(ByVal vNewValue As String)
    mstrMsg = vNewValue
    mfrmProgress.clblInfo.Caption = mstrMsg
    DoEvents
End Property

Public Property Get MaxValue() As Long
    MaxValue = mlngMaxValue
End Property
Public Property Let MaxValue(ByVal vNewValue As Long)
    mlngMaxValue = vNewValue
    mfrmProgress.cpbMain.Max = mlngMaxValue
End Property

Public Property Get MinValue() As Long
    MinValue = mlngMinValue
End Property
Public Property Let MinValue(ByVal vNewValue As Long)
    mlngMinValue = vNewValue
    mfrmProgress.cpbMain.Min = mlngMinValue
End Property

Public Property Get Value() As Long
    Value = mlngValue
End Property
Public Property Let Value(ByVal vNewValue As Long)
    mlngValue = vNewValue
    mfrmProgress.cpbMain.Value = mlngValue
    
End Property

'功能：显示进度界面。
Public Sub subShow()
    mfrmProgress.Show
    DoEvents
End Sub

Public Sub subUnload()
    Unload mfrmProgress
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Unload mfrmProgress
End Sub

