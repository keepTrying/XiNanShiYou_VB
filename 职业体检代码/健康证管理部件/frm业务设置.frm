VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frm业务设置 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "健康证业务设置"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   8760
   Icon            =   "frm业务设置.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   8760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox ctxtFlowNo 
      Height          =   375
      Left            =   2160
      TabIndex        =   12
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox ctxt地区编码 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   2160
      TabIndex        =   9
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton ccmdPrintSetting 
      Caption         =   "打印格式设置"
      Height          =   495
      Left            =   480
      TabIndex        =   8
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CheckBox cchk照相 
      Caption         =   "登记时要照像"
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Caption         =   "调离设置"
      Height          =   5295
      Left            =   3960
      TabIndex        =   3
      Top             =   960
      Width           =   4455
      Begin VB.ListBox clstDisease 
         Height          =   4050
         ItemData        =   "frm业务设置.frx":0442
         Left            =   240
         List            =   "frm业务设置.frx":0449
         Style           =   1  'Checkbox
         TabIndex        =   4
         Top             =   840
         Width           =   3495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "需要调离的病种："
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "健康证格式设置"
      Height          =   1935
      Left            =   360
      TabIndex        =   1
      Top             =   3960
      Visible         =   0   'False
      Width           =   3135
      Begin VB.CheckBox cchk手工 
         Caption         =   "手工输入健康证号"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   2055
      End
      Begin VB.CheckBox cchk条码 
         Caption         =   "健康证带条码"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Value           =   1  'Checked
         Width           =   1455
      End
   End
   Begin MSComctlLib.Toolbar C工具栏 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "cimg按钮图标"
      _Version        =   393216
      Begin MSComctlLib.ImageList cimg按钮图标 
         Left            =   1200
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   15
         ImageHeight     =   15
         MaskColor       =   12632256
         _Version        =   393216
      End
   End
   Begin VB.Label Label3 
      Caption         =   "设置健康证流水号："
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "设置地区编码："
      Height          =   375
      Left            =   480
      TabIndex        =   10
      Top             =   2520
      Width           =   1455
   End
End
Attribute VB_Name = "frm业务设置"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public pblnInUse As Boolean

Private WithEvents mobjGUI As cls界面通用对象 '界面上引用的界面通用对
Attribute mobjGUI.VB_VarHelpID = -1

Private Sub cchk条码_Click()
    On Error Resume Next
    If cchk条码.Value = 1 Then
        cchk手工.Value = 0
        cchk手工.Enabled = False
    Else
        cchk手工.Enabled = True
    End If
End Sub

Private Sub ccmdPrintSetting_Click()
    frm打印格式设置.Show 1
End Sub

Private Sub Form_Load()
    Dim lcolInfo As Collection
    Dim lstrTemp As String
    Dim i As Long
    Dim lobjRec As Object
    On Error GoTo errhandler
    
    If pblnInUse Then Exit Sub
    pblnInUse = True
    
    '初始化工具栏。
    Dim lcol工具栏按钮 As New Collection '工具栏上的按钮初始化集合。
    Set mobjGUI = New cls界面通用对象
    Set mobjGUI.Form = Me
    Set mobjGUI.C工具栏 = C工具栏
    lcol工具栏按钮.Add "保存"
    lcol工具栏按钮.Add "|"
    lcol工具栏按钮.Add "退出"
    mobjGUI.subInitialize lcol工具栏按钮, ""
    
    
    '调离病种。
    lstrTemp = pobj体检管理.业务设置("调离设置")
    '获取检出病种。
    Set lcolInfo = pobj记忆.记忆项值("检出病种", True)
    clstDisease.Clear
    For i = 1 To lcolInfo.Count
        clstDisease.AddItem lcolInfo(i)
        If InStr("," + lstrTemp + ",", "," & lcolInfo(i) & ",") > 0 Then
            clstDisease.Selected(clstDisease.ListCount - 1) = True
        End If
    Next
        
'    lstrTemp = pobj体检管理.业务设置("健康证带条码")
'    If lstrTemp = "是" Then
'        cchk条码.Value = 1
'    Else
'        cchk条码.Value = 0
'        lstrTemp = pobj体检管理.业务设置("手工输入健康证号")
'        If lstrTemp = "是" Then
'            cchk手工.Value = 1
'        Else
'            cchk手工.Value = 0
'        End If
'
'
'    End If
    
    If pobj体检管理.业务设置("是否照相") = "是" Then
        cchk照相.Value = 1
    Else
        cchk照相.Value = 0
    End If
    Set lobjRec = dafuncGetData("select top 1 回函地址 from 健康证_业务配置表")
    ctxt地区编码.Text = lobjRec(0)
            '填充健康证流水号
    Set lobjRec = dafuncGetData("select 当前值 from 系统管理_系统编号生成记录表 where 业务名称='健康证管理' and 编号名称='健康证编号'")
    If lobjRec.RecordCount > 0 Then
        ctxtFlowNo.Text = lobjRec(0)
    End If

    
    Exit Sub
errhandler:
    sfsub错误处理 "健康证管理部件", "frm业务设置", "Form_Load", Err.Number, Err.Description, False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set mobjGUI = Nothing
    pblnInUse = False
End Sub

Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    Dim i As Long
    Dim lstrTemp As String
    Dim lobjRec As Object
    On Error GoTo errhandler
    Select Case Operate
    Case "保存"
        pobj体检管理.业务设置("健康证带条码") = IIf(cchk条码.Value = 1, "是", "否")
        pobj体检管理.业务设置("手工输入健康证号") = IIf(cchk手工.Value = 1, "是", "否")
        
        lstrTemp = ""
        For i = 0 To clstDisease.ListCount - 1
            If clstDisease.Selected(i) Then
                lstrTemp = lstrTemp & clstDisease.List(i) & ","
            End If
        Next
        If lstrTemp <> "" Then lstrTemp = Left(lstrTemp, Len(lstrTemp) - 1)
        '地区编码
        pobj体检管理.业务设置("调离设置") = lstrTemp
        pobj体检管理.业务设置("是否照相") = IIf(cchk照相.Value = 1, "是", "否")
        dafuncGetData ("update 健康证_业务配置表 set 回函地址='" & ctxt地区编码.Text & "'")
        '健康证编号
        Set lobjRec = dafuncGetData("select * from 系统管理_系统编号生成记录表 where 业务名称='健康证管理' and 编号名称='健康证编号'")
        If lobjRec.RecordCount > 0 Then
            dafuncGetData ("update 系统管理_系统编号生成记录表 set 当前值='" & IIf(Trim(ctxtFlowNo.Text) = "", 0, Trim(ctxtFlowNo.Text)) & "' where 业务名称='健康证管理' and 编号名称='健康证编号'")
        Else
            dafuncGetData ("Insert Into 系统管理_系统编号生成记录表(业务名称,编号名称,当前值,数据类型,长度,最大值,是否按年重编,当前年号) values " _
                            & "('健康证管理','健康证编号','" & IIf(Trim(ctxtFlowNo.Text) = "", 0, Trim(ctxtFlowNo.Text)) & "','C','6','999999','否','" & Year(Date) & "')")
        End If
        Cancel = True
        
    End Select
    
    Exit Sub
errhandler:
    sfsub错误处理 "健康证管理部件", "frm业务设置", "mobjGUI_BeforeOperate", Err.Number, Err.Description, False
    
End Sub
