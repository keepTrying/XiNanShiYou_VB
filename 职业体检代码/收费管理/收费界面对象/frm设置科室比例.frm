VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm设置科室比例 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "收费项目科室比例设置"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   9780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton ccmd右向 
      Caption         =   ">>"
      Height          =   495
      Left            =   4320
      TabIndex        =   2
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton ccmd左向 
      Caption         =   "<<"
      Height          =   495
      Left            =   4320
      TabIndex        =   1
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmd比率 
      Caption         =   "比率设置"
      Height          =   495
      Left            =   4320
      TabIndex        =   0
      Top             =   3360
      Width           =   1335
   End
   Begin MSComctlLib.TreeView ctr项目树 
      Height          =   5175
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   9128
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   6
      Appearance      =   1
   End
   Begin MSComctlLib.TreeView ctr科室树 
      Height          =   5295
      Left            =   6000
      TabIndex        =   4
      Top             =   240
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   9340
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   6
      Appearance      =   1
   End
End
Attribute VB_Name = "frm设置科室比例"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim lnodParent As Node
    Dim lnodTemp As Node
    Dim lint级数 As Integer
    Dim lobj项目 As Object
    Dim lstrtvwme As String
    Dim lobj科室 As Object
    Dim lint科目级数 As Integer
    Dim lstr收费项目编号 As String
    
    On Error GoTo errHandler
    ccmd右向.Enabled = False
    ccmd左向.Enabled = False
    cmd比率.Enabled = False
    Dim lrsd科室项目 As Recordset
'    Dim lstr收费项目编号 As String
    Dim lstr科室编号 As String
    '给项目树赋值
    

    Set lnodParent = ctr项目树.Nodes.Add(, , "s", "收费项目")
    lint科目级数 = Val(pobj收费管理.业务设置("科目级数"))
    If lint科目级数 = 0 Then lint科目级数 = 2

    For lint级数 = 1 To lint科目级数
        Set lobj项目 = dafuncGetData("select 收费项目编号,收费项目名称 from 收费管理_收费项目字典表 where len(收费项目编号)=" & (lint级数 * 3) & "  order by 收费项目编号 ")
        
        If (Not lobj项目.BOF) And (Not lobj项目.EOF) Then
            lobj项目.MoveFirst
            Do While (Not lobj项目.EOF)
                lstrtvwme = "s" & lobj项目("收费项目编号").Value
                Set lnodTemp = ctr项目树.Nodes.Add("s" & Mid(lstrtvwme, 2, ((lint级数 - 1) * 3)), tvwChild, lstrtvwme, lobj项目("收费项目名称").Value)

                lstr收费项目编号 = lobj项目("收费项目编号").Value
                Set lrsd科室项目 = func查询科室项目("'" & lstr收费项目编号 & "'")
                If lrsd科室项目.RecordCount > 0 Then
                   If (Not lrsd科室项目.BOF) And (Not lrsd科室项目.EOF) Then
                      lrsd科室项目.MoveFirst
                      Do While (Not lrsd科室项目.EOF)
                         lstr科室编号 = lrsd科室项目("科室编号").Value
                         Set lobj科室 = dafuncGetData("select * from 系统管理_科室字典表 where 编号= '" & lstr科室编号 & "'")
                         Set lnodTemp = ctr项目树.Nodes.Add(lstrtvwme, tvwChild, "k" & lobj科室("名称") & lstrtvwme, lobj科室("名称"))
                         lrsd科室项目.MoveNext
                         Set lnodTemp = Nothing
                      Loop
                   End If
                End If
                lobj项目.MoveNext
                Set lnodTemp = Nothing
                ctr项目树.Refresh
            Loop
        End If
    Next
    
    '给科室树赋值
    Set lnodParent = ctr科室树.Nodes.Add(, , "s", "科室名称")
    Set lobj科室 = dafuncGetData("select * from 系统管理_科室字典表 order by 编号")
    If (Not lobj科室.BOF) And (Not lobj科室.EOF) Then
        lobj科室.MoveFirst
        Do While (Not lobj科室.EOF)
            lstrtvwme = "s" & lobj科室("编号").Value
            Set lnodTemp = ctr科室树.Nodes.Add("s", tvwChild, lstrtvwme, lobj科室("名称").Value)
            lnodTemp.EnsureVisible '展开所有节点
            lobj科室.MoveNext
            Set lnodTemp = Nothing
            ctr科室树.Refresh
        Loop
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "收费界面部件", "frm设置科室比例", "Form_Load", Err.Number, Err.Description, False
End Sub
Public Function func查询科室项目(ByVal para查询条件 As String) As Object
    On Error GoTo errHandler
    If para查询条件 = "ALL" Then
        Set func查询科室项目 = dafuncGetData("select * from  收费管理_科室比率表")
    Else
        Set func查询科室项目 = dafuncGetData("select * from 收费管理_科室比率表 where 收费项目编号=" + para查询条件)
    End If
    
    Exit Function
errHandler:
    sfsub错误处理 "收费界面", "frm设置", "func查询科室项目", Err.Number, Err.Description, True
End Function

Private Sub ccmd左向_Click()
On Error GoTo errHandler
    Dim mstr科室名称 As Recordset
    If (ctr科室树.SelectedItem.Children = 0) And (Len(ctr项目树.SelectedItem.Key) = 7) Then
        ctr项目树.Nodes.Add ctr项目树.SelectedItem.Key, tvwChild, "k" & ctr科室树.SelectedItem.Text & ctr项目树.SelectedItem.Key, ctr科室树.SelectedItem.Text ' ctr项目树.SelectedItem.Key & ctr科室树.SelectedItem.Key
        Set mstr科室名称 = dafuncGetData("select * from 系统管理_科室字典表 where 名称= '" & Trim(ctr科室树.SelectedItem.Text) & "'")
        sub保存 Right(ctr项目树.SelectedItem.Key, 6), mstr科室名称("编号") 'ctr科室树.SelectedItem.Text
        ctr项目树.Refresh
 
    End If
Exit Sub
errHandler:
    If Err.Number = 35602 Then
         Err.Number = 0
         'Err.Raise 6666
         
         Exit Sub
    End If
    sfsub错误处理 "收费界面部件", "frm设置科室比例", "ccmd左向_Click", Err.Number, Err.Description, False
End Sub

Private Sub ccmd右向_Click()
On Error GoTo errHandler
Dim mstr科室名称 As Recordset

If ctr项目树.SelectedItem.Children = 0 And Left(ctr项目树.SelectedItem.Key, 1) = "k" Then
   If MsgBox("确定要从" & ctr项目树.SelectedItem.Parent.Text & "项目中删除该科室吗？", vbOKCancel, "系统提示") = vbOK Then
      Set mstr科室名称 = dafuncGetData("select * from 系统管理_科室字典表 where 名称= '" & Trim(ctr项目树.SelectedItem.Text) & "'")
      sub删除 Right(ctr项目树.SelectedItem.Key, 6), mstr科室名称("编号") 'ctr项目树.SelectedItem.Text
    
    '删除选定的节点
    ctr项目树.Nodes.Remove ctr项目树.SelectedItem.Index
    ctr项目树.Refresh
   End If
End If

Exit Sub
errHandler:
    sfsub错误处理 "收费界面部件", "frm设置科室比例", "ccmd右向_Click", Err.Number, Err.Description, False
End Sub

Private Sub cmd比率_Click()
    If (ctr项目树.SelectedItem.Children > 0) And (Len(ctr项目树.SelectedItem.Key) = 7) Then
      
      frm比率.pstr收费项目编号 = Right(ctr项目树.SelectedItem.Key, 6)
      frm比率.Show 1
    End If
End Sub

Private Sub ctr科室树_NodeClick(ByVal Node As MSComctlLib.Node)
     If Len(ctr项目树.SelectedItem.Key) = 7 Then
        ccmd左向.Enabled = True
     Else
        ccmd左向.Enabled = False
     End If

End Sub

Private Sub ctr项目树_NodeClick(ByVal Node As MSComctlLib.Node)
    If Left(ctr项目树.SelectedItem.Key, 1) = "k" Then
      ccmd右向.Enabled = True
    Else
      ccmd右向.Enabled = False
    End If
    If Len(ctr项目树.SelectedItem.Key) = 7 And ctr项目树.SelectedItem.Children > 0 Then
      cmd比率.Enabled = True
    Else
      cmd比率.Enabled = False
    End If
End Sub


Public Sub sub保存(para收费项目编号, para科室编号 As String)
    On Error GoTo errHandler
    Dim lobjRec As Object '结果集
    Set lobjRec = dafuncGetData("select * from 收费管理_科室比率表 where 收费项目编号 = '" + para收费项目编号 + "'and 科室编号 ='" + para科室编号 + "' ")
    
    If lobjRec.RecordCount > 0 Then
       Exit Sub
    End If
    
    dasubBeginTran
      dafuncGetData ("insert into 收费管理_科室比率表(收费项目编号,科室编号)  values('" + para收费项目编号 + "','" + para科室编号 + "')")
    dasubCommitTran
    
    Exit Sub
errHandler:
    sfsub错误处理 "收费界面", "frm设置", "sub保存", Err.Number, Err.Description, True
End Sub

'功能：从数据库中清除收费项目信息的科室设置
'修改：徐冀川 2006/06/05
Public Sub sub删除(para收费项目编号, para科室编号 As String)
    On Error GoTo errHandler
    dafuncGetData ("delete from 收费管理_科室比率表  where 收费项目编号='" + para收费项目编号 + "' and 科室编号= '" + para科室编号 + "'")
    Exit Sub
errHandler:
    sfsub错误处理 "收费界面", "frm设置", "sub删除", Err.Number, Err.Description, True
End Sub
