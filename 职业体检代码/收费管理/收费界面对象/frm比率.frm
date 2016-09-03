VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm比率 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "比率设置"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   5925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmd保存 
      Caption         =   "保存"
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox txt比率 
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   3120
      Width           =   1215
   End
   Begin MSComctlLib.TreeView ctrwItem 
      Height          =   2895
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   5106
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Label lbl百分比 
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   3
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label lbl比率 
      Caption         =   "设置比率为："
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   3240
      Width           =   1095
   End
End
Attribute VB_Name = "frm比率"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public pstr收费项目编号 As String

'功能： 将设置的收费科室分配比率信息保存到数据库中
'修改：宋宇 2006/01/23
Private Sub cmd保存_Click()
 On Error GoTo errHandler
 If Left(ctrwItem.SelectedItem.Key, 1) = "k" Then
    Dim mstr科室名称 As Recordset
    Set mstr科室名称 = dafuncGetData("select * from 系统管理_科室字典表 where 名称= '" & Left(Trim(ctrwItem.SelectedItem.Text), (InStr(Trim(ctrwItem.SelectedItem.Text), "(") - 1)) & "'")
    sub修改比率 pstr收费项目编号, mstr科室名称("编号")

 End If
 ctrwItem.Nodes.Clear
 Form_Load
 Exit Sub
errHandler:
    sfsub错误处理 "收费界面", "frm比率", "cmd保存_Click", Err.Number, Err.Description, True
End Sub

Private Sub ctrwItem_NodeClick(ByVal Node As MSComctlLib.Node)
  On Error GoTo errHandler
  If Left(ctrwItem.SelectedItem.Key, 1) = "k" Then
    txt比率.SetFocus
    cmd保存.Enabled = True
    Dim mstr科室编号 As Recordset
'    MsgBox Left(Trim(Node.Text), (InStr(Trim(Node.Text), "(") - 1))
    Set mstr科室编号 = dafuncGetData("select * from 系统管理_科室字典表 where 名称= '" & Left(Trim(Node.Text), (InStr(Trim(Node.Text), "(") - 1)) & "'")
    Dim lobjRec As Recordset
    Set lobjRec = func查询科室项目1(pstr收费项目编号, mstr科室编号("编号"))
    txt比率.Text = IIf(IsNull(lobjRec("比率")), "", lobjRec("比率")) 'lobjRec("比率")'left(Trim(Node.Text),(instr(Trim(Node.Text)-1),"("))
  Else
    txt比率.Text = ""
    cmd保存.Enabled = False
  End If
  Exit Sub
errHandler:
    sfsub错误处理 "收费界面", "frm比率", "ctrwItem_NodeClick", Err.Number, Err.Description, True
End Sub

Private Sub Form_Load()
  On Error GoTo errHandler
  Dim lobjRec As Recordset
  Dim mstr收费项目名称 As Recordset
  Dim mstr科室名称 As Recordset
  Dim lobjNode As Node
  cmd保存.Enabled = False
  Set mstr收费项目名称 = dafuncGetData("select * from 收费管理_收费项目字典表 where 收费项目编号= '" & pstr收费项目编号 & "'")
  Set lobjNode = ctrwItem.Nodes.Add(, , "s" & pstr收费项目编号, mstr收费项目名称("收费项目名称"))
  Set lobjRec = func查询科室项目("'" & pstr收费项目编号 & "'")
  Do While Not lobjRec.EOF
     Set mstr科室名称 = dafuncGetData("select * from 系统管理_科室字典表 where 编号= '" & lobjRec("科室编号") & "'")
     Set lobjNode = ctrwItem.Nodes.Add("s" & pstr收费项目编号, tvwChild, "k" & mstr科室名称("名称") & "s" & pstr收费项目编号, mstr科室名称("名称") & "(" & lobjRec("比率") & "%" & ")")
     lobjNode.EnsureVisible

     lobjRec.MoveNext
  Loop
Exit Sub
errHandler:
    sfsub错误处理 "收费界面", "frm比率", "Form_Load", Err.Number, Err.Description, True
End Sub

Public Sub sub修改比率(para收费项目编号, para科室编号 As String)
    On Error GoTo errHandler
    Dim dec As Variant
    Dim dec1 As Variant
    Dim i As Integer
    Dim lobjRec As Object '结果集
    Set lobjRec = dafuncGetData("select 比率 from 收费管理_科室比率表 where 收费项目编号='" + para收费项目编号 + "'")
    dec = 0
    If lobjRec.RecordCount > 0 Then
       Do While Not lobjRec.EOF

          If IsNull(lobjRec(0)) Then
            dec1 = 0
          Else
            dec1 = lobjRec(0)
          End If
          dec = dec + dec1
          lobjRec.MoveNext
       Loop
    End If

    Set lobjRec = dafuncGetData("select 比率 from 收费管理_科室比率表 where 收费项目编号='" + para收费项目编号 + "'and 科室编号='" + para科室编号 + "'")
    If lobjRec.RecordCount > 0 Then
       If IsNull(lobjRec(0)) Then
          dec1 = 0
       Else
          dec1 = lobjRec(0)
       End If
          dec = dec - dec1
          dec = dec + Val(Trim(txt比率.Text))

    Else
       dec = dec + dec(Trim(txt比率.Text))
    End If

    If dec > 100 Then
       MsgBox "总比率已超过100%，请重新设置。", vbOKOnly, "系统提示"
       txt比率.Text = ""
       txt比率.SetFocus
       Exit Sub
    End If
    
    dafuncGetData ("update 收费管理_科室比率表 set 比率='" & Val(Trim(txt比率.Text)) & "' where 收费项目编号='" + para收费项目编号 + "' and 科室编号='" + para科室编号 + "'")
    'MsgBox "保存成功", vbOKOnly, "系统提示"
    Exit Sub
errHandler:
    sfsub错误处理 "收费界面", "frm比率", "sub修改比率", Err.Number, Err.Description, True
End Sub
'func查询科室项目
'功能:根据查询条件查询符合条件的记录
Public Function func查询科室项目1(ByVal para收费项目编号, para科室编号 As String) As Object
    On Error GoTo errHandler
    If para查询条件 = "ALL" Then
        Set func查询科室项目1 = dafuncGetData("select * from  收费管理_科室比率表")
    Else
        Set func查询科室项目1 = dafuncGetData("select * from 收费管理_科室比率表 where 收费项目编号='" + para收费项目编号 + "' and 科室编号='" + para科室编号 + "'")
    End If
    
    Exit Function
errHandler:
    sfsub错误处理 "收费界面", "frm比率", "func查询科室项目", Err.Number, Err.Description, True
End Function
'func查询科室项目
'功能:根据查询条件查询符合条件的记录
Public Function func查询科室项目(ByVal para查询条件 As String) As Object
    On Error GoTo errHandler
    If para查询条件 = "ALL" Then
        Set func查询科室项目 = dafuncGetData("select * from  收费管理_科室比率表")
    Else
        Set func查询科室项目 = dafuncGetData("select * from 收费管理_科室比率表 where 收费项目编号=" + para查询条件)
    End If
    
    Exit Function
errHandler:
    sfsub错误处理 "收费界面", "frm比率", "func查询科室项目1", Err.Number, Err.Description, True
End Function

Private Sub txt比率_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    KeyAscii = gfuncKeyNum(KeyAscii)
    If KeyAscii = 13 Then
        cmd保存_Click
    End If
End Sub
