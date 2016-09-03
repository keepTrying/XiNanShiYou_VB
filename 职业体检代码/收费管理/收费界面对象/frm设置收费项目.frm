VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F48D4DEC-1198-11D5-91BE-0050BA06B70C}#5.9#0"; "录入控件.ocx"
Begin VB.Form frm设置收费项目 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "设置收费项目"
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   9375
   ClipControls    =   0   'False
   Icon            =   "frm设置收费项目.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   9375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   6165
      Left            =   0
      TabIndex        =   8
      Top             =   840
      Width           =   9255
      Begin VB.TextBox ctxtInput 
         Height          =   375
         Left            =   1920
         TabIndex        =   14
         Top             =   5640
         Width           =   2655
      End
      Begin VB.OptionButton coptFind 
         Caption         =   "按助记符查找"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   5640
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   5805
         Left            =   4800
         TabIndex        =   9
         Top             =   120
         Width           =   4365
         Begin VB.ComboBox ccmb收费项目票据类型 
            Height          =   300
            ItemData        =   "frm设置收费项目.frx":0442
            Left            =   1380
            List            =   "frm设置收费项目.frx":0444
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   4260
            Width           =   2835
         End
         Begin 录入控件.ctlInputBox cinb收费项目编号 
            Height          =   360
            Left            =   240
            TabIndex        =   7
            Top             =   255
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   635
            BackColor       =   16777215
            ForeColor       =   0
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            LeftOfTextbox   =   1120
            Text            =   ""
            Label           =   "收费项目编号"
            Enabled         =   0   'False
            名称            =   ""
            长度            =   0
            允许等于最大值  =   0   'False
            允许等于最小值  =   0   'False
            允许多选        =   0   'False
         End
         Begin 录入控件.ctlInputBox cinb收费项目名称 
            Height          =   360
            Left            =   255
            TabIndex        =   0
            Top             =   827
            Width           =   3960
            _ExtentX        =   6985
            _ExtentY        =   635
            BackColor       =   16777215
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            LeftOfTextbox   =   1120
            Text            =   ""
            Label           =   "收费项目名称"
            名称            =   ""
            长度            =   20
            允许等于最大值  =   0   'False
            允许等于最小值  =   0   'False
            允许多选        =   0   'False
         End
         Begin 录入控件.ctlInputBox cinb收费项目助记符 
            Height          =   360
            Left            =   795
            TabIndex        =   1
            Top             =   1399
            Width           =   3420
            _ExtentX        =   6033
            _ExtentY        =   635
            BackColor       =   16777215
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            LeftOfTextbox   =   580
            Text            =   ""
            Label           =   "助记符"
            名称            =   ""
            长度            =   20
            允许等于最大值  =   0   'False
            允许等于最小值  =   0   'False
            允许多选        =   0   'False
         End
         Begin 录入控件.ctlInputBox cinb收费项目单价 
            Height          =   360
            Left            =   990
            TabIndex        =   2
            Top             =   1965
            Width           =   3240
            _ExtentX        =   5715
            _ExtentY        =   635
            BackColor       =   16777215
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            LeftOfTextbox   =   400
            Text            =   ""
            Label           =   "单价"
            名称            =   ""
            长度            =   5
            允许等于最大值  =   0   'False
            允许等于最小值  =   0   'False
            允许多选        =   0   'False
         End
         Begin 录入控件.ctlInputBox cinb收费项目最小单价 
            Height          =   360
            Left            =   615
            TabIndex        =   3
            Top             =   2550
            Width           =   3600
            _ExtentX        =   6350
            _ExtentY        =   635
            BackColor       =   16777215
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            LeftOfTextbox   =   760
            Text            =   ""
            Label           =   "最小单价"
            名称            =   ""
            长度            =   5
            允许等于最大值  =   0   'False
            允许等于最小值  =   0   'False
            允许多选        =   0   'False
         End
         Begin 录入控件.ctlInputBox cinb收费项目最大单价 
            Height          =   360
            Left            =   615
            TabIndex        =   4
            Top             =   3115
            Width           =   3600
            _ExtentX        =   6350
            _ExtentY        =   635
            BackColor       =   16777215
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            LeftOfTextbox   =   760
            Text            =   ""
            Label           =   "最大单价"
            名称            =   ""
            长度            =   5
            允许等于最大值  =   0   'False
            允许等于最小值  =   0   'False
            允许多选        =   0   'False
         End
         Begin 录入控件.ctlInputBox cinb收费项目计量单位 
            Height          =   360
            Left            =   615
            TabIndex        =   5
            Top             =   3690
            Width           =   3600
            _ExtentX        =   6350
            _ExtentY        =   635
            BackColor       =   16777215
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            LeftOfTextbox   =   760
            Text            =   ""
            Label           =   "计量单位"
            名称            =   ""
            长度            =   5
            允许等于最大值  =   0   'False
            允许等于最小值  =   0   'False
            允许多选        =   0   'False
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "所属票据类型"
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   4320
            Width           =   1095
         End
      End
      Begin MSComctlLib.TreeView ctvwMain 
         Height          =   5295
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   4575
         _ExtentX        =   8070
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
   Begin MSComctlLib.ImageList cimg按钮图标 
      Left            =   5760
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar ctlb设置 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "cimg按钮图标"
      _Version        =   393216
   End
End
Attribute VB_Name = "frm设置收费项目"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public pblnInUse As Boolean

Private WithEvents mobjGUI As cls界面通用对象
Attribute mobjGUI.VB_VarHelpID = -1

Public pint科目级数 As Long

'自动生成助记符.
Private Sub cinb收费项目名称_Change()
    Dim lstrTemp As String
    Dim lobj助记符 As Object
    
    On Error Resume Next
    Set lobj助记符 = CreateObject("助记符.cls助记符")
    lstrTemp = lobj助记符.guf_GetFirstLetter(cinb收费项目名称.Text)
    lstrTemp = Left(lstrTemp, 20)
    cinb收费项目助记符.Text = lstrTemp
End Sub

Private Sub ctvwMain_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim lintcount As Integer
    Dim lrsd收费项目 As Object
    Dim lrsd票据类型 As Object
    On Error GoTo errhandler
    
    If Node.Key <> "s" Then
        Set lrsd收费项目 = dafuncGetData("select * from 收费管理_收费项目字典表 where 收费项目编号='" & Right(LTrim(RTrim(Node.Key)), Len(LTrim(RTrim(Node.Key))) - 1) & "'")
        If Not lrsd收费项目.EOF Then
            cinb收费项目编号.Text = lrsd收费项目.Fields("收费项目编号").Value
            cinb收费项目名称.Text = lrsd收费项目.Fields("收费项目名称").Value
            cinb收费项目单价.Text = lrsd收费项目.Fields("单价").Value
            cinb收费项目计量单位.Text = IIf(IsNull(lrsd收费项目.Fields("计量单位").Value), "", lrsd收费项目.Fields("计量单位").Value)
            cinb收费项目助记符.Text = IIf(IsNull(lrsd收费项目.Fields("助记符").Value), "", lrsd收费项目.Fields("助记符").Value)
            cinb收费项目最小单价.Text = lrsd收费项目.Fields("最小单价").Value
            cinb收费项目最大单价.Text = lrsd收费项目.Fields("最大单价").Value
            
            Set lrsd票据类型 = dafuncGetData("select * from 收费管理_票据类型字典视图")
            If (lrsd票据类型.RecordCount > 0) Then
                lrsd票据类型.MoveFirst
                Do While (Not lrsd票据类型.EOF)
                    If lrsd票据类型("InnerID").Value = Val(lrsd收费项目("票据类型编号").Value) Then
                        Exit Do
                    Else
                        lrsd票据类型.MoveNext
                    End If
                Loop
                If lrsd票据类型.EOF Then
                    If (Len(LTrim(RTrim(ctvwMain.SelectedItem.Key))) - 1) / 3 = pint科目级数 Then
                        MsgBox "如果修改了票据类型字典表，请重新录入该项目的票据类型，该项目票据类型有问题!", vbExclamation, "票据类型设置"
                    End If
                    Exit Sub
                End If
                ccmb收费项目票据类型.Text = lrsd票据类型("名称").Value
            End If

        End If
        cinb收费项目编号.SetFocus
    End If
    
    Exit Sub
errhandler:
    sfsub错误处理 "收费界面部件", "frm设置收费项目", "ctvwMain_MouseDown", Err.Number, Err.Description, False
End Sub

Private Sub ctxtInput_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo errhandler
    Dim i As Long
    If KeyCode = 13 Then
        If Not ctvwMain.SelectedItem Is Nothing Then
            ctvwMain_NodeClick ctvwMain.SelectedItem
        End If
    Else
        '定位。
        Dim lnodeParent As Node
        Dim lNode As Node
        Dim lstrTemp As String
        
        If ctvwMain.SelectedItem.Children = 0 Then
            Set lnodeParent = ctvwMain.SelectedItem.Parent
        Else
            Set lnodeParent = ctvwMain.SelectedItem
        End If
        lnodeParent.Selected = True
        If ctxtInput.Text <> "" Then
            If lnodeParent.Children > 0 Then
                Set lNode = lnodeParent.Child
                For i = 1 To lnodeParent.Children
                    lstrTemp = Right(lNode.Text, Len(lNode.Text) - InStr(lNode.Text, " "))
                    If UCase(Left(lstrTemp, Len(ctxtInput.Text))) = UCase(ctxtInput.Text) Then
                        lNode.Selected = True
                        Exit For
                    Else
                        Set lNode = lNode.Next
                    End If
                Next
            End If
        End If
    End If
    Exit Sub
errhandler:
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 39 Then
        KeyAscii = 0
    ElseIf KeyAscii = 13 Then
        If ActiveControl = ctxtInput Then
        Else
            SendKeys Chr(vbKeyTab)
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim lobjRec As Object
    
    On Error GoTo errhandler
    
    If pblnInUse = True Then Exit Sub
    
    pblnInUse = True

    '初始化工具栏
    Dim lcol工具栏按钮 As Collection
    Set lcol工具栏按钮 = New Collection
    Set mobjGUI = New cls界面通用对象
    Set mobjGUI.Form = Me
    Set mobjGUI.c工具栏 = ctlb设置

    lcol工具栏按钮.Add "添加"
    lcol工具栏按钮.Add "删除"
    lcol工具栏按钮.Add "|"
    lcol工具栏按钮.Add "保存"
    lcol工具栏按钮.Add "|"
    lcol工具栏按钮.Add "退出"
    
    mobjGUI.subInitialize lcol工具栏按钮, ""
    
    
    '初始化票据类型
    Set lobjRec = dafuncGetData("select * from 收费管理_票据类型字典视图")
    If (Not lobjRec.EOF) And (Not lobjRec.BOF) Then
        Do While (Not lobjRec.EOF)
            ccmb收费项目票据类型.AddItem lobjRec.Fields("名称").Value
            ccmb收费项目票据类型.ItemData(ccmb收费项目票据类型.NewIndex) = lobjRec.Fields("innerId").Value
            lobjRec.MoveNext
        Loop
        lobjRec.MoveFirst
    End If
    If ccmb收费项目票据类型.ListCount > 0 Then
        ccmb收费项目票据类型.ListIndex = 0
    End If
    
    '初始化收费项目树。
    Dim lnodParent As Node
    Dim lint级数 As Long
    Dim lstrKey As String
    
    pint科目级数 = Val(pobj收费管理.业务设置("科目级数"))
    If pint科目级数 = 0 Then pint科目级数 = 2
    
    Set lnodParent = ctvwMain.Nodes.Add(, , "s", "收费项目")
    For lint级数 = 1 To pint科目级数
        Set lobjRec = dafuncGetData("select * from 收费管理_收费项目字典表 where Len(收费项目编号) =" & lint级数 * 3 & " order by 助记符")
        Do While (Not lobjRec.EOF)
            lstrKey = "s" & lobjRec("收费项目编号").Value
            ctvwMain.Nodes.Add "s" & Mid(lstrKey, 2, ((lint级数 - 1) * 3)), tvwChild, lstrKey, lobjRec("收费项目名称").Value & " " & IIf(IsNull(lobjRec("助记符")), "", lobjRec("助记符"))
            lobjRec.MoveNext
        Loop
    Next
    ctvwMain.Nodes(1).Expanded = True
    Exit Sub
errhandler:
    sfsub错误处理 "收费界面部件", "frm设置收费项目", "Form_Load", Err.Number, Err.Description, False
    Exit Sub
    Resume

End Sub

Private Sub Form_Unload(Cancel As Integer)
    pblnInUse = False
End Sub

Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    On Error GoTo errhandler
    Select Case Operate
    Case "添加"
        Cancel = True
        cinb收费项目编号.Text = ""
        cinb收费项目名称.Text = ""
        cinb收费项目单价.Text = ""
        cinb收费项目计量单位.Text = ""
        cinb收费项目助记符.Text = ""
        cinb收费项目最大单价.Text = ""
        cinb收费项目最小单价.Text = ""
        cinb收费项目名称.SetFocus
        
    Case "保存"
        Dim lstrParent As String
        Cancel = True
        If ctvwMain.SelectedItem Is Nothing Then
            lstrParent = ""
        ElseIf ctvwMain.SelectedItem.Key = "s" Then
            lstrParent = ""
        Else
            lstrParent = Right(ctvwMain.SelectedItem.Key, Len(ctvwMain.SelectedItem.Key) - 1)
            If Len(lstrParent) = pint科目级数 * 3 Then
                '选中的是末级编码。
                lstrParent = Left(lstrParent, Len(lstrParent) - 3)
            End If
        End If
        
        '校验
        subValidate lstrParent
        
        '向数据库中添加记录
        Dim lobjItem As Object
        Set lobjItem = CreateObject("收费对象部件.cls收费项目")
        lobjItem.收费项目编号 = cinb收费项目编号.Text
        lobjItem.收费项目名称 = cinb收费项目名称.Text
        lobjItem.单价 = cinb收费项目单价.Text
        lobjItem.计量单位 = cinb收费项目计量单位.Text
        lobjItem.票据类型编号 = ccmb收费项目票据类型.ItemData(ccmb收费项目票据类型.ListIndex)
        lobjItem.最小单价 = cinb收费项目最小单价.Text
        lobjItem.最大单价 = cinb收费项目最大单价.Text
        lobjItem.助记符 = cinb收费项目助记符.Text
        lobjItem.sub保存 lstrParent
        
        
        '添加成功则增加节点.
        If cinb收费项目编号.Text = "" Then
            Call ctvwMain.Nodes.Add("s" & lstrParent, tvwChild, "s" & lobjItem.收费项目编号, cinb收费项目名称.Text)
        End If
        '自动新增。
        mobjGUI_BeforeOperate "添加", True
    
    Case "删除"
        Dim lstrKey As String
        Cancel = True
        If ctvwMain.SelectedItem Is Nothing Then
            Err.Raise 6666, , "请选择要删除的收费项目！"
        ElseIf ctvwMain.SelectedItem.Key = "s" Then
            Err.Raise 6666, , "请选择要删除的收费项目！"
        ElseIf ctvwMain.SelectedItem.Children > 0 Then
            Err.Raise 6666, , "请从最下级开始删除！"
        Else
            If MsgBox("你确认要删除收费项目“" & ctvwMain.SelectedItem.Text & "”吗？", vbYesNo + vbQuestion, "系统询问") = vbYes Then
                lstrKey = Right(ctvwMain.SelectedItem.Key, Len(ctvwMain.SelectedItem.Key) - 1)
                pobj收费管理.sub删除收费项目 (lstrKey)
                
                ctvwMain.Nodes.Remove (ctvwMain.SelectedItem.Key)
            End If
        End If
    
    End Select
    Exit Sub
errhandler:
    sfsub错误处理 "收费界面部件", "frm设置收费项目", "mobjGUI_BeforeOperate", Err.Number, Err.Description, False
    Exit Sub
    Resume

End Sub

Private Sub subValidate(ByVal paraParent As String)
    
    If Trim(cinb收费项目名称.Text) = "" Then Err.Raise 6666, , "收费项目名称不能为空，请录入！"
    If ccmb收费项目票据类型.ListIndex = -1 Then Err.Raise 6666, , "必须选择票据类型！"
    If Len(paraParent) = (pint科目级数 - 1) * 3 Then
        '末级科目，必须输入单价。
        If cinb收费项目单价.Text = "" Then Err.Raise 6666, , "最后一级单价不能为空，请修改！"
        If cinb收费项目单价.Text = 0 Then Err.Raise 6666, , "最后一级单价不能为零！"
        
        If IsNumeric(cinb收费项目单价.Text) And IsNumeric(cinb收费项目最大单价.Text) And IsNumeric(cinb收费项目最小单价.Text) Then
               If CDbl(cinb收费项目单价.Text) < 0 Or CDbl(cinb收费项目最大单价.Text) < 0 Or CDbl(cinb收费项目最小单价.Text) < 0 Then
                   Err.Raise 6666, , "单价、最小单价、最大单价必须大于零，请修改！"
               End If
        Else
            Err.Raise 6666, , "单价、最小单价、最大单价必须为数值，请修改！"
        End If
        If CDbl(cinb收费项目单价.Text) > CDbl(cinb收费项目最大单价.Text) Or CDbl(cinb收费项目单价.Text) < CDbl(cinb收费项目最小单价.Text) Then
            Err.Raise 6666, , "单价必须在最小单价和最大单价之间，请修改！"
        End If
             
        If cinb收费项目最大单价.Text = "" And cinb收费项目最小单价.Text = "" Then
            cinb收费项目最大单价.Text = cinb收费项目单价.Text
            cinb收费项目最小单价.Text = cinb收费项目单价.Text
        ElseIf cinb收费项目最大单价.Text = "" And cinb收费项目最小单价.Text <> "" Then
            If CDbl(cinb收费项目最小单价.Text) < CDbl(cinb收费项目单价.Text) Then
                cinb收费项目最大单价.Text = cinb收费项目单价.Text
            End If
        ElseIf cinb收费项目最大单价.Text <> "" And cinb收费项目最小单价.Text = "" Then
            If CDbl(cinb收费项目最大单价.Text) > CDbl(cinb收费项目单价.Text) Then
                cinb收费项目最小单价.Text = cinb收费项目单价.Text
            End If
        End If
             
        If cinb收费项目计量单位.Text = "" Or IsNull(cinb收费项目计量单位.Text) Then Err.Raise 6666, , "请输入计量单位！"
             
    Else
        If Len(LTrim(RTrim(cinb收费项目编号.Text))) / 3 > pint科目级数 Then
            Err.Raise 6666, , "收费项目级数超过限制，该项目设置无效"
        Else
            cinb收费项目单价.Text = "0"
            cinb收费项目最大单价.Text = "0"
            cinb收费项目最小单价.Text = "0"
        End If
    End If

End Sub
