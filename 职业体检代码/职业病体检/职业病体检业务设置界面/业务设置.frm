VERSION 5.00
Begin VB.Form frmConfigureMedicalExam 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "体检管理业务设置"
   ClientHeight    =   2850
   ClientLeft      =   1170
   ClientTop       =   870
   ClientWidth     =   7155
   ClipControls    =   0   'False
   Icon            =   "业务设置.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   7155
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton ccmdAction 
      Caption         =   "条码大小设置"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   9
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   960
      Width           =   2500
   End
   Begin VB.CommandButton ccmdAction 
      Caption         =   "重新赋予修改权限"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   8
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "添加、修改、删除体检人员附加项目，以便设置体检表时使用。"
      Top             =   2280
      Width           =   2500
   End
   Begin VB.Frame Frame1 
      Caption         =   "复查"
      Height          =   735
      Left            =   6000
      TabIndex        =   22
      Top             =   4680
      Visible         =   0   'False
      Width           =   2655
      Begin VB.TextBox ctxtdatenumber 
         Height          =   375
         Left            =   1080
         TabIndex        =   23
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "复查周期："
         Height          =   180
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "天"
         Height          =   180
         Left            =   2280
         TabIndex        =   24
         Top             =   360
         Width           =   180
      End
   End
   Begin VB.CommandButton ccmdAction 
      Caption         =   "最终结论模板设置"
      Height          =   450
      Index           =   7
      Left            =   4080
      TabIndex        =   21
      Top             =   2160
      Visible         =   0   'False
      Width           =   2500
   End
   Begin VB.CommandButton ccmdAction 
      Caption         =   "权限设置"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   6
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   360
      Width           =   2500
   End
   Begin VB.CommandButton ccmdAction 
      Caption         =   "体检表设置"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   5
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1560
      Width           =   2500
   End
   Begin VB.CommandButton ccmdAction 
      Caption         =   "体检人员附加项目设置(&O)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   3
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "添加、修改、删除体检人员附加项目，以便设置体检表时使用。"
      Top             =   5520
      Visible         =   0   'False
      Width           =   2500
   End
   Begin VB.CommandButton ccmdAction 
      Caption         =   "体检项目设置(&I)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   2
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "添加、修改、删除体检项目。"
      Top             =   360
      Width           =   2500
   End
   Begin VB.CommandButton ccmdAction 
      Caption         =   "体检结论判断条件设置(&F)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   1
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "设置各体检结论的自动判断条件。"
      Top             =   1560
      Visible         =   0   'False
      Width           =   2500
   End
   Begin VB.CommandButton ccmdAction 
      Caption         =   "体检医师设置(&D)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   0
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "设置各体检医师可以操作的体检项目。"
      Top             =   960
      Width           =   2500
   End
   Begin VB.CommandButton ccmdAction 
      Caption         =   "退 出 (&X)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   4
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1560
      Width           =   2500
   End
   Begin VB.Frame cfraSet 
      Caption         =   "体检登记时是否打印"
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   3
      Left            =   5040
      TabIndex        =   11
      Top             =   3600
      Visible         =   0   'False
      Width           =   4695
      Begin VB.OptionButton coptPrint 
         Caption         =   "打印体检表"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   17
         Top             =   420
         Width           =   1215
      End
      Begin VB.OptionButton coptPrint 
         Caption         =   "不打印"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   3600
         TabIndex        =   7
         Top             =   420
         Width           =   975
      End
      Begin VB.OptionButton coptPrint 
         Caption         =   "打印体检单"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   6
         Top             =   420
         Width           =   1395
      End
   End
   Begin VB.Frame cfraSet 
      Caption         =   "是否快速登记"
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   240
      TabIndex        =   10
      Top             =   4320
      Visible         =   0   'False
      Width           =   4695
      Begin VB.CheckBox chk快速登记计费 
         Caption         =   "快速登记计费"
         Height          =   255
         Left            =   2760
         TabIndex        =   18
         Top             =   240
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.OptionButton coptQuick 
         Caption         =   "是"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   4
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton coptQuick 
         Caption         =   "否"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   5
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame cfraSet 
      Caption         =   "是否照相"
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   5280
      Visible         =   0   'False
      Width           =   4695
      Begin VB.OptionButton coptPhoto 
         Caption         =   "是"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   2
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton coptPhoto 
         Caption         =   "否"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   3
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame cfraSet 
      Caption         =   "是否收费"
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   360
      TabIndex        =   8
      Top             =   2880
      Visible         =   0   'False
      Width           =   4695
      Begin VB.OptionButton coptCharge 
         Caption         =   "否"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
      Begin VB.OptionButton coptCharge 
         Caption         =   "是"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   0
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmConfigureMedicalExam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'功能：职业病体检基本设置

Option Explicit

Private mblnInUse As Boolean '表明当前窗体是否已加载。

Private mblnSys As Boolean

Public Property Get pblnInUse() As Boolean
    pblnInUse = mblnInUse
End Property

Private Sub chk快速登记计费_Click()
    On Error GoTo errHandler
    Dim lstrTemp As String '业务设置值。
    
    If mblnSys Then Exit Sub
    
    If chk快速登记计费.Value = 1 Then
        lstrTemp = "是"
    Else
        lstrTemp = "否"
    End If
    pobj业务对象.Sub修改业务配置 "快速登记是否计费", lstrTemp
    
    Exit Sub
errHandler:
    sfsub错误处理 "职业病设置界面", "frmConfigureMedicalExam", "chk快速登记计费_Click", Err.Number, Err.Description, False
    
End Sub



Private Sub Form_Load()
    On Error GoTo errHandler
    
    '设置窗体已加载标志。
    mblnInUse = True
    
    If um用户编号 = "8882" Then
        ccmdAction(9).Top = ccmdAction(6).Top
        ccmdAction(0).Visible = False
        ccmdAction(6).Visible = False
        ccmdAction(8).Visible = False
    End If
    
    '显示业务设置。
    mblnSys = True
    If pobj业务对象.业务设置("是否收费") = "是" Then
        coptCharge(0).Value = True
    Else
        coptCharge(1).Value = True
    End If
    
    If pobj业务对象.业务设置("是否照像") = "是" Then
        coptPhoto(0).Value = True
    Else
        coptPhoto(1).Value = True
    End If

    
    If pobj业务对象.业务设置("快速登记是否计费") = "是" Then
        chk快速登记计费.Value = 1
    Else
        chk快速登记计费.Value = 0
    End If
    'chk快速登记计费.Enabled = False
    
    If pobj业务对象.业务设置("是否快速登记") = "是" Then
        coptQuick(0).Value = True
        If coptCharge(0).Value Then
            '只有要收费，并且快速登记时该选择项目才可操作
            chk快速登记计费.Enabled = True
        Else
            chk快速登记计费.Value = 0
        End If
    Else
        coptQuick(1).Value = True
        chk快速登记计费.Value = 0
    End If
    
    '增加业务设置“是否使用体检单”。
    If pobj业务对象.业务设置("是否使用体检单") = "否" Then
        coptPrint(0).Visible = False
        coptPrint(2).Left = coptPrint(1).Left
        coptPrint(1).Left = coptPrint(0).Left
    End If
    If pobj业务对象.业务设置("是否打印体检单") = "是" And pobj业务对象.业务设置("是否使用体检单") <> "否" Then
        coptPrint(0).Value = True
    ElseIf pobj业务对象.业务设置("是否打印体检表") = "是" Then
        coptPrint(1).Value = True
    Else
        coptPrint(2).Value = True
    End If
    
    ctxtdatenumber.Text = pobj业务对象.业务设置("复查周期")
    
    mblnSys = False
    
    Exit Sub
errHandler:
    sfsub错误处理 "职业病设置界面", "frmConfigureMedicalExam", "Form_Load", Err.Number, Err.Description, False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '设置窗体未加载标志。
    mblnInUse = False
End Sub

Private Sub ccmdAction_Click(Index As Integer)
    On Error GoTo errHandler
    Dim lobjFrm As Form
    Select Case Index
    Case 0
        Set lobjFrm = frmSetDoctor
    Case 1
        Set lobjFrm = frmSetConclusionFilter
    Case 2
        Set lobjFrm = frmSetTestItem
    Case 3
        Set lobjFrm = frmSetBaseItem
    Case 5
        Set lobjFrm = frmSetMedicalExamTemplate
    Case 6
        Set lobjFrm = frmSetDoctorPermission
    Case 4
        Unload Me
    Case 7
        Set lobjFrm = frmSetConclusion
    Case 8
        Set lobjFrm = frm重新赋予修改权限
    Case 9
        Set lobjFrm = frm条码大小设置
    End Select
    If Not lobjFrm Is Nothing Then
        lobjFrm.Move Me.Left, Me.Top
        lobjFrm.Show 1
    End If
    Set lobjFrm = Nothing
    Exit Sub
errHandler:
    sfsub错误处理 "职业病设置界面", "frmConfigureMedicalExam", "ccmdAction_Click", Err.Number, Err.Description, False
End Sub

Private Sub coptCharge_Click(Index As Integer)
    On Error GoTo errHandler
    Dim lstrTemp As String '业务设置值。
    
    If mblnSys Then Exit Sub
    
    If coptCharge(0).Value Then
        lstrTemp = "是"
     '   If coptQuick(0).Value Then
      '      chk快速登记计费.Enabled = True
       ' End If
    Else
        lstrTemp = "否"
        
     '   chk快速登记计费.Enabled = False
     '   chk快速登记计费.Value = 0
        
    End If
    pobj业务对象.Sub修改业务配置 "是否收费", lstrTemp
    'chk快速登记计费_Click
    Exit Sub
errHandler:
    sfsub错误处理 "职业病设置界面", "frmConfigureMedicalExam", "coptCharge_Click", Err.Number, Err.Description, False
End Sub

Private Sub coptPhoto_Click(Index As Integer)
    Dim lstrTemp As String '业务设置值。
    
    On Error GoTo errHandler
    If mblnSys Then Exit Sub
    
    If coptPhoto(0).Value Then
        lstrTemp = "是"
    Else
        lstrTemp = "否"
    End If
    pobj业务对象.Sub修改业务配置 "是否照像", lstrTemp

    Exit Sub
errHandler:
    sfsub错误处理 "职业病设置界面", "frmConfigureMedicalExam", "coptPhoto_Click", Err.Number, Err.Description, False
End Sub

Private Sub coptPrint_Click(Index As Integer)
    Dim lstr体检单  As String '业务设置值。
    Dim lstr体检表  As String '业务设置值。
    
    On Error GoTo errHandler
    
    If mblnSys Then Exit Sub
    
    lstr体检单 = "否"
    lstr体检表 = "否"
    If coptPrint(0).Value Then
        lstr体检单 = "是"
    ElseIf coptPrint(1).Value Then
        lstr体检表 = "是"
    End If
    pobj业务对象.Sub修改业务配置 "是否打印体检单", lstr体检单
    pobj业务对象.Sub修改业务配置 "是否打印体检表", lstr体检表
    
    Exit Sub
errHandler:
    sfsub错误处理 "职业病设置界面", "frmConfigureMedicalExam", "coptPrint_Click", Err.Number, Err.Description, False
End Sub

Private Sub coptQuick_Click(Index As Integer)
    On Error GoTo errHandler
    Dim lstrTemp As String '业务设置值。
    
    If mblnSys Then Exit Sub
    If coptQuick(0).Value Then
        lstrTemp = "是"
    '    If coptCharge(0).Value Then
    '        chk快速登记计费.Enabled = True
    '    End If
    Else
        lstrTemp = "否"
        
    '    chk快速登记计费.Enabled = False
    '    chk快速登记计费.Value = 0
        
    End If
    pobj业务对象.Sub修改业务配置 "是否快速登记", lstrTemp
    'chk快速登记计费_Click
    Exit Sub
errHandler:
    sfsub错误处理 "职业病设置界面", "frmConfigureMedicalExam", "coptQuick_Click", Err.Number, Err.Description, False
End Sub
Private Sub ctxtDateNumber_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandler
    If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 13 Or KeyAscii = vbKeyBack Then
    Else
        KeyAscii = 0
        Err.Raise 6666, , "复查周期必须输入数字。"
    End If
    
    Exit Sub
errHandler:
    sfsub错误处理 "职业病设置界面", "frmConfigureMedicalExam", "ctxtDateNumber_KeyPress", Err.Number, Err.Description, False
    Exit Sub
    Resume
End Sub


Private Sub ctxtDateNumber_LostFocus()
    On Error GoTo errHandler
   '判断复查周期不能为空
    If Val(ctxtdatenumber.Text) <= 0 Then
        sffuncMsg "请输入复查周期。复查周期必须>0。", sf警告
        ctxtdatenumber.SetFocus
        Exit Sub
    Else
        '判断是否为数字
        If IsNumeric(Trim(ctxtdatenumber.Text)) = False Then
            sffuncMsg "复查周期只能为数字，请重新录入", sf警告
            With ctxtdatenumber
                .SelStart = 0
                .SelLength = Len(Trim(ctxtdatenumber.Text))
                .SetFocus
            End With
            Exit Sub
        Else
            '判断是否大于0
            If ctxtdatenumber.Text < 0 Then
                sffuncMsg "复查周期不能为负，请重新录入", sf警告
                With ctxtdatenumber
                    .SelStart = 0
                    .SelLength = Len(Trim(ctxtdatenumber.Text))
                    .SetFocus
                End With
            End If
        End If
    End If
    If pobj业务对象.业务设置("复查周期") <> Trim(ctxtdatenumber) Then
        pobj业务对象.Sub修改业务配置 "复查周期", Trim(ctxtdatenumber)
    End If
    
    Exit Sub
errHandler:
    sfsub错误处理 "职业病设置界面", "frmConfigureMedicalExam", "ctxtDateNumber_LostFocus", Err.Number, Err.Description, False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errHandler
    '热键处理。
    If Shift And vbAltMask = vbAltMask Then
        Select Case KeyCode
        Case vbKeyD
            ccmdAction_Click 0
        Case vbKeyF
            ccmdAction_Click 1
        Case vbKeyI
            ccmdAction_Click 2
        Case vbKeyO
            ccmdAction_Click 3
        Case vbKeyX
            ccmdAction_Click 4
        End Select
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "职业病设置界面", "frmConfigureMedicalExam", "Form_KeyDown", Err.Number, Err.Description, False
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 39 Then
        '不允许输入“'”。
        KeyAscii = 0
    End If

End Sub
