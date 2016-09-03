VERSION 5.00
Begin VB.Form FrmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3525
   ClientLeft      =   1725
   ClientTop       =   1965
   ClientWidth     =   5550
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3525
   ScaleWidth      =   5550
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox ctxtUserNo 
      BeginProperty Font 
         Name            =   "System"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Left            =   2415
      MaxLength       =   8
      TabIndex        =   0
      Top             =   1560
      Width           =   2100
   End
   Begin VB.TextBox ctxtPassWord 
      BeginProperty Font 
         Name            =   "System"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      IMEMode         =   3  'DISABLE
      Left            =   2415
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2115
      Width           =   2100
   End
   Begin VB.Image ccmdCancel 
      Height          =   300
      Left            =   3165
      Picture         =   "frmLogin.frx":0000
      Top             =   2745
      Width           =   945
   End
   Begin VB.Image ccmdOk 
      Height          =   300
      Left            =   1290
      Picture         =   "frmLogin.frx":27B9
      Top             =   2745
      Width           =   945
   End
   Begin VB.Label clblSysName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   0
      TabIndex        =   5
      Top             =   885
      Width           =   5535
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "用户登录"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3960
      TabIndex        =   4
      Top             =   210
      Width           =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "口    令："
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   1155
      TabIndex        =   3
      Top             =   2115
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "用户编号："
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   1185
      TabIndex        =   2
      Top             =   1620
      Width           =   900
   End
   Begin VB.Image Image2 
      Height          =   645
      Left            =   -60
      Picture         =   "frmLogin.frx":4FC7
      Top             =   0
      Width           =   5610
   End
   Begin VB.Image Image1 
      Height          =   2985
      Left            =   -60
      Picture         =   "frmLogin.frx":8F65
      Stretch         =   -1  'True
      Top             =   645
      Width           =   5610
   End
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



'须先输入用户编号
Private Sub ctxtpassword_Change()
    If ctxtUserNo.Text = "" Then
        Call sffuncMsg("请先输入用户编号", sf警告)
        ctxtUserNo.SetFocus
    End If
End Sub

'得到焦点，选中所有字符
Private Sub ctxtPassWord_GotFocus()
    ctxtPassWord.SelStart = 0
    ctxtPassWord.SelLength = Len(ctxtPassWord.Text)
End Sub

'得到焦点，选中所有字符
Private Sub ctxtUserNo_GotFocus()
    ctxtUserNo.SelStart = 0
    ctxtUserNo.SelLength = Len(ctxtUserNo.Text)
End Sub

'响应回车键
Private Sub ctxtUSERNO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ctxtPassWord.SetFocus
    End If
    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

'响应回车键
Private Sub ctxtPAssword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        '调用校验用户
        Call ccmdOk_Click
    End If
End Sub

'用户取消，退出系统，释放对象
Private Sub ccmdCancel_Click()
    Set pobj平台结构 = Nothing
    Unload Me
End Sub

Private Sub ccmdOk_Click()
    On Error GoTo errHandle
    'Dim lfrm主界面 As New frm主界面
    Dim lfrm主界面 As New frmMain
    Dim i As Long
    
    '验证身份
    If umfunc校验身份(Trim(ctxtUserNo.Text), ctxtPassWord.Text) Then
        '合法用户
        '修改：2002-1-28（登记配置表，当前用户编号）。
        On Error Resume Next
        sfsubSaveSetting "系统管理", "本地配置", "用户编号", Trim(ctxtUserNo.Text)
        
        On Error GoTo errHandle
        If func是否注册(Trim(ctxtUserNo.Text)) Then
            If Trim(um用户所属科室编号) = "" Then
                If Trim(ctxtUserNo.Text) = "0000" Then
                    '显示主界面
                    Unload Me
                    lfrm主界面.Show
                    plngMainHwnd = lfrm主界面.hWnd
                Else
                    Call sffuncMsg("由于您没有配置科室，不能使用本系统！", sf警告)
                    Call ccmdCancel_Click
                End If
            Else
                Unload Me
                lfrm主界面.Show
                plngMainHwnd = lfrm主界面.hWnd
            End If
        Else
            If Trim(ctxtUserNo.Text) = "0000" Then
                Call sffuncMsg("该工作站尚未注册，只有系统管理员能在未注册的工作站上登录，请先进入工作站管理注册后再使用本系统！", sf警告)
                '显示主界面
                Unload Me
                lfrm主界面.Show
                plngMainHwnd = lfrm主界面.hWnd
            Else
                Call sffuncMsg("该工作站尚未注册，请先注册再使用本系统！", sf警告)
                Call ccmdCancel_Click
            End If
        End If
        
        Call oesubSave("用户登录系统", "登录")
    Else '非法用户或口令输错
        Call sffuncMsg("用户编号或口令输错！", sf警告)
        ctxtUserNo.SetFocus
    End If
errHandle:
    Set lfrm主界面 = Nothing
    If Err.Number = 0 Then Exit Sub
    Set pobj平台结构 = Nothing
    Call sfsub错误处理("主程序", "frmLogin", "ccmdOk_click", Err.Number, Err.Description, False)
End Sub

' 修改说明：不判断工作站是否注册，始终返回True值。
' 修改人：  罗庆
' 修改时间：2001-8-6
Private Function func是否注册(ByVal para用户编号 As String) As Boolean
    func是否注册 = True
    Exit Function
    If um工作站编号 = "" Then
        func是否注册 = False
    End If
End Function

'修改：2002-1-28（杨春）从注册表中获取上次登录的用户编号。
Private Sub Form_Load()
    Dim lstrUser As String
    
    On Error Resume Next
    lstrUser = sffuncGetSetting("系统管理", "本地配置", "用户编号")
    If lstrUser = "" Then lstrUser = "0001"
    ctxtUserNo.Text = lstrUser
    'Label1 = "请输入用户编号和口令以登录" & pstrSysName
    
End Sub
