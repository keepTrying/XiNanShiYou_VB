VERSION 5.00
Begin VB.Form FrmLogin 
   BackColor       =   &H00F0F0EC&
   BorderStyle     =   0  'None
   Caption         =   "用户登录"
   ClientHeight    =   2520
   ClientLeft      =   1710
   ClientTop       =   1950
   ClientWidth     =   5400
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "frmLogin.frx":0000
   ScaleHeight     =   2520
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox ctxtUserNo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
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
      Left            =   1740
      MaxLength       =   8
      TabIndex        =   0
      Top             =   1180
      Width           =   2100
   End
   Begin VB.TextBox ctxtPassWord 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
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
      Left            =   1740
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1780
      Width           =   2100
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "请输入用户编号和口令以登录道源卫生防疫信息管理系统"
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label ccmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3960
      TabIndex        =   3
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label ccmdOk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3960
      TabIndex        =   2
      Top             =   1080
      Width           =   855
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

    Unload Me
End Sub

Private Sub ccmdOk_Click()
    Dim lstr模式 As String
    
    On Error GoTo errHandle
    Dim lfrm主界面 As Form
    
    '验证身份
    If umfunc校验身份(Trim(ctxtUserNo.Text), ctxtPassWord.Text) Then
        '合法用户
        '修改：2002-1-28（登记配置表，当前用户编号）。
        On Error Resume Next
        
        Set lfrm主界面 = New frmMain
        
        On Error GoTo errHandle
        If Trim(um用户所属科室编号) = "" Then
            If Trim(ctxtUserNo.Text) = "0000" Then
                '显示主界面
                Unload Me
                lfrm主界面.Show
            Else
                Call sffuncMsg("由于您没有配置科室，不能使用本系统！", sf警告)
                Call ccmdCancel_Click
            End If
        Else
            Unload Me
            lfrm主界面.Show
        End If
        
    Else '非法用户或口令输错
        Call sffuncMsg("用户编号或口令输错！", sf警告)
        ctxtUserNo.SetFocus
    End If
errHandle:
    Set lfrm主界面 = Nothing
    If Err.Number = 0 Then Exit Sub

    Call sfsub错误处理("体检数据导入导出", "frmLogin", "ccmdOk_click", Err.Number, Err.Description, False)
End Sub



'修改：2002-1-28（杨春）从注册表中获取上次登录的用户编号。
Private Sub Form_Load()
    Dim lstrUser As String
    
    On Error Resume Next
    lstrUser = sffuncGetSetting("系统管理", "本地配置", "用户编号")
    If lstrUser = "" Then lstrUser = "0001"
    ctxtUserNo.Text = lstrUser
    Label1 = "请输入用户编号和口令以登录体检输入导入导出工具"
    
End Sub
