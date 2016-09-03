VERSION 5.00
Begin VB.Form frm密码修改 
   BackColor       =   &H00E4E8C6&
   Caption         =   "口令修改"
   ClientHeight    =   2565
   ClientLeft      =   1530
   ClientTop       =   1800
   ClientWidth     =   4800
   Icon            =   "frm密码修改.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2565
   ScaleWidth      =   4800
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox ctxtOldPassWord 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2730
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   330
      Width           =   1335
   End
   Begin VB.TextBox ctxtFirstPassWord 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2730
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox ctxtSecPassWord 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2730
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1350
      Width           =   1335
   End
   Begin VB.Image ccmdOk 
      Height          =   300
      Left            =   1020
      Picture         =   "frm密码修改.frx":0442
      Top             =   1965
      Width           =   945
   End
   Begin VB.Image ccmdCancel 
      Height          =   300
      Left            =   2895
      Picture         =   "frm密码修改.frx":2C50
      Top             =   1965
      Width           =   945
   End
   Begin VB.Image Image1 
      Height          =   510
      Left            =   120
      Picture         =   "frm密码修改.frx":5409
      Stretch         =   -1  'True
      Top             =   120
      Width           =   510
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "请输入原来的口令："
      Height          =   180
      Left            =   990
      TabIndex        =   5
      Top             =   360
      Width           =   1620
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "请输入新口令："
      Height          =   180
      Left            =   990
      TabIndex        =   4
      Top             =   870
      Width           =   1260
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "请再次输入新口令："
      Height          =   180
      Left            =   990
      TabIndex        =   3
      Top             =   1350
      Width           =   1620
   End
End
Attribute VB_Name = "frm密码修改"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub ccmdCancel_Click()
    Unload Me
End Sub


Private Sub ccmdOk_Click()
    On Error GoTo errHandle
    If ctxtSecPassWord.Text <> ctxtFirstPassWord.Text Then  '两次输入的口令是否一致
        Call sffuncMsg("两次输入的新口令不一样，请重新输入！", sf警告)
        ctxtFirstPassWord.SetFocus
        Exit Sub
    Else
        If umfunc修改口令(ctxtOldPassWord.Text, ctxtFirstPassWord.Text) Then
            Call sffuncMsg("修改成功，请记住您的新口令！", sf警告)
            Unload Me
        Else
            Call sffuncMsg("口令修改失败，确认旧口令是否正确！", sf警告)
            ctxtOldPassWord.SetFocus
            Exit Sub
        End If
    End If
    Exit Sub
errHandle:
    Call sfsub错误处理("主程序", "frm密码修改", "ccmdOk_Click", Err.Number, Err.Description, False)
End Sub

Private Sub ctxtFirstPassWord_GotFocus()
    ctxtFirstPassWord.SelStart = 0
    ctxtFirstPassWord.SelLength = Len(ctxtFirstPassWord.Text)
End Sub

'响应回车键
Private Sub ctxtFirstPassWord_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ctxtSecPassWord.SetFocus
    End If
End Sub

'去掉两端的空格
Private Sub ctxtFirstPassWord_LostFocus()
    ctxtFirstPassWord.Text = Trim(ctxtFirstPassWord.Text)
End Sub


Private Sub ctxtOldPassWord_GotFocus()
    ctxtOldPassWord.SelStart = 0
    ctxtOldPassWord.SelLength = Len(ctxtOldPassWord.Text)
    
End Sub

'响应回车键
Private Sub ctxtOldPassWord_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
         ctxtFirstPassWord.SetFocus
    End If
End Sub


Private Sub ctxtSecPassWord_GotFocus()
    ctxtSecPassWord.SelStart = 0
    ctxtSecPassWord.SelLength = Len(ctxtSecPassWord.Text)
End Sub


'响应回车键
Private Sub ctxtSecPassWord_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call ccmdOk_Click
    End If
End Sub

'去掉两端的空格
Private Sub ctxtSecPassWord_LostFocus()
    ctxtSecPassWord.Text = Trim(ctxtSecPassWord.Text)
End Sub


