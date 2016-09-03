VERSION 5.00
Begin VB.Form frmdrawline 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "纯音听阈测试"
   ClientHeight    =   9225
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9225
   ScaleWidth      =   9585
   StartUpPosition =   2  '屏幕中心
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "骨导：X"
      Height          =   255
      Left            =   8160
      TabIndex        =   24
      Top             =   120
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "气导：O"
      Height          =   255
      Left            =   6960
      TabIndex        =   23
      Top             =   120
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "体检结果录入"
      Height          =   3135
      Left            =   360
      TabIndex        =   7
      Top             =   6000
      Width           =   8775
      Begin VB.CommandButton ccmdsave 
         BackColor       =   &H00FFFFFF&
         Caption         =   "保存结果"
         Height          =   495
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   2400
         Width           =   2415
      End
      Begin VB.TextBox ctxt斯氏 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4800
         TabIndex        =   19
         Top             =   1920
         Width           =   3015
      End
      Begin VB.TextBox ctxt韦氏 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4800
         TabIndex        =   17
         Top             =   1440
         Width           =   3015
      End
      Begin VB.TextBox ctxt任氏 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4800
         TabIndex        =   15
         Top             =   960
         Width           =   3015
      End
      Begin VB.TextBox ctxt备注 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1380
         Left            =   1200
         TabIndex        =   12
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox ctxt语音测试 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1200
         TabIndex        =   11
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox ctxt耳语测试 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1200
         TabIndex        =   8
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "左         右"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5640
         TabIndex        =   21
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         Caption         =   "斯氏:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4200
         TabIndex        =   20
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "韦氏:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4200
         TabIndex        =   18
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "任氏:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4200
         TabIndex        =   16
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "音叉测试:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3840
         TabIndex        =   14
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "备注:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "语音测试:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "耳语测试:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   360
      Picture         =   "frmdrawline.frx":0000
      ScaleHeight     =   4335
      ScaleWidth      =   8835
      TabIndex        =   6
      Top             =   720
      Width           =   8835
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "保存为左耳测试结果"
      Height          =   495
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5400
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "保存为左耳测试结果"
      Height          =   495
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5400
      Width           =   2295
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFFFF&
      Caption         =   "传导方式："
      Height          =   225
      Left            =   5880
      TabIndex        =   28
      Top             =   120
      Width           =   975
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   8280
      X2              =   9000
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line1 
      X1              =   6960
      X2              =   7680
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFFFF&
      Caption         =   "线条宽度："
      Height          =   225
      Left            =   5880
      TabIndex        =   27
      Top             =   360
      Width           =   975
   End
   Begin VB.Label lblno 
      BackColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   1440
      TabIndex        =   26
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFFFF&
      Caption         =   "系统编号："
      Height          =   225
      Left            =   360
      TabIndex        =   25
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "未保存"
      Height          =   225
      Left            =   8640
      TabIndex        =   3
      Top             =   5160
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "听觉损失单位Db  右："
      Height          =   225
      Left            =   6840
      TabIndex        =   2
      Top             =   5160
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "未保存"
      Height          =   225
      Left            =   2160
      TabIndex        =   1
      Top             =   5160
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "听觉损失单位Db  左："
      Height          =   225
      Left            =   480
      TabIndex        =   0
      Top             =   5160
      Width           =   1935
   End
End
Attribute VB_Name = "frmdrawline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim imgleft As StdPicture
Dim imgright As StdPicture
'把图片先预存到imgleft/imgright 最后和其它信息一起存入数据库。
Dim pidleft As String
Dim pidright As String
Dim lobjRec照片left As Object
Dim lobjRec照片right As Object
Dim tip As Integer
    
Private Sub ccmdsave_Click()
Dim sqlstr As String
sqlstr = "insert into 职业病体检_结果信息_纯音听阈测试(系统编号,传导方式,耳语测试,语音测试,备注) values( '" & lblno.Caption & "','','" & ctxt耳语测试.Text & "','" & ctxt语音测试.Text & "','" & ctxt备注.Text & "')"
dafuncGetData (sqlstr)
End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
Check2.Value = 0
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
Check1.Value = 0
Picture1.DrawWidth = 2
End If
End Sub

Private Sub Command1_Click()
pidleft = lblno.Caption + "left"
Set imgleft = Picture1.Image
Label3.Caption = "已保存!"
Label3.ForeColor = vbRed
Set lobjRec照片left = CreateObject("职业病对象.clsPersonExamed")
lobjRec照片left.func保存身份证照片 imgleft, pidleft, "职业病体检"
End Sub

Private Sub Command2_Click()
pidright = lblno.Caption + "right"
Set imgright = Picture1.Image
Label5.Caption = "已保存!"
Label5.ForeColor = vbRed
Set lobjRec照片right = CreateObject("职业病对象.clsPersonExamed")
lobjRec照片right.func保存身份证照片 imgright, pidright, "职业病体检"
End Sub
Private Sub Form_Load()
tip = 1
End Sub
Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then
Picture1.Cls
tip = 1
Picture1.DrawWidth = 1
Check1.Value = 1
ElseIf Picture1.CurrentX = 0 And Picture1.CurrentY = 0 Then
Picture1.PSet (X, Y), vbBlack
Else
    If tip < 6 Then
    Picture1.Line -(X, Y), vbBlack
    tip = tip + 1
        If Check2.Value = 1 And tip = 6 Then
            Check2.Value = 0
        End If
    ElseIf Check2.Value = 1 And tip > 5 Then
    Picture1.PSet (X, Y), vbBlack
    tip = 1
    End If
End If
End Sub

