VERSION 5.00
Begin VB.Form frmEyeDraw 
   Caption         =   "晶状体环面及正面图"
   ClientHeight    =   3870
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8070
   LinkTopic       =   "Form1"
   ScaleHeight     =   3870
   ScaleWidth      =   8070
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command1 
      Caption         =   "最粗"
      Height          =   375
      Index           =   3
      Left            =   3240
      TabIndex        =   16
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "粗"
      Height          =   375
      Index           =   2
      Left            =   2520
      TabIndex        =   15
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "中"
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   14
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "细"
      Height          =   375
      Index           =   0
      Left            =   840
      TabIndex        =   13
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton ccmdDraw 
      Caption         =   "画图"
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton ccmdEraser 
      Caption         =   "橡皮擦"
      Height          =   375
      Left            =   5280
      TabIndex        =   4
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton ccmdClearPicture 
      Caption         =   "清空此次修改"
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton ccmdSavePicture 
      BackColor       =   &H00C0FFC0&
      Caption         =   "保存图像"
      Height          =   375
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2220
      Left            =   240
      ScaleHeight     =   2160
      ScaleWidth      =   7515
      TabIndex        =   1
      Top             =   1560
      Width           =   7575
   End
   Begin VB.CommandButton ccmdLoadOriginalPicture 
      Caption         =   "载入原图"
      Height          =   375
      Left            =   6480
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFC0C0&
      Caption         =   "线条/橡皮擦粗细："
      Height          =   375
      Left            =   1320
      TabIndex        =   12
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFC0C0&
      Caption         =   "细"
      Height          =   375
      Index           =   0
      Left            =   960
      TabIndex        =   11
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0E0FF&
      Caption         =   "右眼"
      Height          =   255
      Left            =   1560
      TabIndex        =   10
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0E0FF&
      Caption         =   "左眼"
      Height          =   255
      Left            =   5400
      TabIndex        =   9
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFC0C0&
      Caption         =   "中"
      Height          =   375
      Index           =   1
      Left            =   1560
      TabIndex        =   8
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFC0C0&
      Caption         =   "粗"
      Height          =   375
      Index           =   2
      Left            =   2160
      TabIndex        =   7
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFC0C0&
      Caption         =   "最粗"
      Height          =   375
      Index           =   3
      Left            =   2760
      TabIndex        =   6
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "frmEyeDraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private DrawLineWidth As Integer        '画图线条粗细的标记
Private DrawState As Integer            '记录当前画图状态，未做修改不保存 -1、未画图0、画图1、橡皮擦2
Private mstr结果图片项目编号 As String  '标记五官科结果图片项目编号，数据库中记录的值为”01069“
Public mstr结果图片名称 As String

Private lobj批量操作对象 As Object    '为批量操作提供对象函数
Private EyeMapCheck(56050, 2) As Integer
Private pointCnt As Long

Public pubSysNo As String
Private pri业务对象 As Object

Private Sub ccmdClearPicture_Click()
    Picture3.Cls
    Picture3.ForeColor = vbRed
    Picture3.DrawWidth = DrawLineWidth
    DrawState = -1
End Sub

Private Sub ccmdDraw_Click()
    Picture3.ForeColor = vbRed
    Picture3.DrawWidth = DrawLineWidth
    DrawState = 1
End Sub

Private Sub ccmdEraser_Click()
    Picture3.DrawWidth = DrawLineWidth
    DrawState = 2
End Sub

'修改人：罗李奎 2012-12-10 ↓
'说明：得到用户画线的粗细
'bug好：0000019
Private Sub Command1_Click(Index As Integer)
    DrawLineWidth = Pow_2(Index)
    Picture3.DrawWidth = DrawLineWidth
End Sub
'修改人：罗李奎 2012-12-10 ↑



Private Sub Form_Load()
    sub公共变量初始化
    sub原图解析
    sub载入已有图片
End Sub

'Private Sub Label9_Click(Index As Integer)
'    DrawLineWidth = Pow_2(Index)
'    Picture3.DrawWidth = DrawLineWidth
'End Sub

Private Function Pow_2(ByVal paraExp As Integer) As Integer
    Dim i, resultTmp As Integer     'paraExp最好小于10,否则会溢出
    resultTmp = 1
    If paraExp > 10 Then paraExp = 10
    For i = 1 To paraExp: resultTmp = resultTmp * 2: Next
    Pow_2 = resultTmp
End Function

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CurrentX = X: CurrentY = Y
    If DrawState = -1 Then ccmdDraw_Click
    
End Sub

Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And DrawState = 1 Then
        Picture3.Line (CurrentX, CurrentY)-(X, Y)
    End If
    If Button = 1 And DrawState = 2 Then
        Picture3.ForeColor = vbWhite
        Picture3.MousePointer = 4
        Picture3.Line (CurrentX, CurrentY)-(X, Y)
    End If
    CurrentX = X
    CurrentY = Y
    'MousePointer = 2   '十字架形鼠标指针
End Sub

'复杂度略高啊~~~~~~~6040像素
Private Sub Picture3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If DrawState <> 2 Then Exit Sub
    
    Dim i, xx, yy As Integer
    For i = 1 To pointCnt
        xx = EyeMapCheck(i, 1)
        yy = EyeMapCheck(i, 2)
        Call SetPixel(Picture3.hdc, xx, yy, 0)
    Next
    Picture3.Refresh
End Sub


Private Sub ccmdSavePicture_Click()
    pri业务对象.func保存结果图片 Picture3.Image, pubSysNo, mstr结果图片项目编号, Now    '01069 是眼睛体检结果图的项目编号。在“体检项目设置”的表里有记录。
    MsgBox ("结果图片保存成功！")
    
    DrawState = 0
    Unload Me
    Exit Sub
End Sub

Private Sub ccmdLoadOriginalPicture_Click()
    Dim lobjRec As Object
    Dim isOk As Integer
    Set lobjRec = pri业务对象.func查找结果图片(pubSysNo, mstr结果图片项目编号)
    If lobjRec.RecordCount > 0 Then
        isOk = MsgBox("载入原图会删除现有的结果，确定继续吗？", vbOKCancel)
        If isOk = 2 Then Exit Sub
        pri业务对象.func删除结果图片 pubSysNo, mstr结果图片项目编号
        DrawState = 0
    End If
    Picture3.Picture = pri业务对象.func获取结果图片(pubSysNo, mstr结果图片项目编号, "晶状体环面及正面图.bmp")
End Sub

Sub sub公共变量初始化()
    '画图部分设定
    DrawState = -1       'form_load时，画图状态为“未做修改不保存”
    DrawLineWidth = 2
    
    Set pri业务对象 = CreateObject("职业病体检结果录入.clscommon")
    
    
    mstr结果图片项目编号 = pri业务对象.func获取体检项目编号(mstr结果图片名称)
    
End Sub

Sub sub原图解析()
    Dim i, j, rows, cols As Integer
    Picture3.ScaleMode = 3
    
    Set Picture3.Picture = LoadPicture(App.Path & "\晶状体环面及正面图.bmp")
    cols = Picture3.ScaleWidth - 1
    rows = Picture3.ScaleHeight - 1

    pointCnt = 0
    For i = 1 To cols
        For j = 1 To rows
            'true时像素为黑色，false时像素为白色
            If Hex(GetPixel(Picture3.hdc, i, j)) = Hex(&H0) Then
                pointCnt = pointCnt + 1
                EyeMapCheck(pointCnt, 1) = i
                EyeMapCheck(pointCnt, 2) = j
            End If
        Next
    Next
    Set Picture3.Picture = Nothing
End Sub

Sub sub载入已有图片()
    Picture3.Picture = pri业务对象.func获取结果图片(pubSysNo, mstr结果图片项目编号, "晶状体环面及正面图.bmp")  '01069是眼睛检查结果图的项目编号。
    Picture3.AutoSize = True
    frmEyeDraw.Width = 2.5 * Picture3.Left + Picture3.Width
    frmEyeDraw.Height = Picture3.Top + Picture3.Height + 600
End Sub
