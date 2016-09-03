VERSION 5.00
Begin VB.Form frmPrintCard 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4785
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7770
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   7770
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame FrameM 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2625
      Index           =   0
      Left            =   120
      TabIndex        =   17
      Top             =   1080
      Width           =   5010
      Begin VB.Label clbl检出病种 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   38
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Line Line5 
         X1              =   2880
         X2              =   3240
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label cLab年龄 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   7.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   2400
         TabIndex        =   37
         Top             =   1080
         Width           =   450
      End
      Begin VB.Line Line4 
         X1              =   1200
         X2              =   2160
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line3 
         X1              =   1440
         X2              =   3240
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Label cLbl体检单位 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "健康体检单位："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   7.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   360
         TabIndex        =   36
         Top             =   2160
         Width           =   1050
      End
      Begin VB.Line Line1 
         X1              =   1200
         X2              =   3240
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label cinfo备注 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   150
         Left            =   3720
         TabIndex        =   35
         Top             =   2040
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label cInfo民族 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   150
         Left            =   3960
         TabIndex        =   34
         Top             =   2160
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label cInfo单位 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   150
         Left            =   2280
         TabIndex        =   33
         Top             =   2040
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label cInfo试管编号 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   300
         TabIndex        =   32
         Top             =   1920
         Width           =   90
      End
      Begin VB.Label cInfo发证日期 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   150
         Left            =   1200
         TabIndex        =   31
         Top             =   1680
         Width           =   90
      End
      Begin VB.Image Cphoto 
         Height          =   1425
         Left            =   3660
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1080
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00E38B5B&
         BorderWidth     =   3
         Height          =   1455
         Index           =   0
         Left            =   3660
         Top             =   360
         Width           =   1110
      End
      Begin VB.Label cLab发证日期 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "发证日期："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   7.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   360
         TabIndex        =   30
         Top             =   1800
         Width           =   750
      End
      Begin VB.Line cLne发证日期 
         Index           =   0
         X1              =   1200
         X2              =   3240
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label cLab健康证号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "编   号："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   7.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Index           =   0
         Left            =   360
         TabIndex        =   29
         Top             =   360
         Width           =   675
      End
      Begin VB.Label cLab工种 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "类   别："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   7.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Index           =   0
         Left            =   360
         TabIndex        =   28
         Top             =   1080
         Width           =   675
      End
      Begin VB.Label cInfo工种 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   150
         Left            =   1200
         TabIndex        =   27
         Top             =   960
         Width           =   90
      End
      Begin VB.Label cInfo年龄 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   150
         Left            =   2880
         TabIndex        =   26
         Top             =   960
         Width           =   90
      End
      Begin VB.Line cLne性别 
         Index           =   0
         X1              =   2880
         X2              =   3240
         Y1              =   800
         Y2              =   800
      End
      Begin VB.Label cInfo性别 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   150
         Left            =   2880
         TabIndex        =   25
         Top             =   600
         Width           =   90
      End
      Begin VB.Label cLab性别 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   7.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Index           =   0
         Left            =   2400
         TabIndex        =   24
         Top             =   720
         Width           =   450
      End
      Begin VB.Line cLne姓名 
         Index           =   0
         X1              =   1200
         X2              =   2160
         Y1              =   800
         Y2              =   800
      End
      Begin VB.Label cInfo姓名 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   150
         Left            =   1200
         TabIndex        =   23
         Top             =   600
         Width           =   90
      End
      Begin VB.Label cLab姓名 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓    名："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   7.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Index           =   0
         Left            =   360
         TabIndex        =   22
         Top             =   660
         Width           =   750
      End
      Begin VB.Label cinfo证号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1200
         TabIndex        =   21
         Top             =   240
         Width           =   60
      End
      Begin VB.Label cLbl体检 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "体    检："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   7.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Index           =   0
         Left            =   360
         TabIndex        =   20
         Top             =   1395
         Width           =   750
      End
      Begin VB.Label cLbl合格 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "合格"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   7.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Index           =   0
         Left            =   3720
         TabIndex        =   19
         Top             =   2160
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label Cinfo体检单位 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   1440
         TabIndex        =   18
         Top             =   2040
         Width           =   90
      End
      Begin VB.Line cLne体检 
         Index           =   0
         X1              =   1200
         X2              =   3240
         Y1              =   1560
         Y2              =   1560
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "选择模板"
      Height          =   2295
      Left            =   5280
      TabIndex        =   7
      Top             =   1080
      Width           =   2415
      Begin VB.Frame cframPos 
         Caption         =   "起始张"
         Height          =   735
         Left            =   60
         TabIndex        =   11
         Top             =   1560
         Visible         =   0   'False
         Width           =   2295
         Begin VB.ComboBox ccmbIndex 
            Height          =   300
            ItemData        =   "frmPrintCard.frx":0000
            Left            =   1320
            List            =   "frmPrintCard.frx":0013
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   360
            Width           =   600
         End
         Begin VB.ComboBox ccmbSide 
            Height          =   300
            ItemData        =   "frmPrintCard.frx":0026
            Left            =   120
            List            =   "frmPrintCard.frx":0030
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   360
            Width           =   600
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "张"
            Height          =   180
            Index           =   1
            Left            =   2000
            TabIndex        =   15
            Top             =   480
            Width           =   180
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "边第"
            Height          =   180
            Index           =   0
            Left            =   800
            TabIndex        =   14
            Top             =   480
            Width           =   360
         End
      End
      Begin VB.OptionButton opt模板类型 
         Caption         =   "单张打印"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Width           =   1095
      End
      Begin VB.OptionButton opt模板类型 
         Caption         =   "1 * 5(从上到下)"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   1695
      End
      Begin VB.OptionButton opt模板类型 
         Caption         =   "2 * 5(从左到右)"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Value           =   -1  'True
         Width           =   1695
      End
   End
   Begin VB.CommandButton cComNext 
      Caption         =   "下一张>&>"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cComPrev 
      Caption         =   "&<<上一张"
      Enabled         =   0   'False
      Height          =   375
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cComCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(Esc)"
      Height          =   375
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cComPrintCard 
      Caption         =   "打印(&P)"
      Default         =   -1  'True
      Height          =   375
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label cInfo有效截止日期 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   840
      TabIndex        =   16
      Top             =   3480
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   6600
      Picture         =   "frmPrintCard.frx":003C
      Top             =   120
      Width           =   480
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   7800
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "系统在此处预览健康证的效果并提供最后的打印确认。你可以通过单击""上一张""、""下一张""预览所有你将要打印的健康证。"
      Height          =   495
      Left            =   390
      TabIndex        =   5
      Top             =   120
      Width           =   5295
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "以上为健康证的打印模版，单击""打印""确认，单击""取消""退出。"
      Height          =   180
      Left            =   480
      TabIndex        =   2
      Top             =   3840
      Width           =   5040
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   120
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   7755
   End
End
Attribute VB_Name = "frmPrintCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Cards As New Collection
Private m发证单位 As String             '定义变量记录发证单位

Dim intIndex As Long

Public pblnPrint As Boolean

Private Sub cComCancel_Click()

On Error GoTo errordeal

    Me.Hide
    Set Cards = New Collection
    pblnPrint = False
Exit Sub

errordeal:
    MsgBox Err.Description, vbInformation, "健康证系统"
    On Error Resume Next
End Sub

Private Sub cComNext_Click()

On Error GoTo errordeal

    intIndex = intIndex + 1
    FillForm intIndex
    
Exit Sub

errordeal:
    MsgBox Err.Description, vbInformation, "健康证系统"
    On Error Resume Next

End Sub

Private Sub cComPrev_Click()

On Error GoTo errordeal

    intIndex = intIndex - 1
    FillForm intIndex

Exit Sub

errordeal:
    MsgBox Err.Description, vbInformation, "健康证系统"
    On Error Resume Next

End Sub

Private Sub FillForm(ByVal Index As Long)
On Error GoTo errordeal
    
    Dim lobjRectemp As Object       '定义变量记录临时记录集
    
    If Index >= 1 And Index <= Cards.Count Then
        cInfo姓名.Caption = Cards(Index).姓名
        cInfo性别.Caption = Cards(Index).性别
        cInfo年龄.Caption = Cards(Index).年龄
        
        cInfo工种.Caption = Cards(Index).种类
        Cinfo体检单位.Caption = Cards(Index).发证单位
        cInfo发证日期.Caption = Cards(Index).发证日期
        
        clbl检出病种.Caption = IIf(Cards(Index).检出病种 = "无", "无从业禁忌症", Cards(Index).检出病种)
        cinfo证号.Caption = Cards(Index).健康证号
        cInfo单位.Caption = Cards(Index).单位名称
        cInfo民族.Caption = Cards(Index).民族
        cinfo备注.Caption = IIf(Cards(Index).检出病种 = "正常", "", Cards(Index).检出病种)
        Cphoto.Picture = Cards(Index).照片
    End If
    
    If Index = 1 Then
        cComPrev.Enabled = False
        cComNext.Enabled = True
    ElseIf Index = Cards.Count Then
        cComNext.Enabled = False
        cComPrev.Enabled = True
    ElseIf Index > 1 And Index < Cards.Count Then
        cComPrev.Enabled = True
        cComNext.Enabled = True
    ElseIf Cards.Count = 1 Then
        cComPrev.Enabled = False
        cComNext.Enabled = False
    End If

Exit Sub

errordeal:
    MsgBox Err.Description, vbInformation, "健康证系统"
    On Error Resume Next
    Resume
End Sub

Private Sub cComPrintCard_Click()

On Error GoTo errordeal


    Me.Hide
    
    pblnPrint = True
Exit Sub

errordeal:
    MsgBox Err.Description, vbInformation, "健康证系统"
    On Error Resume Next

End Sub



Private Sub Form_Activate()
'    intIndex = 1
'    Select Case Cards.Count
'    Case 1
'        cComNext.Enabled = False
'        cComPrev.Enabled = False
'    Case Is > 1
'        cComNext.Enabled = True
'        cComPrev.Enabled = False
'    End Select
'
'    If Cards.Count <= 5 Then
'        opt模板类型(1).Value = True
'    ElseIf Cards.Count > 5 Then
'        opt模板类型(0).Value = True
'    End If
'
'    If Cards.Count >= 1 Then
'        FillForm intIndex
'    End If
'
'    '修改：2002-5-23（增加打印位置设置）。
'    If Cards.Count < 10 Then
'        cframPos.Visible = True
'        cframPos.Enabled = True
'    Else
'        cframPos.Visible = False
'        cframPos.Enabled = False
'    End If
End Sub

Private Sub Form_Load()

On Error GoTo errordeal
    
    intIndex = 1
    Select Case Cards.Count
    Case 1
        cComNext.Enabled = False
        cComPrev.Enabled = False
    Case Is > 1
        cComNext.Enabled = True
        cComPrev.Enabled = False
    End Select
    If Cards.Count = 1 Then
        opt模板类型(2).Value = True
    ElseIf Cards.Count <= 5 Then
        opt模板类型(1).Value = True
    ElseIf Cards.Count > 5 Then
        opt模板类型(0).Value = True
    End If
    
    If Cards.Count >= 1 Then
        FillForm intIndex
    End If
    
    '修改：2002-5-23（增加打印位置设置）。
    ccmbSide.ListIndex = 0
    ccmbIndex.ListIndex = 0
    
    If Cards.Count < 10 Then
        cframPos.Visible = True
        cframPos.Enabled = True
    Else
        cframPos.Visible = False
        cframPos.Enabled = False
    End If
    
    
Exit Sub

errordeal:
    MsgBox Err.Description, vbInformation, "健康证系统"
    On Error Resume Next

End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error GoTo errordeal

    Set Cards = Nothing

Exit Sub

errordeal:
    MsgBox Err.Description, vbInformation, "健康证系统"
    On Error Resume Next

End Sub

'修改：2002-5-23（增加打印位置设置）。
Private Sub opt模板类型_Click(Index As Integer)
    Dim i As Long
    Dim lstrCN As String
    If Index = 0 Then
        If Cards.Count < 10 Then
            cframPos.Visible = True
            cframPos.Enabled = True
        Else
            cframPos.Visible = False
            cframPos.Enabled = False
        End If
        
    Else
        cframPos.Visible = False
        cframPos.Enabled = False
        
    End If
    
    lstrCN = Cards(1).健康证号
    If opt模板类型(0).Value Then
        For i = 2 To Cards.Count
            lstrCN = Format(Val(lstrCN) + 1, String(Len(lstrCN), "0"))
            Cards(i).健康证号 = lstrCN
        Next
        FillForm intIndex
    ElseIf opt模板类型(1).Value Then
        For i = 2 To Cards.Count
            lstrCN = Format(Val(lstrCN) + 2, String(Len(lstrCN), "0"))
            Cards(i).健康证号 = lstrCN
        Next
        FillForm intIndex
    End If
End Sub
