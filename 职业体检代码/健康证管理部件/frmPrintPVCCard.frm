VERSION 5.00
Object = "{D9347025-9612-11D1-9D75-00C04FCC8CDC}#1.0#0"; "MSBCODE9.OCX"
Begin VB.Form frmPrintPVCCard 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3585
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6945
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   6945
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton ccmdNext 
      Caption         =   "下一张(&>)"
      Height          =   495
      Left            =   5520
      TabIndex        =   24
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton ccmdPrve 
      Caption         =   "上一张(&<)"
      Height          =   495
      Left            =   5520
      TabIndex        =   23
      Top             =   840
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   240
      Picture         =   "frmPrintPVCCard.frx":0000
      ScaleHeight     =   135
      ScaleWidth      =   5055
      TabIndex        =   22
      Top             =   3120
      Width           =   5055
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   240
      Picture         =   "frmPrintPVCCard.frx":0433
      ScaleHeight     =   495
      ScaleWidth      =   5055
      TabIndex        =   21
      Top             =   240
      Width           =   5055
   End
   Begin VB.CommandButton ccmdCancel 
      Caption         =   "退出(&E)"
      Height          =   495
      Left            =   5520
      TabIndex        =   2
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton ccmdPrint 
      Caption         =   "打印(&P)"
      Height          =   495
      Left            =   5520
      TabIndex        =   1
      Top             =   2280
      Width           =   1215
   End
   Begin VB.PictureBox PICCard 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   240
      ScaleHeight     =   3015
      ScaleWidth      =   5055
      TabIndex        =   0
      Top             =   240
      Width           =   5055
      Begin VB.Label clbl血号 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3720
         TabIndex        =   25
         Top             =   2280
         Width           =   615
      End
      Begin VB.Image cPhoto 
         Height          =   1605
         Left            =   3360
         Stretch         =   -1  'True
         Top             =   600
         Width           =   1200
      End
      Begin VB.Label cinfo发证机构2 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   960
         TabIndex        =   20
         Top             =   2230
         Width           =   2655
      End
      Begin VB.Label cinfo发证机构1 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   960
         TabIndex        =   19
         Top             =   2040
         Width           =   4335
      End
      Begin VB.Label cinfo编号 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   600
         TabIndex        =   18
         Top             =   1770
         Width           =   3375
      End
      Begin VB.Label cinfo有效期止 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   960
         TabIndex        =   17
         Top             =   1515
         Width           =   2775
      End
      Begin VB.Label cinfo有效期起 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   960
         TabIndex        =   16
         Top             =   1245
         Width           =   2775
      End
      Begin VB.Label cinfo年龄 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   2760
         TabIndex        =   15
         Top             =   990
         Width           =   735
      End
      Begin VB.Label cinfo工种 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   960
         TabIndex        =   14
         Top             =   990
         Width           =   975
      End
      Begin VB.Label cinfo性别 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   2760
         TabIndex        =   13
         Top             =   720
         Width           =   735
      End
      Begin VB.Label cinfo姓名 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   600
         TabIndex        =   12
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "至"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   720
         TabIndex        =   10
         Top             =   1515
         Width           =   195
      End
      Begin VB.Label cLab健康证号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "编号："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   9
         Top             =   1770
         Width           =   585
      End
      Begin VB.Label cLab发证日期 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "发证机构："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   8
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label cLab工种 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "从业类别："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   7
         Top             =   990
         Width           =   975
      End
      Begin VB.Label cLab年龄 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   2280
         TabIndex        =   6
         Top             =   990
         Width           =   585
      End
      Begin VB.Label cLab性别 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   2280
         TabIndex        =   5
         Top             =   720
         Width           =   585
      End
      Begin VB.Label cLab姓名 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   585
      End
      Begin VB.Label cLbl体检 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "有效期限："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   3
         Top             =   1245
         Width           =   975
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000002&
         Height          =   1605
         Left            =   3360
         Top             =   600
         Width           =   1200
      End
      Begin BARCODELibCtl.BarCodeCtrl cBar健康证编号 
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   2400
         Width           =   4335
         Style           =   7
         SubStyle        =   -1
         Validation      =   0
         LineWeight      =   1
         Direction       =   0
         ShowData        =   0
         Value           =   "123456 Code-128"
         ForeColor       =   0
         BackColor       =   16777215
      End
   End
End
Attribute VB_Name = "frmPrintPVCCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Cards As New Collection
Private m发证单位 As String             '定义变量记录发证单位

Dim intIndex As Long

Private Sub Command2_Click()

End Sub

Private Sub ccmdCancel_Click()
On Error GoTo errordeal

'    IFPrint = False
    Me.Hide
    dafuncGetData ("update 系统管理_系统编号生成记录表 set 当前值=当前值-" & Cards.Count & " where 业务名称='健康证管理' and 编号名称='健康证编号'")
    Set Cards = New Collection
    Unload Me
Exit Sub

errordeal:
    MsgBox Err.Description, vbInformation, "健康证系统"
    On Error Resume Next
End Sub


Private Sub FillForm(ByVal Index As Long)
On Error GoTo errordeal
   Dim lstr有效期限 As String
   Dim lstr前缀 As String
   Dim lobjtemp As Object
    m发证单位 = um防疫站名
    Set lobjtemp = dafuncGetData("select top 1 回函地址 from 健康证_业务配置表")
    If lobjtemp.RecordCount > 0 Then
    lstr前缀 = lobjtemp(0)
    Else
        lstr前缀 = ""
    End If
    If Index >= 1 And Index <= Cards.Count Then
        cinfo姓名.Caption = Cards(Index).姓名
        cinfo性别.Caption = Cards(Index).性别
        cinfo年龄.Caption = Cards(Index).年龄
        cinfo工种.Caption = Cards(Index).种类
        
        cinfo有效期起 = Left(Cards(Index).发证日期, 4) + "年" + Mid(Cards(Index).发证日期, 6, 2) + "月" + Right(Cards(Index).发证日期, 2) + "日"
        lstr有效期限 = DateAdd("d", -1, DateAdd("yyyy", 1, Cards(Index).发证日期))
        cinfo有效期止 = Left(lstr有效期限, 4) + "年" + Mid(lstr有效期限, 6, 2) + "月" + Right(lstr有效期限, 2) + "日"
        cinfo编号.Caption = "川" + "(" + Left(Cards(Index).发证日期, 4) + ")" + lstr前缀 + "-" + Cards(Index).健康证号
'        cinfo编号.Caption = "川" + lstr前缀 + "(" + Left(Cards(Index).Date, 4) + ")第" + "00000000" + "号"
        cinfo发证机构1.Caption = Left(m发证单位, 12)
        If Len(m发证单位) > 12 Then
            cinfo发证机构2.Caption = Right(m发证单位, Len(m发证单位) - 12)
        End If
        
'        cBar健康证编号.Value = Right(Cards(Index).SN, 8)
        cBar健康证编号.Value = IIf(Cards(Index).二代身份证编号 = "", Cards(Index).体检系统编号, Cards(Index).二代身份证编号)
'        cBar健康证编号.Value = "AFEdU/ZW13N0UDAA"

        cPhoto.Picture = Cards(Index).照片
'        clabSysNo.Caption = Cards(Index).体检系统编号
        clbl血号.Caption = Cards(Index).民族
    End If
    

Exit Sub

errordeal:
    MsgBox Err.Description, vbInformation, "健康证系统"
    On Error Resume Next

End Sub

Private Sub ccmdNext_Click()
intIndex = intIndex + 1
If intIndex = Cards.Count Then
    ccmdNext.Enabled = False
End If
If Cards.Count > 1 Then
    ccmdPrve.Enabled = True
   
End If

 FillForm intIndex
End Sub

Private Sub ccmdPrint_Click()
On Error GoTo errordeal
Dim i As Integer
'    IFPrint = True
Me.Hide
Me.BackColor = vbWhite
PICCard.Top = 140
PICCard.Left = 150
Picture1.Visible = False
Picture2.Visible = False

For i = 1 To Cards.Count
    FillForm i

    Me.PrintForm
    
    dafuncGetData "Update 健康证管理_办证申请信息表  Set 状态='已打印',健康证号 ='" & Cards(i).健康证号 & "',身份证号 ='" & Cards(i).身份证号 & "',发证日期='" & Cards(i).发证日期 & "',有效期至='" & Cards(i).有效期至 & "', 发证单位='" & Cards(i).发证单位 & "' where 系统编号 ='" & Cards(i).系统编号 & "'"
Next
    
Set Cards = New Collection
Exit Sub

errordeal:
    MsgBox Err.Description, vbInformation, "健康证系统"
    On Error Resume Next
End Sub

Private Sub ccmdPrve_Click()
intIndex = intIndex - 1
If intIndex = 1 Then
    ccmdPrve.Enabled = False
End If
If intIndex < Cards.Count Then
    ccmdNext.Enabled = True
   
End If

 FillForm intIndex
End Sub

Private Sub Form_Load()
On Error GoTo errordeal
    
    intIndex = 1

    
    If Cards.Count >= 1 Then
        FillForm intIndex
    End If
    ccmdPrve.Enabled = False
    If Cards.Count = 1 Then
        ccmdNext.Enabled = False
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
