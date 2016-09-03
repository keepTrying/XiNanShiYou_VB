VERSION 5.00
Object = "{D9347025-9612-11D1-9D75-00C04FCC8CDC}#1.0#0"; "MSBCODE9.OCX"
Begin VB.Form frm打印体检登记表 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9945
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   9945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin BARCODELibCtl.BarCodeCtrl cbccMain 
      Height          =   765
      Left            =   4320
      TabIndex        =   10
      Top             =   240
      Width           =   3015
      Style           =   7
      SubStyle        =   0
      Validation      =   0
      LineWeight      =   3
      Direction       =   0
      ShowData        =   1
      Value           =   "123456 Code-128"
      ForeColor       =   0
      BackColor       =   16777215
   End
   Begin VB.Label clbl民族 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8400
      TabIndex        =   9
      Top             =   2280
      Width           =   120
   End
   Begin VB.Label clbl从业类别 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7200
      TabIndex        =   8
      Top             =   2880
      Width           =   120
   End
   Begin VB.Label clblUnit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2250
      TabIndex        =   7
      Top             =   2880
      Width           =   120
   End
   Begin VB.Label clblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6480
      TabIndex        =   6
      Top             =   2280
      Width           =   120
   End
   Begin VB.Label clblIDCard 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2265
      TabIndex        =   5
      Top             =   2340
      Width           =   120
   End
   Begin VB.Label clblPhotoNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7320
      TabIndex        =   4
      Top             =   1755
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label clblSysNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4560
      TabIndex        =   3
      Top             =   1755
      Width           =   120
   End
   Begin VB.Label clblDay 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3030
      TabIndex        =   2
      Top             =   1755
      Width           =   120
   End
   Begin VB.Label clblMonth 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2190
      TabIndex        =   1
      Top             =   1755
      Width           =   120
   End
   Begin VB.Label clblYear 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1305
      TabIndex        =   0
      Top             =   1755
      Width           =   120
   End
   Begin VB.Image cimgPhoto 
      Height          =   1500
      Left            =   8160
      Top             =   240
      Width           =   1185
   End
End
Attribute VB_Name = "frm打印体检登记表"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public pobj文书内容 As Object 'recordset[系统编号，姓名，身份证号，性别，单位名称]

Private Sub Form_Load()
    Dim lstrTmp As String
    
    On Error GoTo errhandler
    
    '填文书内容。
    clblYear.Caption = Left(Format(pobj文书内容("体检日期"), "yyyy-mm-dd"), 4)
    clblMonth.Caption = Format(Month(pobj文书内容("体检日期")), "00")
    clblDay.Caption = Format(Day(pobj文书内容("体检日期")), "00")
    
    clblSysNo.Caption = pobj文书内容("系统编号")
    clblPhotoNo.Caption = pobj文书内容("健康档案编号")
    
    clblIDCard.Caption = IIf(IsNull(pobj文书内容("公民身份号码")), "", pobj文书内容("公民身份号码"))
    
    clblName.Caption = IIf(IsNull(pobj文书内容("姓名")), "", pobj文书内容("姓名"))
    
    clblUnit.Caption = IIf(IsNull(pobj文书内容("单位名称")), "", pobj文书内容("单位名称"))
    
    
    clbl民族 = IIf(IsNull(pobj文书内容("民族")), "", pobj文书内容("民族"))
    
    lstrTmp = IIf(IsNull(pobj文书内容("卫生种类")), "", pobj文书内容("卫生种类"))
    If lstrTmp <> "" Then
        If Right(lstrTmp, 2) = "卫生" Then
            lstrTmp = Left(lstrTmp, Len(lstrTmp) - 2)
        End If
    End If
    clbl从业类别.Caption = lstrTmp
    
    cbccMain.Value = pobj文书内容("系统编号")
    
    '创建体检对象，获取照片。
    Dim lobj体检 As Object
    Set lobj体检 = CreateObject("体检对象.clsMedicalExam")
    lobj体检.系统编号 = pobj文书内容("系统编号")
    
    '显示像片。
    If Not lobj体检.体检人员.像片 Is Nothing Then
        cimgPhoto.Picture = lobj体检.体检人员.像片
    End If
    
    Exit Sub
errhandler:
    sfsub错误处理 "职业病文书管理", "frm打印体检登记表", "Form_Load", Err.Number, Err.Description, False
    Exit Sub
    Resume
End Sub


