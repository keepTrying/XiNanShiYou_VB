VERSION 5.00
Object = "{D9347025-9612-11D1-9D75-00C04FCC8CDC}#1.0#0"; "MSBCODE9.OCX"
Begin VB.Form frm打印体检表 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   19425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11415
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   20903.3
   ScaleMode       =   0  'User
   ScaleWidth      =   11415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Label clbl从业类别2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   15
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label clblAge2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   14
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label clblsex2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   13
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label clblName2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   12
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label clbl从业类别 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   11
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label clblName1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   10
      Top             =   13680
      Width           =   1095
   End
   Begin VB.Label clblUnit 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   9
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label clblAge 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   8
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label clblSex 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label clblName 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label clblDate1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
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
      Left            =   2880
      TabIndex        =   5
      Top             =   14280
      Width           =   720
   End
   Begin VB.Label clblDate 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
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
      Left            =   2400
      TabIndex        =   4
      Top             =   1680
      Width           =   120
   End
   Begin VB.Label clbl体检表号 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   0
      Left            =   1920
      TabIndex        =   3
      Top             =   1200
      Width           =   1035
   End
   Begin VB.Label clbl体检表号 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   1
      Left            =   9600
      TabIndex        =   2
      Top             =   14160
      Width           =   795
   End
   Begin VB.Label clbl试管号 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   1560
      TabIndex        =   1
      Top             =   1440
      Visible         =   0   'False
      Width           =   90
   End
   Begin BARCODELibCtl.BarCodeCtrl cbccMain 
      Height          =   765
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
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
   Begin VB.Image cimgPhoto 
      Height          =   1845
      Left            =   9360
      Top             =   960
      Width           =   1410
   End
End
Attribute VB_Name = "frm打印体检表"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public pobj文书内容 As Object 'recordset[系统编号，姓名，身份证号，性别，单位名称]

Private Sub Form_Load()
    Dim lstrTmp As String
    
    On Error GoTo errhandler
    
    
    lstrTmp = IIf(IsNull(pobj文书内容("卫生种类")), "", pobj文书内容("卫生种类"))
    If lstrTmp <> "" Then
        If Right(lstrTmp, 2) = "卫生" Then
            lstrTmp = Left(lstrTmp, Len(lstrTmp) - 2)
        End If
    End If
    Select Case lstrTmp
      Case "食品", 1
        clbl从业类别2.Caption = "食品"
        clblName2.Caption = IIf(IsNull(pobj文书内容("姓名")), "", pobj文书内容("姓名"))
        clblsex2.Caption = IIf(IsNull(pobj文书内容("性别")), "", pobj文书内容("性别"))
        clblAge2.Caption = IIf(IsNull(pobj文书内容("出生日期")), "", DateDiff("yyyy", pobj文书内容("出生日期"), Date))
        
        clblName.Visible = False
        clblSex.Visible = False
        clblUnit.Visible = False
        clblAge.Visible = False
        clbl从业类别.Visible = False
      Case "公共", 2
        clbl从业类别.Caption = "公共场所"
        clblName.Caption = IIf(IsNull(pobj文书内容("姓名")), "", pobj文书内容("姓名"))
        clblSex.Caption = IIf(IsNull(pobj文书内容("性别")), "", pobj文书内容("性别"))
        clblUnit.Caption = IIf(IsNull(pobj文书内容("单位名称")), "", pobj文书内容("单位名称"))
        clblAge.Caption = IIf(IsNull(pobj文书内容("出生日期")), "", DateDiff("yyyy", pobj文书内容("出生日期"), Date))
        
        clblName2.Visible = False
        clblsex2.Visible = False
        clbl从业类别2.Visible = False
        clblAge2.Visible = False
      Case "药品", 4
        clbl从业类别.Caption = "药品"
        clblName.Caption = IIf(IsNull(pobj文书内容("姓名")), "", pobj文书内容("姓名"))
        clblSex.Caption = IIf(IsNull(pobj文书内容("性别")), "", pobj文书内容("性别"))
        clblUnit.Caption = IIf(IsNull(pobj文书内容("单位名称")), "", pobj文书内容("单位名称"))
        clblAge.Caption = IIf(IsNull(pobj文书内容("出生日期")), "", DateDiff("yyyy", pobj文书内容("出生日期"), Date))
        
        clblName.Visible = False
        clblSex.Visible = False
        clblUnit.Visible = False
        clblAge.Visible = False
        clbl从业类别2.Visible = False
      Case Else
        clbl从业类别.Caption = lstrTmp
    End Select
    
    
    
    
    '填文书内容。
    cbccMain.Value = pobj文书内容("系统编号")
    
    '创建体检对象，获取照片。
    Dim lobj体检 As Object
    Set lobj体检 = CreateObject("体检对象.clsMedicalExam")
    lobj体检.系统编号 = pobj文书内容("系统编号")
    
      
    '显示像片。
    If Not lobj体检.体检人员.像片 Is Nothing Then
        cimgPhoto.Picture = lobj体检.体检人员.像片
    End If
    
    On Error Resume Next
    Dim lobjRec As Object
    Set lobjRec = dafuncGetData("select isnull(项目值,'') from 体检管理_体检附加信息表 where 系统编号='" & pobj文书内容("系统编号") & "' and 附加项目='血号'")
    If lobjRec.recordcount > 0 Then
        clbl体检表号(0).Caption = lobjRec(0)
        clbl体检表号(1).Caption = lobjRec(0)
    End If
    
    clblDate = Format(Date, "yyyy-mm-dd")
    clblDate1 = clblDate
    
    clblName1.Caption = IIf(IsNull(pobj文书内容("姓名")), "", pobj文书内容("姓名"))
    
    
    
    Exit Sub
errhandler:
    sfsub错误处理 "体检文书管理", "frm打印体检登记表", "Form_Load", Err.Number, Err.Description, False
    Exit Sub
    Resume
End Sub


