VERSION 5.00
Object = "{D9347025-9612-11D1-9D75-00C04FCC8CDC}#1.0#0"; "MSBCODE9.OCX"
Begin VB.Form frm打印体检单 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000009&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   14625
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   14625
   ScaleWidth      =   9870
   ShowInTaskbar   =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame cfram体检单 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6120
      Left            =   300
      TabIndex        =   28
      Top             =   570
      Width           =   9345
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "体 检 号："
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
         Index           =   2
         Left            =   90
         TabIndex        =   48
         Top             =   915
         Width           =   1050
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓    名："
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
         Index           =   2
         Left            =   90
         TabIndex        =   47
         Top             =   1267
         Width           =   1050
      End
      Begin VB.Label label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "身份证号："
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
         Index           =   2
         Left            =   90
         TabIndex        =   46
         Top             =   1971
         Width           =   1050
      End
      Begin VB.Label clblSysNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Index           =   0
         Left            =   1425
         TabIndex        =   45
         Top             =   915
         Width           =   105
      End
      Begin VB.Label clblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Index           =   0
         Left            =   1425
         TabIndex        =   44
         Top             =   1267
         Width           =   105
      End
      Begin VB.Label clblIDCard 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   1425
         TabIndex        =   43
         Top             =   1971
         Width           =   105
      End
      Begin VB.Image cimgPhoto 
         Height          =   1845
         Left            =   7080
         Stretch         =   -1  'True
         Top             =   1005
         Width           =   1410
      End
      Begin VB.Label clblUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "苏州工业园区体检中心"
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
         Index           =   1
         Left            =   6120
         TabIndex        =   42
         Top             =   5640
         Width           =   2100
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "谢谢合作！"
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
         Index           =   1
         Left            =   3360
         TabIndex        =   41
         Top             =   5250
         Width           =   1050
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "5.如有代检等作弊行为，取消体检结果，并通报单位及人力资源公司。"
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
         Index           =   1
         Left            =   90
         TabIndex        =   40
         Top             =   4440
         Width           =   6510
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "4.体检必须当日完成，隔日无效。"
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
         Index           =   1
         Left            =   90
         TabIndex        =   39
         Top             =   4110
         Width           =   3150
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3.检测血压偏高者，可休息一刻钟，多测几次，选其中低值记入。但必须以体检当日血压为准"
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
         Index           =   1
         Left            =   90
         TabIndex        =   38
         Top             =   3810
         Width           =   8610
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2.尿检验只需把装有尿液的尿杯放到指定的地方，化验单压在下面即可"
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
         Index           =   1
         Left            =   90
         TabIndex        =   37
         Top             =   3525
         Width           =   6510
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         Index           =   1
         X1              =   105
         X2              =   9000
         Y1              =   3435
         Y2              =   3420
      End
      Begin VB.Label clbl科室 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "检验室   放射科   外科   心电图   眼科   内科"
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
         Left            =   90
         TabIndex        =   36
         Top             =   3150
         Width           =   8685
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   1
         X1              =   105
         X2              =   9000
         Y1              =   3045
         Y2              =   3045
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1.办理完手续后凭本体检单到以下个科室进行体检"
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
         Index           =   1
         Left            =   90
         TabIndex        =   35
         Top             =   2775
         Width           =   4620
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "体检单"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3270
         TabIndex        =   34
         Top             =   195
         Width           =   1365
      End
      Begin BARCODELibCtl.BarCodeCtrl cbccMain 
         Height          =   765
         Index           =   0
         Left            =   6600
         TabIndex        =   33
         Top             =   120
         Width           =   3015
         Style           =   7
         SubStyle        =   -1
         Validation      =   0
         LineWeight      =   3
         Direction       =   0
         ShowData        =   1
         Value           =   "123456 Code-128"
         ForeColor       =   0
         BackColor       =   16777215
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单位名称："
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
         Left            =   90
         TabIndex        =   32
         Top             =   1619
         Width           =   1050
      End
      Begin VB.Label clbl单位名称 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
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
         Left            =   1425
         TabIndex        =   31
         Top             =   1619
         Width           =   105
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "收费编号："
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
         Left            =   90
         TabIndex        =   30
         Top             =   2325
         Width           =   1050
      End
      Begin VB.Label clbl收费批号 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
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
         Left            =   1425
         TabIndex        =   29
         Top             =   2325
         Width           =   105
      End
   End
   Begin VB.Frame cfram血液规检验单 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   1725
      Left            =   300
      TabIndex        =   18
      Top             =   11880
      Width           =   9465
      Begin VB.Line Line3 
         BorderWidth     =   2
         Index           =   5
         X1              =   75
         X2              =   9360
         Y1              =   30
         Y2              =   30
      End
      Begin BARCODELibCtl.BarCodeCtrl cbccMain 
         Height          =   765
         Index           =   3
         Left            =   6720
         TabIndex        =   26
         Top             =   120
         Width           =   3015
         Style           =   7
         SubStyle        =   -1
         Validation      =   0
         LineWeight      =   3
         Direction       =   0
         ShowData        =   1
         Value           =   "123456 Code-128"
         ForeColor       =   0
         BackColor       =   16777215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "体 检 号："
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
         Index           =   5
         Left            =   90
         TabIndex        =   25
         Top             =   870
         Width           =   1050
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓    名："
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
         Index           =   5
         Left            =   90
         TabIndex        =   24
         Top             =   1125
         Width           =   1050
      End
      Begin VB.Label label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性    别："
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
         Index           =   5
         Left            =   90
         TabIndex        =   23
         Top             =   1395
         Width           =   1050
      End
      Begin VB.Label clblSysNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Index           =   3
         Left            =   1200
         TabIndex        =   22
         Top             =   870
         Width           =   105
      End
      Begin VB.Label clblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Index           =   3
         Left            =   1200
         TabIndex        =   21
         Top             =   1125
         Width           =   105
      End
      Begin VB.Label clblSex 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Index           =   2
         Left            =   1200
         TabIndex        =   20
         Top             =   1395
         Width           =   105
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "血液检验单"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2715
         TabIndex        =   19
         Top             =   360
         Width           =   1650
      End
   End
   Begin VB.Frame cfram大便培养规检验单 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   1740
      Left            =   240
      TabIndex        =   9
      Top             =   9480
      Width           =   9465
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "大便培养检验单"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2715
         TabIndex        =   17
         Top             =   360
         Width           =   2310
      End
      Begin VB.Label clblSex 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Index           =   1
         Left            =   1200
         TabIndex        =   16
         Top             =   1395
         Width           =   105
      End
      Begin VB.Label clblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Index           =   2
         Left            =   1200
         TabIndex        =   15
         Top             =   1125
         Width           =   105
      End
      Begin VB.Label clblSysNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Index           =   2
         Left            =   1200
         TabIndex        =   14
         Top             =   810
         Width           =   105
      End
      Begin VB.Label label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性    别："
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
         Index           =   4
         Left            =   90
         TabIndex        =   13
         Top             =   1395
         Width           =   1050
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓    名："
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
         Index           =   4
         Left            =   90
         TabIndex        =   12
         Top             =   1125
         Width           =   1050
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "体 检 号："
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
         Index           =   4
         Left            =   90
         TabIndex        =   11
         Top             =   870
         Width           =   1050
      End
      Begin BARCODELibCtl.BarCodeCtrl cbccMain 
         Height          =   765
         Index           =   2
         Left            =   6720
         TabIndex        =   10
         Top             =   120
         Width           =   3015
         Style           =   7
         SubStyle        =   -1
         Validation      =   0
         LineWeight      =   3
         Direction       =   0
         ShowData        =   1
         Value           =   "123456 Code-128"
         ForeColor       =   0
         BackColor       =   16777215
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         Index           =   4
         X1              =   90
         X2              =   9240
         Y1              =   30
         Y2              =   30
      End
   End
   Begin VB.Frame cfram尿常规检验单 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   1965
      Left            =   300
      TabIndex        =   0
      Top             =   6960
      Width           =   9345
      Begin VB.Label clblSex 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Index           =   0
         Left            =   1200
         TabIndex        =   27
         Top             =   1395
         Width           =   105
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         Index           =   1
         X1              =   90
         X2              =   9240
         Y1              =   30
         Y2              =   30
      End
      Begin BARCODELibCtl.BarCodeCtrl cbccMain 
         Height          =   765
         Index           =   1
         Left            =   6600
         TabIndex        =   8
         Top             =   120
         Width           =   3015
         Style           =   7
         SubStyle        =   -1
         Validation      =   0
         LineWeight      =   3
         Direction       =   0
         ShowData        =   1
         Value           =   "123456 Code-128"
         ForeColor       =   0
         BackColor       =   16777215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "体 检 号："
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
         Index           =   1
         Left            =   90
         TabIndex        =   7
         Top             =   870
         Width           =   1050
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓    名："
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
         Index           =   1
         Left            =   90
         TabIndex        =   6
         Top             =   1125
         Width           =   1050
      End
      Begin VB.Label label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性    别："
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
         Index           =   1
         Left            =   90
         TabIndex        =   5
         Top             =   1395
         Width           =   1050
      End
      Begin VB.Label clblSysNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Index           =   1
         Left            =   1200
         TabIndex        =   4
         Top             =   810
         Width           =   105
      End
      Begin VB.Label clblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Index           =   1
         Left            =   1200
         TabIndex        =   3
         Top             =   1125
         Width           =   105
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "尿常规检验单"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2715
         TabIndex        =   2
         Top             =   360
         Width           =   1995
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "注：本检验单用于尿常规检验，压在装有尿液的尿杯下，放在指定地方即可"
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
         Index           =   0
         Left            =   90
         TabIndex        =   1
         Top             =   1665
         Width           =   6930
      End
   End
End
Attribute VB_Name = "frm打印体检单"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'作者：邓恒。

Public pobj文书内容 As Object 'recordset[系统编号，姓名，身份证号，性别，单位名称]

Public pbln是否有尿常规 As Boolean
Public pbln是否有大便培养 As Boolean
Public pbln是否有验血 As Boolean

'功能：填文书内容。
'作者：邓恒。
Private Sub Form_Load()
    On Error GoTo errHandler
    Dim i As Integer
    Label1.Left = (Me.Width - Label1.Width) / 2
    Label5.Left = (Me.Width - Label5.Width) / 2
    Label7.Left = (Me.Width - Label7.Width) / 2
    Label17.Left = (Me.Width - Label17.Width) / 2
    
    '填文书内容。
    For i = 0 To clblSysNo.Count - 1
        clblSysNo(i).Caption = pobj文书内容("系统编号")
    Next i
    For i = 0 To clblName.Count - 1
        clblName(i).Caption = pobj文书内容("姓名")
    Next i
    clblIDCard.Caption = pobj文书内容("公民身份号码")
    For i = 0 To clblSex.Count - 1
        clblSex(i).Caption = pobj文书内容("性别")
    Next i

    For i = 0 To cbccMain.Count - 1
        cbccMain(i).Value = pobj文书内容("系统编号")
    Next i
    clbl收费批号.Caption = pobj文书内容("收费批号")
    clbl单位名称.Caption = pobj文书内容("单位名称")
    
    '防疫站名。
    clblUnit(1).Caption = um防疫站名
    
    '检查需要做得检查。
    If Not pbln是否有尿常规 Then
        cfram尿常规检验单.Visible = False
    End If
    If Not pbln是否有大便培养 Then
        cfram大便培养规检验单.Visible = False
    End If
    If Not pbln是否有验血 Then
        cfram血液规检验单.Visible = False
    End If
    
    '创建体检对象，获取照片。
    Dim lobj体检 As Object
    Set lobj体检 = CreateObject("体检对象.clsMedicalExam")
    lobj体检.系统编号 = pobj文书内容("系统编号")
    
    '显示像片。
    cimgPhoto.Picture = lobj体检.体检人员.像片
    
    '修改：2002-8-16（杨春）动态显示科室。
    On Error Resume Next
    clbl科室.Caption = pobj文书内容("体检科室名串")
    
    Exit Sub
errHandler:
    sfsub错误处理 "体检文书管理", "frm打印体检单3", "Form_Load", Err.Number, Err.Description, False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    
End Sub
