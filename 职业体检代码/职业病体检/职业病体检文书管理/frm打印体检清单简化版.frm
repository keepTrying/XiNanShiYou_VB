VERSION 5.00
Object = "{D9347025-9612-11D1-9D75-00C04FCC8CDC}#1.0#0"; "MSBCODE9.OCX"
Begin VB.Form frm打印体检清单简化版 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "职业病体检清单"
   ClientHeight    =   13515
   ClientLeft      =   -225
   ClientTop       =   -1965
   ClientWidth     =   12540
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   15867.48
   ScaleMode       =   0  'User
   ScaleWidth      =   11882.77
   StartUpPosition =   3  '窗口缺省
   Begin VB.CheckBox DeptName 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DeptName20"
      Height          =   255
      Index           =   20
      Left            =   240
      TabIndex        =   72
      Top             =   1320
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CheckBox DeptName 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DeptName15"
      Height          =   255
      Index           =   15
      Left            =   10560
      TabIndex        =   71
      Top             =   10440
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CheckBox DeptName 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DeptName14"
      Height          =   255
      Index           =   14
      Left            =   7920
      TabIndex        =   70
      Top             =   11280
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CheckBox DeptName 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DeptName19"
      Height          =   255
      Index           =   19
      Left            =   10440
      TabIndex        =   58
      Top             =   10800
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CheckBox DeptName 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DeptName18"
      Height          =   255
      Index           =   18
      Left            =   10440
      TabIndex        =   57
      Top             =   9600
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CheckBox DeptName 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DeptName17"
      Height          =   255
      Index           =   17
      Left            =   6720
      TabIndex        =   56
      Top             =   11280
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CheckBox DeptName 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DeptName16"
      Height          =   255
      Index           =   16
      Left            =   5880
      TabIndex        =   55
      Top             =   10320
      Width           =   3135
   End
   Begin VB.CheckBox DeptName 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DeptName13"
      Height          =   255
      Index           =   13
      Left            =   9960
      TabIndex        =   54
      Top             =   11280
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CheckBox DeptName 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DeptName12"
      Height          =   255
      Index           =   12
      Left            =   1200
      TabIndex        =   53
      Top             =   10320
      Width           =   3135
   End
   Begin VB.CheckBox DeptName 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DeptName11"
      Height          =   255
      Index           =   11
      Left            =   5880
      TabIndex        =   52
      Top             =   7920
      Width           =   3135
   End
   Begin VB.CheckBox DeptName 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DeptName10"
      Height          =   255
      Index           =   10
      Left            =   1200
      TabIndex        =   51
      Top             =   9840
      Width           =   3135
   End
   Begin VB.CheckBox DeptName 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DeptName9"
      Height          =   255
      Index           =   9
      Left            =   1200
      TabIndex        =   50
      Top             =   8880
      Width           =   3135
   End
   Begin VB.CheckBox DeptName 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DeptName8"
      Height          =   255
      Index           =   8
      Left            =   1200
      TabIndex        =   49
      Top             =   9360
      Width           =   3135
   End
   Begin VB.CheckBox DeptName 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DeptName7"
      Height          =   255
      Index           =   7
      Left            =   5880
      TabIndex        =   48
      Top             =   7440
      Width           =   3135
   End
   Begin VB.CheckBox DeptName 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DeptName6"
      Height          =   255
      Index           =   6
      Left            =   5880
      TabIndex        =   47
      Top             =   9840
      Width           =   3135
   End
   Begin VB.CheckBox DeptName 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DeptName5"
      Height          =   255
      Index           =   5
      Left            =   5880
      TabIndex        =   46
      Top             =   9360
      Width           =   3135
   End
   Begin VB.CheckBox DeptName 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DeptName4"
      Height          =   255
      Index           =   4
      Left            =   5880
      TabIndex        =   45
      Top             =   8400
      Width           =   3135
   End
   Begin VB.CheckBox DeptName 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DeptName3"
      Height          =   255
      Index           =   3
      Left            =   5880
      TabIndex        =   44
      Top             =   8880
      Width           =   3135
   End
   Begin VB.CheckBox DeptName 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DeptName2"
      Height          =   255
      Index           =   2
      Left            =   1200
      TabIndex        =   43
      Top             =   8400
      Width           =   3135
   End
   Begin VB.CheckBox DeptName 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DeptName1"
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   42
      Top             =   7920
      Width           =   3135
   End
   Begin VB.CheckBox DeptName 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DeptName0"
      Height          =   255
      Index           =   0
      Left            =   1200
      TabIndex        =   41
      Top             =   7440
      Width           =   3135
   End
   Begin VB.CommandButton ccmdPrint 
      Caption         =   "打印"
      Height          =   495
      Left            =   480
      TabIndex        =   20
      Top             =   2640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton ccmdExit 
      Caption         =   "退出"
      Height          =   495
      Left            =   2160
      TabIndex        =   19
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   1935
      Left            =   8880
      ScaleHeight     =   1875
      ScaleWidth      =   1515
      TabIndex        =   18
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label32 
      Caption         =   "  年   月    日"
      Height          =   375
      Left            =   1680
      TabIndex        =   76
      Top             =   12960
      Width           =   2175
   End
   Begin VB.Label Label31 
      Caption         =   "受检人签名："
      Height          =   375
      Left            =   1080
      TabIndex        =   75
      Top             =   12360
      Width           =   1815
   End
   Begin VB.Label clbChekData 
      BackColor       =   &H80000005&
      Height          =   255
      Left            =   8760
      TabIndex        =   74
      Top             =   12360
      Width           =   1935
   End
   Begin VB.Line Line24 
      X1              =   8187.171
      X2              =   10233.96
      Y1              =   14934.1
      Y2              =   14934.1
   End
   Begin VB.Label Label23 
      BackColor       =   &H00FFFFFF&
      Caption         =   "      2、体检结束后请将体检指引单交到领表处。"
      Height          =   375
      Index           =   2
      Left            =   1080
      TabIndex        =   73
      Top             =   11640
      Width           =   5535
   End
   Begin VB.Line Line10 
      Index           =   2
      X1              =   6026.668
      X2              =   6026.668
      Y1              =   12539.01
      Y2              =   13102.56
   End
   Begin VB.Line Line10 
      Index           =   1
      X1              =   3411.321
      X2              =   3411.321
      Y1              =   12539.01
      Y2              =   13102.56
   End
   Begin VB.Label Label30 
      BackColor       =   &H00FFFFFF&
      Caption         =   "次/分"
      Height          =   255
      Left            =   8160
      TabIndex        =   69
      Top             =   10800
      Width           =   495
   End
   Begin VB.Label Label28 
      BackColor       =   &H00FFFFFF&
      Caption         =   "mmHg"
      Height          =   255
      Left            =   5520
      TabIndex        =   68
      Top             =   10800
      Width           =   495
   End
   Begin VB.Label Label27 
      BackColor       =   &H00FFFFFF&
      Caption         =   "收缩压:"
      Height          =   255
      Left            =   1200
      TabIndex        =   67
      Top             =   10800
      Width           =   855
   End
   Begin VB.Label Label26 
      BackColor       =   &H00FFFFFF&
      Caption         =   "mmHg"
      Height          =   255
      Left            =   2880
      TabIndex        =   66
      Top             =   10800
      Width           =   495
   End
   Begin VB.Line Line8 
      Index           =   9
      X1              =   1023.396
      X2              =   9892.832
      Y1              =   13102.56
      Y2              =   13102.56
   End
   Begin VB.Label Label23 
      BackColor       =   &H00FFFFFF&
      Caption         =   "备注：1、进行受检者个人信息录入前请先测血压、心率。"
      Height          =   375
      Index           =   0
      Left            =   1080
      TabIndex        =   63
      Top             =   11280
      Width           =   5535
   End
   Begin VB.Label Label25 
      BackColor       =   &H00FFFFFF&
      Caption         =   "心率:"
      Height          =   255
      Left            =   6480
      TabIndex        =   65
      Top             =   10800
      Width           =   615
   End
   Begin VB.Label Label24 
      BackColor       =   &H00FFFFFF&
      Caption         =   "舒张压:"
      Height          =   255
      Left            =   3720
      TabIndex        =   64
      Top             =   10800
      Width           =   735
   End
   Begin VB.Label Label22 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "体检日期："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   7440
      TabIndex        =   62
      Top             =   12480
      Width           =   1200
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "注：打勾的为需体检的项目。"
      Height          =   180
      Left            =   1080
      TabIndex        =   61
      Top             =   6720
      Width           =   2340
   End
   Begin VB.Line Line23 
      X1              =   1933.082
      X2              =   9892.832
      Y1              =   7748.825
      Y2              =   7748.825
   End
   Begin VB.Label clblCompany 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2040
      TabIndex        =   60
      Top             =   6240
      Width           =   8415
   End
   Begin VB.Label Label20 
      BackColor       =   &H00FFFFFF&
      Caption         =   "单位名称："
      Height          =   255
      Left            =   1080
      TabIndex        =   59
      Top             =   6360
      Width           =   975
   End
   Begin VB.Line Line8 
      Index           =   8
      X1              =   1023.396
      X2              =   9892.832
      Y1              =   12539.01
      Y2              =   12539.01
   End
   Begin VB.Label Label29 
      BackColor       =   &H00FFFFFF&
      Caption         =   "医师签字"
      Height          =   225
      Index           =   3
      Left            =   9480
      TabIndex        =   40
      Top             =   7065
      Width           =   855
   End
   Begin VB.Label Label29 
      BackColor       =   &H00FFFFFF&
      Caption         =   "体检科室名称"
      Height          =   225
      Index           =   2
      Left            =   6360
      TabIndex        =   39
      Top             =   7065
      Width           =   1215
   End
   Begin VB.Label Label29 
      BackColor       =   &H00FFFFFF&
      Caption         =   "医师签字"
      Height          =   225
      Index           =   1
      Left            =   4680
      TabIndex        =   38
      Top             =   7080
      Width           =   855
   End
   Begin VB.Label Label29 
      BackColor       =   &H00FFFFFF&
      Caption         =   "体检科室名称"
      Height          =   225
      Index           =   0
      Left            =   1680
      TabIndex        =   37
      Top             =   7065
      Width           =   1215
   End
   Begin VB.Line Line7 
      X1              =   7050.064
      X2              =   8187.171
      Y1              =   6058.172
      Y2              =   6058.172
   End
   Begin VB.Line Line22 
      X1              =   3866.164
      X2              =   4775.85
      Y1              =   6058.172
      Y2              =   6058.172
   End
   Begin VB.Line Line21 
      X1              =   7618.618
      X2              =   9892.832
      Y1              =   7185.274
      Y2              =   7185.274
   End
   Begin VB.Line Line20 
      X1              =   6140.378
      X2              =   8300.882
      Y1              =   6621.723
      Y2              =   6621.723
   End
   Begin VB.Line Line19 
      X1              =   4321.007
      X2              =   6367.8
      Y1              =   7185.274
      Y2              =   7185.274
   End
   Begin VB.Line Line18 
      X1              =   2729.057
      X2              =   4093.586
      Y1              =   6621.723
      Y2              =   6621.723
   End
   Begin VB.Line Line17 
      X1              =   1933.082
      X2              =   3070.189
      Y1              =   7185.274
      Y2              =   7185.274
   End
   Begin VB.Label clblMedicalTable 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8040
      TabIndex        =   36
      Top             =   5880
      Width           =   2415
   End
   Begin VB.Label clblMedicalCategory 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4560
      TabIndex        =   35
      Top             =   5880
      Width           =   2175
   End
   Begin VB.Label clblMedicalType 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2040
      TabIndex        =   34
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label clblMarried 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4200
      TabIndex        =   33
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label clblHagardAge 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6480
      TabIndex        =   32
      Top             =   5400
      Width           =   2175
   End
   Begin VB.Label clblMoney 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6240
      TabIndex        =   31
      Top             =   3840
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label clblChargeNumber 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6240
      TabIndex        =   30
      Top             =   3360
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label clblHagard 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2880
      TabIndex        =   29
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Line Line13 
      X1              =   9892.832
      X2              =   9892.832
      Y1              =   8171.488
      Y2              =   13102.56
   End
   Begin VB.Line Line12 
      X1              =   1023.396
      X2              =   1023.396
      Y1              =   8171.488
      Y2              =   13102.56
   End
   Begin VB.Line Line8 
      Index           =   0
      X1              =   1023.396
      X2              =   9892.832
      Y1              =   8171.488
      Y2              =   8171.488
   End
   Begin VB.Line Line8 
      Index           =   1
      X1              =   1023.396
      X2              =   9892.832
      Y1              =   8594.151
      Y2              =   8594.151
   End
   Begin VB.Line Line8 
      Index           =   2
      X1              =   1023.396
      X2              =   9892.832
      Y1              =   9157.702
      Y2              =   9157.702
   End
   Begin VB.Line Line8 
      Index           =   3
      X1              =   1023.396
      X2              =   9892.832
      Y1              =   9721.253
      Y2              =   9721.253
   End
   Begin VB.Line Line8 
      Index           =   4
      X1              =   1023.396
      X2              =   9892.832
      Y1              =   10284.8
      Y2              =   10284.8
   End
   Begin VB.Line Line8 
      Index           =   5
      X1              =   1023.396
      X2              =   9892.832
      Y1              =   10848.36
      Y2              =   10848.36
   End
   Begin VB.Line Line9 
      X1              =   5458.114
      X2              =   5458.114
      Y1              =   8171.488
      Y2              =   12539.01
   End
   Begin VB.Line Line8 
      Index           =   6
      X1              =   1023.396
      X2              =   9892.832
      Y1              =   11411.91
      Y2              =   11411.91
   End
   Begin VB.Line Line8 
      Index           =   7
      X1              =   1023.396
      X2              =   9892.832
      Y1              =   11975.46
      Y2              =   11975.46
   End
   Begin VB.Line Line10 
      Index           =   0
      X1              =   4207.296
      X2              =   4207.296
      Y1              =   8171.488
      Y2              =   12539.01
   End
   Begin VB.Line Line11 
      X1              =   8642.015
      X2              =   8642.015
      Y1              =   8171.488
      Y2              =   13102.56
   End
   Begin VB.Label Label19 
      BackColor       =   &H00FFFFFF&
      Caption         =   "婚否："
      Height          =   255
      Left            =   3600
      TabIndex        =   28
      Top             =   4920
      Width           =   615
   End
   Begin VB.Label Label18 
      BackColor       =   &H00FFFFFF&
      Caption         =   "职业危害工龄："
      Height          =   255
      Left            =   5280
      TabIndex        =   27
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFFFF&
      Caption         =   "危害因素或特殊作业："
      Height          =   255
      Left            =   1080
      TabIndex        =   26
      Top             =   5400
      Width           =   2175
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFFFF&
      Caption         =   "收费金额："
      Height          =   255
      Left            =   5280
      TabIndex        =   25
      Top             =   3840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      Caption         =   "收费批号："
      Height          =   255
      Left            =   5280
      TabIndex        =   24
      Top             =   3360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "检查种类："
      Height          =   255
      Left            =   3600
      TabIndex        =   23
      Top             =   5880
      Width           =   975
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "检查类型："
      Height          =   255
      Left            =   1080
      TabIndex        =   22
      Top             =   5880
      Width           =   975
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "检查表名称："
      Height          =   255
      Left            =   6960
      TabIndex        =   21
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Line Line6 
      X1              =   5458.114
      X2              =   6367.8
      Y1              =   6058.172
      Y2              =   6058.172
   End
   Begin VB.Line Line5 
      X1              =   1933.082
      X2              =   3070.189
      Y1              =   6058.172
      Y2              =   6058.172
   End
   Begin VB.Line Line4 
      X1              =   7050.064
      X2              =   8187.171
      Y1              =   5494.622
      Y2              =   5494.622
   End
   Begin VB.Line Line3 
      X1              =   5458.114
      X2              =   6367.8
      Y1              =   5494.622
      Y2              =   5494.622
   End
   Begin VB.Line Line2 
      X1              =   3866.164
      X2              =   4775.85
      Y1              =   5494.622
      Y2              =   5494.622
   End
   Begin VB.Line Line1 
      X1              =   1478.239
      X2              =   3070.189
      Y1              =   5494.622
      Y2              =   5494.622
   End
   Begin VB.Label clblWorkAge 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7440
      TabIndex        =   17
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFFFFF&
      Caption         =   "工龄："
      Height          =   255
      Left            =   6960
      TabIndex        =   16
      Top             =   4920
      Width           =   735
   End
   Begin VB.Label clblProfession 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5880
      TabIndex        =   15
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFFFF&
      Caption         =   "工种："
      Height          =   255
      Left            =   5280
      TabIndex        =   14
      Top             =   4920
      Width           =   615
   End
   Begin VB.Label clblDegree 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2040
      TabIndex        =   13
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFFF&
      Caption         =   "文化程度："
      Height          =   255
      Left            =   1080
      TabIndex        =   12
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label clblNationality 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7560
      TabIndex        =   11
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "民族："
      Height          =   255
      Left            =   6960
      TabIndex        =   10
      Top             =   4440
      Width           =   615
   End
   Begin VB.Label clblAge 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5880
      TabIndex        =   9
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "年龄："
      Height          =   255
      Left            =   5280
      TabIndex        =   8
      Top             =   4440
      Width           =   615
   End
   Begin VB.Label clblSex 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4200
      TabIndex        =   7
      Top             =   4440
      Width           =   735
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "性别："
      Height          =   255
      Left            =   3600
      TabIndex        =   6
      Top             =   4440
      Width           =   615
   End
   Begin VB.Label clblName 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "姓名："
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   4440
      Width           =   615
   End
   Begin BARCODELibCtl.BarCodeCtrl BarCode 
      Height          =   735
      Left            =   1920
      TabIndex        =   3
      Top             =   3360
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
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "体检编号："
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "职业健康检查指引单"
      Height          =   495
      Left            =   3240
      TabIndex        =   1
      Top             =   1560
      Width           =   4575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "西南石油大学校医院"
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   1080
      Width           =   7575
   End
End
Attribute VB_Name = "frm打印体检清单简化版"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''
'2012-06-18 于登淼
'添加整个窗体，用来实现单个体检人员体检清单赋值与初始化。
'预览与打印功能在其它窗口实现。
'''''''''''''''

Option Explicit

Public sysNo As String
Private paraDept(0 To 20) As String '记录每个科室体检项目
Private numInOneLine As Integer     '每行填写多少个项目(不包括医师签名)

Private Sub ccmdExit_Click()
    Unload Me
End Sub

Private Sub ccmdPrint_Click()
    Me.ccmdPrint.Visible = False
    Me.ccmdExit.Visible = False
   ' Form_Load
'    Me.PrintForm
    Me.ccmdPrint.Visible = True
    Me.ccmdExit.Visible = True
End Sub

Private Sub Form_Load()
 Dim devPrinter As Printer
 Dim bqprt As String
 Dim qdprt As String
' frmSelect.Show 1
'subSelect
Dim c As String
c = MsgBox("选择“是”打印西南石油大学指引单，选择“否”打印南充市疾控马市铺门诊指引单，选择“取消”打印中原石油指引单", vbYesNoCancel, "提示")
If c = vbYes Then
    Label1.Caption = "西南石油大学（南充）校医院"
    Label2.Caption = "健康检查指引单"
ElseIf c = vbNo Then
    Label1.Caption = "南充市疾病预防控制中心马市铺门诊"
Else
    Label1.Caption = "中原油田疾病预防控制中心"
End If
 bqprt = ""
 qdprt = ""
     For Each devPrinter In Printers
'        Me.Show    '测试   牟俊
        If devPrinter.DeviceName = "清单打印机" Then
          Set Printer = devPrinter
          qdprt = "清单打印机"
          subFontInit
        subPersonalInfoInit
    subDeptInit
    subDeptFill
'    Me.Show     '测试显示打印窗口 2015-12-28 by 牟俊
    Me.PrintForm
 
       Exit For
       
        End If
     Next
   If qdprt = "清单打印机" Then
   Else
        Dim tips
         tips = MsgBox("没有设置清单打印机！", vbOKOnly + vbCritical, "提示")
      Exit Sub

   End If
   

End Sub
Sub subSelect()
    frmSelect.Show 1
End Sub

'2012-06-18 于登淼
'控制标题字体大小。label1与label2右侧对齐时，字体刚好也对齐
Sub subFontInit()
    Label1.FontSize = 20
    Label2.FontSize = 20
End Sub

Sub subPersonalInfoInit()
    Dim lobjRec As Object
    Dim strSQL As String
    Dim chargeNumber As String
    
    If sysNo = "" Then sysNo = "0011201200004554" '''''''''''''测试行''''''''''''
    BarCode.Value = sysNo
    
    '载入文字类基本信息
    strSQL = "select * from 职业病体检_体检人员基本信息表 where 系统编号='" & sysNo & "'"
    dasubSetQueryTimeout 6000
    Set lobjRec = dafuncGetData(strSQL)
    clblName.Caption = lobjRec("姓名")
    clblSex.Caption = lobjRec("性别")
    clblAge.Caption = lobjRec("年龄")
    clblNationality.Caption = lobjRec("民族") '& "族"
    clblDegree.Caption = lobjRec("文化程度")
    clblProfession.Caption = lobjRec("现工种")
    clblWorkAge.Caption = lobjRec("工龄")
    clblMarried.Caption = lobjRec("婚否")
'    clblWorkClassify.Caption = lobjRec("职业分类")
    clblHagard.Caption = lobjRec("危害因素")
    clblHagardAge.Caption = lobjRec("职业危害工龄")
    clblCompany.Caption = lobjRec("单位名称")
    strSQL = "select * from 职业病体检_体检基本信息表 where 系统编号='" & sysNo & "'"
    dasubSetQueryTimeout 6000
    Set lobjRec = dafuncGetData(strSQL)
    clblMedicalType.Caption = lobjRec("体检类型")
    clblMedicalCategory.Caption = lobjRec("体检类别")
    clblMedicalTable.Caption = Mid(lobjRec("体检表编号"), 1, 4) & "体检表"
    clblChargeNumber.Caption = lobjRec("收费批号")
    chargeNumber = lobjRec("收费批号")
    '增加体检日期 add by lanchao 2015.9.6
    clbChekData.Caption = lobjRec("体检日期")
    
    
    '20150325
    'strSQL = "select * from 收费管理_费用信息表 where 收费批号='" & chargeNumber & "'"
    'dasubSetQueryTimeout 6000
    'Set lobjRec = dafuncGetData(strSQL)
    'If Not (lobjRec.bof Or lobjRec.EOF) Then
       ' clblMoney.Caption = lobjRec("金额") & " 元"
    'End If
    
    '载入体检照片（现场拍照，不是身份证照片）
    Set lobjRec = CreateObject("职业病对象.clspersonexamed")
    lobjRec.系统编号 = sysNo
    Picture1.Picture = lobjRec.像片
    
    Set lobjRec = Nothing
End Sub

Sub subDeptInit()
    Dim lobjRec As Object
    Dim strSQL As String
    Dim i As Integer
    Dim pCnt As Integer
    Dim preDept, curDept As Integer
    
    For i = LBound(paraDept) To UBound(paraDept): paraDept(i) = "0": Next
    strSQL = "select 体检项目 from 职业病体检_体检结果视图 where 系统编号='" & sysNo & "'"
    dasubSetQueryTimeout 6000
    Set lobjRec = dafuncGetData(strSQL)
    lobjRec.movefirst
    For i = 0 To lobjRec.RecordCount - 1
        If i = 0 Then
            preDept = CInt(Left(lobjRec("体检项目"), 2))
            pCnt = 0
        End If
        curDept = CInt(Left(lobjRec("体检项目"), 2))
        If preDept <> curDept Then
            paraDept(preDept) = pCnt
            pCnt = 0
            preDept = curDept
        End If
        pCnt = pCnt + 1
        lobjRec.movenext
        If lobjRec.EOF = True Then paraDept(curDept) = pCnt
    Next
End Sub

Sub subDeptFill()
    Dim i As Integer
    Dim strSQL As String
    Dim lobjRec As Object
    dasubSetQueryTimeout 6000
    strSQL = "select * from 系统管理_字典_字典内容表 where 描述='职业病体检_科室' and right(名称,1)='科' order by 编号"
    Set lobjRec = dafuncGetData(strSQL)
    
    lobjRec.movefirst
    For i = 0 To lobjRec.RecordCount - 1
        DeptName(CInt(lobjRec("编号")) - 1).Value = IIf(CInt(paraDept(CInt(lobjRec("编号")))) > 0, 1, 0)
        DeptName(CInt(lobjRec("编号")) - 1).Caption = lobjRec("名称")
        lobjRec.movenext
    Next
    For i = 0 To DeptName.Count - 1
        If DeptName.Item(i).Value = 0 Then
'            DeptName.Item(i).Caption = DeptName.Item(i).Caption & "（不需体检）"
            DeptName.Item(i).Visible = False
        End If
    Next i
    Set lobjRec = Nothing
End Sub

Public Function subPrint()
    ccmdPrint_Click
End Function


