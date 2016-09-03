VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmCareerHstRegt 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "受检者个人信息录入"
   ClientHeight    =   9315
   ClientLeft      =   4830
   ClientTop       =   2430
   ClientWidth     =   14055
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9315
   ScaleWidth      =   14055
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   5895
      Left            =   120
      TabIndex        =   151
      TabStop         =   0   'False
      Top             =   2880
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   10398
      _Version        =   393216
      Tabs            =   6
      Tab             =   4
      TabsPerRow      =   6
      TabHeight       =   520
      BackColor       =   0
      ForeColor       =   16711680
      TabCaption(0)   =   "个人生活史"
      TabPicture(0)   =   "frmCareerHstRegister.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "freRadiation"
      Tab(0).Control(1)=   "freOrdinary"
      Tab(0).Control(2)=   "freNuclear"
      Tab(0).Control(3)=   "Frame6"
      Tab(0).Control(4)=   "Frame1"
      Tab(0).Control(5)=   "ctxtOther"
      Tab(0).Control(6)=   "Label5(0)"
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "职业史"
      TabPicture(1)   =   "frmCareerHstRegister.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cgrd职业史"
      Tab(1).Control(1)=   "Frame10"
      Tab(1).Control(2)=   "Ccmd修改"
      Tab(1).Control(3)=   "Ccmd删除"
      Tab(1).Control(4)=   "Frame11"
      Tab(1).Control(5)=   "Command4"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "既往病史(包括职业病史)"
      TabPicture(2)   =   "frmCareerHstRegister.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame13"
      Tab(2).Control(1)=   "Cmmdmody病史"
      Tab(2).Control(2)=   "Ccmdcancel病史"
      Tab(2).Control(3)=   "cgrd病史"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "自觉症状"
      TabPicture(3)   =   "frmCareerHstRegister.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame15"
      Tab(3).Control(1)=   "cgrd症状"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "体格一般情况"
      TabPicture(4)   =   "frmCareerHstRegister.frx":0070
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Frame7"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Frame5"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Text临时"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "病状询问"
      TabPicture(5)   =   "frmCareerHstRegister.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "cgrdzzxw"
      Tab(5).Control(1)=   "Command1"
      Tab(5).Control(2)=   "Command2"
      Tab(5).ControlCount=   3
      Begin VB.TextBox Text临时 
         Height          =   1335
         Left            =   8400
         TabIndex        =   321
         Text            =   "用来存从体重秤读取的字符串"
         Top             =   1800
         Width           =   2775
      End
      Begin VB.Frame freRadiation 
         BorderStyle     =   0  'None
         Height          =   3735
         Left            =   -74760
         TabIndex        =   154
         Top             =   600
         Width           =   11895
         Begin VB.Frame Frame19 
            Caption         =   "烟酒史"
            ForeColor       =   &H000080FF&
            Height          =   2175
            Index           =   0
            Left            =   6000
            TabIndex        =   187
            Top             =   1560
            Width           =   5175
            Begin VB.TextBox ctxtMore 
               Height          =   375
               Index           =   0
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   93
               Top             =   1680
               Width           =   4695
            End
            Begin VB.TextBox ctxt戒烟 
               Height          =   270
               Index           =   0
               Left            =   960
               TabIndex        =   90
               Top             =   1200
               Width           =   975
            End
            Begin VB.TextBox ctxt烟龄 
               Height          =   270
               Index           =   0
               Left            =   960
               TabIndex        =   88
               Top             =   840
               Width           =   975
            End
            Begin VB.TextBox ctxt酒龄 
               Height          =   270
               Index           =   0
               Left            =   3360
               TabIndex        =   89
               Top             =   960
               Width           =   975
            End
            Begin VB.TextBox ctxt吸烟量 
               Height          =   270
               Index           =   0
               Left            =   960
               TabIndex        =   92
               Top             =   480
               Width           =   975
            End
            Begin VB.TextBox ctxt饮酒量 
               Height          =   270
               Index           =   0
               Left            =   3360
               TabIndex        =   91
               Top             =   600
               Width           =   975
            End
            Begin VB.ComboBox ccmb饮酒 
               Height          =   300
               Index           =   0
               Left            =   3360
               TabIndex        =   87
               Top             =   120
               Width           =   1335
            End
            Begin VB.ComboBox ccmb吸烟 
               Height          =   300
               Index           =   0
               Left            =   960
               TabIndex        =   86
               Top             =   120
               Width           =   1335
            End
            Begin VB.Label Label101 
               AutoSize        =   -1  'True
               Caption         =   "ML/日"
               Height          =   300
               Index           =   0
               Left            =   4440
               TabIndex        =   190
               Top             =   600
               Width           =   810
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               Caption         =   "多年居住地区、饮食习惯、烟酒嗜好用量："
               Height          =   180
               Index           =   1
               Left            =   120
               TabIndex        =   241
               Top             =   1440
               Width           =   3420
            End
            Begin VB.Label Label110 
               AutoSize        =   -1  'True
               Caption         =   "年"
               Height          =   180
               Index           =   0
               Left            =   2040
               TabIndex        =   199
               Top             =   1200
               Width           =   180
            End
            Begin VB.Label Label109 
               AutoSize        =   -1  'True
               Caption         =   "年"
               Height          =   180
               Index           =   0
               Left            =   4440
               TabIndex        =   198
               Top             =   960
               Width           =   180
            End
            Begin VB.Label Label108 
               AutoSize        =   -1  'True
               Caption         =   "年"
               Height          =   180
               Index           =   0
               Left            =   2040
               TabIndex        =   197
               Top             =   840
               Width           =   180
            End
            Begin VB.Label Label107 
               AutoSize        =   -1  'True
               Caption         =   "戒烟时长："
               Height          =   180
               Index           =   0
               Left            =   120
               TabIndex        =   196
               Top             =   1200
               Width           =   900
            End
            Begin VB.Label Label106 
               AutoSize        =   -1  'True
               Caption         =   "烟龄："
               Height          =   180
               Index           =   0
               Left            =   240
               TabIndex        =   195
               Top             =   840
               Width           =   540
            End
            Begin VB.Label Label105 
               AutoSize        =   -1  'True
               Caption         =   "酒龄："
               Height          =   180
               Index           =   0
               Left            =   2640
               TabIndex        =   194
               Top             =   960
               Width           =   540
            End
            Begin VB.Label Label104 
               AutoSize        =   -1  'True
               Caption         =   "吸烟量："
               Height          =   180
               Index           =   0
               Left            =   120
               TabIndex        =   193
               Top             =   480
               Width           =   720
            End
            Begin VB.Label Label103 
               AutoSize        =   -1  'True
               Caption         =   "饮酒量："
               Height          =   180
               Index           =   0
               Left            =   2640
               TabIndex        =   192
               Top             =   600
               Width           =   720
            End
            Begin VB.Label Label102 
               AutoSize        =   -1  'True
               Caption         =   "支/天"
               Height          =   180
               Index           =   0
               Left            =   2040
               TabIndex        =   191
               Top             =   480
               Width           =   450
            End
            Begin VB.Label Label100 
               AutoSize        =   -1  'True
               Caption         =   "饮酒程度："
               Height          =   180
               Index           =   0
               Left            =   2520
               TabIndex        =   189
               Top             =   195
               Width           =   900
            End
            Begin VB.Label Label99 
               AutoSize        =   -1  'True
               Caption         =   "吸烟程度："
               Height          =   180
               Index           =   0
               Left            =   120
               TabIndex        =   188
               Top             =   195
               Width           =   900
            End
         End
         Begin VB.ComboBox Combo11 
            Height          =   300
            ItemData        =   "frmCareerHstRegister.frx":00A8
            Left            =   5040
            List            =   "frmCareerHstRegister.frx":00B5
            TabIndex        =   275
            Text            =   "健康"
            Top             =   360
            Width           =   855
         End
         Begin VB.Frame Frame2 
            Caption         =   "婚姻史"
            ForeColor       =   &H000080FF&
            Height          =   1095
            Index           =   0
            Left            =   120
            TabIndex        =   167
            Top             =   120
            Width           =   11055
            Begin VB.TextBox ctxtmarrydate 
               Height          =   270
               Index           =   0
               Left            =   960
               TabIndex        =   276
               Text            =   "  年 月"
               Top             =   720
               Width           =   1575
            End
            Begin VB.TextBox ctxtmatehelh 
               Height          =   270
               Index           =   0
               Left            =   3960
               TabIndex        =   71
               Text            =   "健康"
               Top             =   240
               Width           =   855
            End
            Begin VB.TextBox ctxtmateradioac 
               Height          =   495
               Index           =   0
               Left            =   5880
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   74
               Text            =   "frmCareerHstRegister.frx":00CB
               Top             =   480
               Width           =   4815
            End
            Begin VB.TextBox ctxtmatejob 
               Height          =   270
               Index           =   0
               Left            =   3960
               TabIndex        =   73
               Top             =   720
               Width           =   1695
            End
            Begin VB.ComboBox Ccmb婚否 
               Height          =   300
               Index           =   0
               Left            =   960
               TabIndex        =   70
               Top             =   240
               Width           =   1695
            End
            Begin MSComCtl2.DTPicker ctxtmarrydate1 
               CausesValidation=   0   'False
               Height          =   300
               Index           =   0
               Left            =   7560
               TabIndex        =   72
               Top             =   120
               Visible         =   0   'False
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   0   'False
               CustomFormat    =   "yyyy/MM"
               Format          =   60030976
               CurrentDate     =   41013
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "配偶健康状况："
               Height          =   180
               Index           =   0
               Left            =   2760
               TabIndex        =   172
               Top             =   300
               Width           =   1260
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "配偶接触放射线情况："
               Height          =   180
               Index           =   0
               Left            =   5880
               TabIndex        =   171
               Top             =   240
               Width           =   1800
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "配偶职业："
               Height          =   180
               Index           =   0
               Left            =   2760
               TabIndex        =   170
               Top             =   760
               Width           =   900
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "结婚日期："
               Height          =   180
               Index           =   0
               Left            =   120
               TabIndex        =   169
               Top             =   760
               Width           =   900
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "是否结婚："
               Height          =   180
               Index           =   0
               Left            =   120
               TabIndex        =   168
               Top             =   300
               Width           =   900
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "生育史(或配偶生育史)"
            ForeColor       =   &H000080FF&
            Height          =   2295
            Index           =   0
            Left            =   120
            TabIndex        =   155
            Top             =   1440
            Width           =   5775
            Begin VB.TextBox ctxt女孩出生日期 
               Height          =   270
               Index           =   0
               Left            =   2880
               TabIndex        =   285
               Top             =   1920
               Width           =   855
            End
            Begin VB.TextBox ctxt现有女孩 
               Height          =   270
               Index           =   0
               Left            =   1200
               TabIndex        =   283
               Text            =   "0"
               Top             =   1920
               Width           =   495
            End
            Begin VB.TextBox ctxt男孩出生日期 
               Height          =   270
               Index           =   0
               Left            =   2880
               TabIndex        =   281
               Top             =   1560
               Width           =   855
            End
            Begin VB.ComboBox Combo8 
               Height          =   300
               ItemData        =   "frmCareerHstRegister.frx":00D0
               Left            =   4800
               List            =   "frmCareerHstRegister.frx":00DD
               TabIndex        =   269
               Text            =   "健康"
               Top             =   1800
               Width           =   855
            End
            Begin VB.ComboBox Combo7 
               Height          =   300
               ItemData        =   "frmCareerHstRegister.frx":00F3
               Left            =   3960
               List            =   "frmCareerHstRegister.frx":0103
               TabIndex        =   268
               Text            =   "不孕不育原因模板"
               Top             =   1080
               Width           =   1815
            End
            Begin VB.TextBox ctxt不孕不育 
               Height          =   375
               Index           =   0
               Left            =   2040
               MultiLine       =   -1  'True
               TabIndex        =   85
               Text            =   "frmCareerHstRegister.frx":0135
               Top             =   1080
               Width           =   1815
            End
            Begin VB.TextBox ctxt死产 
               Height          =   270
               Index           =   0
               Left            =   2880
               TabIndex        =   76
               Text            =   "0"
               Top             =   480
               Width           =   735
            End
            Begin VB.TextBox ctxt早产 
               Height          =   270
               Index           =   0
               Left            =   1920
               TabIndex        =   75
               Text            =   "0"
               Top             =   480
               Width           =   735
            End
            Begin VB.TextBox ctxt活产 
               Height          =   270
               Index           =   0
               Left            =   1080
               TabIndex        =   79
               Text            =   "0"
               Top             =   480
               Width           =   735
            End
            Begin VB.TextBox ctxt异位妊娠 
               Height          =   270
               Index           =   0
               Left            =   1080
               TabIndex        =   83
               Text            =   "0"
               Top             =   1080
               Width           =   855
            End
            Begin VB.TextBox ctxt孕次 
               Height          =   270
               Index           =   0
               Left            =   240
               TabIndex        =   77
               Text            =   "0"
               Top             =   480
               Width           =   735
            End
            Begin VB.TextBox ctxt多胎 
               Height          =   270
               Index           =   0
               Left            =   240
               TabIndex        =   78
               Text            =   "0"
               Top             =   1080
               Width           =   735
            End
            Begin VB.TextBox ctxt畸胎 
               Height          =   270
               Index           =   0
               Left            =   4800
               TabIndex        =   84
               Text            =   "0"
               Top             =   480
               Width           =   735
            End
            Begin VB.TextBox ctxt流产 
               Height          =   270
               Index           =   0
               Left            =   3720
               TabIndex        =   80
               Text            =   "0"
               Top             =   480
               Width           =   855
            End
            Begin VB.TextBox ctxt现有子女 
               Height          =   270
               Index           =   0
               Left            =   1200
               TabIndex        =   81
               Text            =   "0"
               Top             =   1560
               Width           =   495
            End
            Begin VB.TextBox ctxt子女健康 
               Height          =   270
               Index           =   0
               Left            =   3960
               TabIndex        =   82
               Text            =   "健康"
               Top             =   1800
               Width           =   855
            End
            Begin VB.Label Label73 
               Caption         =   "出生日期："
               Height          =   255
               Left            =   1920
               TabIndex        =   284
               Top             =   1920
               Width           =   975
            End
            Begin VB.Label Label72 
               Caption         =   "现有女孩："
               Height          =   255
               Left            =   240
               TabIndex        =   282
               Top             =   1920
               Width           =   975
            End
            Begin VB.Label Label71 
               Caption         =   "出生日期："
               Height          =   255
               Left            =   1920
               TabIndex        =   280
               Top             =   1560
               Width           =   975
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               Caption         =   "不孕不育原因："
               Height          =   180
               Index           =   0
               Left            =   2040
               TabIndex        =   166
               Top             =   840
               Width           =   1500
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "死产："
               Height          =   180
               Index           =   0
               Left            =   2880
               TabIndex        =   165
               Top             =   240
               Width           =   540
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "早产："
               Height          =   180
               Index           =   0
               Left            =   1920
               TabIndex        =   164
               Top             =   240
               Width           =   540
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "活产："
               Height          =   180
               Index           =   0
               Left            =   1080
               TabIndex        =   163
               Top             =   240
               Width           =   540
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               Caption         =   "异位妊娠："
               Height          =   180
               Index           =   0
               Left            =   1080
               TabIndex        =   162
               Top             =   840
               Width           =   900
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "孕次："
               Height          =   180
               Index           =   0
               Left            =   240
               TabIndex        =   161
               Top             =   240
               Width           =   540
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               Caption         =   "多胎："
               Height          =   180
               Index           =   0
               Left            =   240
               TabIndex        =   160
               Top             =   840
               Width           =   540
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               Caption         =   "畸胎："
               Height          =   180
               Index           =   0
               Left            =   4800
               TabIndex        =   159
               Top             =   240
               Width           =   540
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               Caption         =   "流产："
               Height          =   180
               Index           =   0
               Left            =   3720
               TabIndex        =   158
               Top             =   240
               Width           =   540
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               Caption         =   "现有男孩："
               Height          =   180
               Index           =   0
               Left            =   240
               TabIndex        =   157
               Top             =   1560
               Width           =   900
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               Caption         =   "子女健康状况："
               Height          =   180
               Index           =   0
               Left            =   3960
               TabIndex        =   156
               Top             =   1560
               Width           =   1260
            End
         End
      End
      Begin VB.Frame freOrdinary 
         BorderStyle     =   0  'None
         Height          =   3735
         Left            =   -74640
         TabIndex        =   207
         Top             =   600
         Width           =   11175
         Begin VB.TextBox ctxt出生地 
            Height          =   375
            Index           =   2
            Left            =   6840
            TabIndex        =   307
            Top             =   1800
            Width           =   4215
         End
         Begin VB.Frame Frame3 
            Caption         =   "生育史(或配偶生育史)"
            ForeColor       =   &H000080FF&
            Height          =   2655
            Index           =   2
            Left            =   120
            TabIndex        =   224
            Top             =   960
            Width           =   5775
            Begin VB.TextBox ctxt异常胎 
               Height          =   270
               Left            =   1680
               TabIndex        =   61
               Text            =   "0"
               Top             =   1080
               Width           =   1215
            End
            Begin VB.TextBox ctxt现有子女 
               Height          =   270
               Index           =   2
               Left            =   1680
               TabIndex        =   57
               Text            =   "0"
               Top             =   360
               Width           =   1215
            End
            Begin VB.TextBox ctxt流产 
               Height          =   270
               Index           =   2
               Left            =   1680
               TabIndex        =   59
               Text            =   "0"
               Top             =   720
               Width           =   1215
            End
            Begin VB.TextBox ctxt早产 
               Height          =   270
               Index           =   2
               Left            =   4200
               TabIndex        =   58
               Text            =   "0"
               Top             =   360
               Width           =   975
            End
            Begin VB.TextBox ctxt死产 
               Height          =   270
               Index           =   2
               Left            =   4200
               TabIndex        =   60
               Text            =   "0"
               Top             =   720
               Width           =   975
            End
            Begin VB.Label Label异常胎 
               AutoSize        =   -1  'True
               Caption         =   "异常胎："
               Height          =   180
               Left            =   840
               TabIndex        =   229
               Top             =   1080
               Width           =   720
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               Caption         =   "现有子女数目："
               Height          =   180
               Index           =   1
               Left            =   480
               TabIndex        =   228
               Top             =   360
               Width           =   1260
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               Caption         =   "流产："
               Height          =   180
               Index           =   1
               Left            =   1080
               TabIndex        =   227
               Top             =   720
               Width           =   540
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "早产："
               Height          =   180
               Index           =   1
               Left            =   3480
               TabIndex        =   226
               Top             =   360
               Width           =   540
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "死产："
               Height          =   180
               Index           =   1
               Left            =   3480
               TabIndex        =   225
               Top             =   720
               Width           =   540
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "婚姻史"
            ForeColor       =   &H000080FF&
            Height          =   615
            Index           =   2
            Left            =   120
            TabIndex        =   222
            Top             =   240
            Width           =   5775
            Begin VB.ComboBox Ccmb婚否 
               Height          =   300
               Index           =   2
               Left            =   1680
               TabIndex        =   51
               Top             =   240
               Width           =   1695
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "是否结婚："
               Height          =   180
               Index           =   2
               Left            =   480
               TabIndex        =   223
               Top             =   300
               Width           =   900
            End
         End
         Begin VB.Frame Frame19 
            Caption         =   "烟酒史"
            ForeColor       =   &H000080FF&
            Height          =   1455
            Index           =   1
            Left            =   6000
            TabIndex        =   209
            Top             =   240
            Width           =   5055
            Begin VB.ComboBox ccmb吸烟 
               Height          =   300
               Index           =   2
               Left            =   960
               TabIndex        =   62
               Top             =   120
               Width           =   1335
            End
            Begin VB.ComboBox ccmb饮酒 
               Height          =   300
               Index           =   2
               Left            =   3360
               TabIndex        =   63
               Top             =   120
               Width           =   1335
            End
            Begin VB.TextBox ctxt饮酒量 
               Height          =   270
               Index           =   2
               Left            =   3360
               TabIndex        =   68
               Top             =   600
               Width           =   975
            End
            Begin VB.TextBox ctxt吸烟量 
               Height          =   270
               Index           =   2
               Left            =   960
               TabIndex        =   67
               Top             =   480
               Width           =   975
            End
            Begin VB.TextBox ctxt酒龄 
               Height          =   270
               Index           =   2
               Left            =   3360
               TabIndex        =   65
               Top             =   960
               Width           =   975
            End
            Begin VB.TextBox ctxt烟龄 
               Height          =   270
               Index           =   2
               Left            =   960
               TabIndex        =   64
               Top             =   840
               Width           =   975
            End
            Begin VB.TextBox ctxt戒烟 
               Height          =   270
               Index           =   2
               Left            =   960
               TabIndex        =   66
               Top             =   1200
               Width           =   975
            End
            Begin VB.Label Label99 
               AutoSize        =   -1  'True
               Caption         =   "吸烟程度："
               Height          =   180
               Index           =   1
               Left            =   120
               TabIndex        =   221
               Top             =   195
               Width           =   900
            End
            Begin VB.Label Label100 
               AutoSize        =   -1  'True
               Caption         =   "饮酒程度："
               Height          =   180
               Index           =   1
               Left            =   2520
               TabIndex        =   220
               Top             =   195
               Width           =   900
            End
            Begin VB.Label Label101 
               AutoSize        =   -1  'True
               Caption         =   "ML/日"
               Height          =   180
               Index           =   1
               Left            =   4440
               TabIndex        =   219
               Top             =   600
               Width           =   450
            End
            Begin VB.Label Label102 
               AutoSize        =   -1  'True
               Caption         =   "支/天"
               Height          =   180
               Index           =   1
               Left            =   2040
               TabIndex        =   218
               Top             =   480
               Width           =   450
            End
            Begin VB.Label Label103 
               AutoSize        =   -1  'True
               Caption         =   "饮酒量："
               Height          =   180
               Index           =   1
               Left            =   2640
               TabIndex        =   217
               Top             =   600
               Width           =   720
            End
            Begin VB.Label Label104 
               AutoSize        =   -1  'True
               Caption         =   "吸烟量："
               Height          =   180
               Index           =   1
               Left            =   120
               TabIndex        =   216
               Top             =   480
               Width           =   720
            End
            Begin VB.Label Label105 
               AutoSize        =   -1  'True
               Caption         =   "酒龄："
               Height          =   180
               Index           =   1
               Left            =   2640
               TabIndex        =   215
               Top             =   960
               Width           =   540
            End
            Begin VB.Label Label106 
               AutoSize        =   -1  'True
               Caption         =   "烟龄："
               Height          =   180
               Index           =   1
               Left            =   120
               TabIndex        =   214
               Top             =   840
               Width           =   540
            End
            Begin VB.Label Label107 
               AutoSize        =   -1  'True
               Caption         =   "戒烟时长："
               Height          =   180
               Index           =   1
               Left            =   120
               TabIndex        =   213
               Top             =   1200
               Width           =   900
            End
            Begin VB.Label Label108 
               AutoSize        =   -1  'True
               Caption         =   "年"
               Height          =   180
               Index           =   1
               Left            =   2040
               TabIndex        =   212
               Top             =   840
               Width           =   180
            End
            Begin VB.Label Label109 
               AutoSize        =   -1  'True
               Caption         =   "年"
               Height          =   180
               Index           =   1
               Left            =   4440
               TabIndex        =   211
               Top             =   960
               Width           =   180
            End
            Begin VB.Label Label110 
               AutoSize        =   -1  'True
               Caption         =   "年"
               Height          =   180
               Index           =   1
               Left            =   2040
               TabIndex        =   210
               Top             =   1200
               Width           =   180
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "过敏史"
            ForeColor       =   &H000080FF&
            Height          =   1215
            Index           =   1
            Left            =   6000
            TabIndex        =   208
            Top             =   2280
            Width           =   5055
            Begin VB.ComboBox Combo2 
               Height          =   300
               Left            =   3240
               TabIndex        =   252
               Text            =   "Combo2"
               Top             =   600
               Width           =   1575
            End
            Begin VB.TextBox ctxt过敏史 
               Height          =   735
               Index           =   2
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   69
               Top             =   240
               Width           =   3135
            End
            Begin VB.Label Label63 
               Caption         =   "请选择过敏药源:"
               Height          =   255
               Left            =   3240
               TabIndex        =   253
               Top             =   240
               Width           =   1455
            End
         End
         Begin VB.Label Lab出生地 
            Caption         =   "出生地："
            Height          =   255
            Left            =   6120
            TabIndex        =   306
            Top             =   1920
            Width           =   735
         End
      End
      Begin VB.Frame freNuclear 
         BorderStyle     =   0  'None
         Height          =   3615
         Left            =   -74760
         TabIndex        =   173
         Top             =   600
         Width           =   10815
         Begin VB.Frame Frame2 
            Caption         =   "婚姻史"
            ForeColor       =   &H000080FF&
            Height          =   1095
            Index           =   1
            Left            =   120
            TabIndex        =   201
            Top             =   120
            Width           =   10935
            Begin VB.TextBox ctxtmarrydate 
               Height          =   270
               Index           =   1
               Left            =   960
               TabIndex        =   286
               Text            =   "年月"
               Top             =   720
               Width           =   1695
            End
            Begin VB.ComboBox Ccmb婚否 
               Height          =   300
               Index           =   1
               Left            =   960
               TabIndex        =   40
               Top             =   240
               Width           =   1695
            End
            Begin VB.TextBox ctxtmatejob 
               Height          =   270
               Index           =   1
               Left            =   3960
               TabIndex        =   43
               Top             =   720
               Width           =   1695
            End
            Begin VB.TextBox ctxtmateradioac 
               Height          =   495
               Index           =   1
               Left            =   5880
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   44
               Top             =   480
               Width           =   4815
            End
            Begin VB.TextBox ctxtmatehelh 
               Height          =   270
               Index           =   1
               Left            =   3960
               TabIndex        =   41
               Text            =   "健康"
               Top             =   240
               Width           =   1695
            End
            Begin MSComCtl2.DTPicker ctxtmarrydate2 
               Height          =   300
               Index           =   1
               Left            =   7680
               TabIndex        =   42
               Top             =   120
               Visible         =   0   'False
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   529
               _Version        =   393216
               CustomFormat    =   "yyyy/MM/dd"
               Format          =   60030976
               CurrentDate     =   41013
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "是否结婚："
               Height          =   180
               Index           =   1
               Left            =   120
               TabIndex        =   206
               Top             =   300
               Width           =   900
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "结婚日期："
               Height          =   180
               Index           =   1
               Left            =   120
               TabIndex        =   205
               Top             =   760
               Width           =   900
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "配偶职业："
               Height          =   180
               Index           =   1
               Left            =   2760
               TabIndex        =   204
               Top             =   760
               Width           =   900
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "配偶接触放射线情况："
               Height          =   180
               Index           =   1
               Left            =   5880
               TabIndex        =   203
               Top             =   240
               Width           =   1800
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "配偶健康状况："
               Height          =   180
               Index           =   1
               Left            =   2760
               TabIndex        =   202
               Top             =   300
               Width           =   1260
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "生育史(或配偶生育史)"
            ForeColor       =   &H000080FF&
            Height          =   2295
            Index           =   1
            Left            =   120
            TabIndex        =   39
            Top             =   1320
            Width           =   5775
            Begin VB.TextBox ctxt流产 
               Height          =   270
               Index           =   1
               Left            =   120
               TabIndex        =   305
               Top             =   1200
               Width           =   735
            End
            Begin VB.TextBox ctxt女孩出生日期 
               Height          =   270
               Index           =   1
               Left            =   2280
               TabIndex        =   297
               Top             =   2040
               Width           =   1215
            End
            Begin VB.TextBox ctxt男孩出生日期 
               Height          =   270
               Index           =   1
               Left            =   2280
               TabIndex        =   296
               Top             =   1680
               Width           =   1215
            End
            Begin VB.TextBox ctxt现有女孩 
               Height          =   270
               Index           =   1
               Left            =   960
               TabIndex        =   293
               Top             =   2040
               Width           =   495
            End
            Begin VB.TextBox ctxt现有子女 
               Height          =   270
               Index           =   1
               Left            =   960
               TabIndex        =   292
               Top             =   1680
               Width           =   495
            End
            Begin VB.ComboBox Combo10 
               Height          =   300
               ItemData        =   "frmCareerHstRegister.frx":0139
               Left            =   4680
               List            =   "frmCareerHstRegister.frx":0146
               TabIndex        =   271
               Text            =   "健康"
               Top             =   1920
               Width           =   975
            End
            Begin VB.ComboBox Combo9 
               Height          =   300
               ItemData        =   "frmCareerHstRegister.frx":015C
               Left            =   3600
               List            =   "frmCareerHstRegister.frx":016C
               TabIndex        =   270
               Text            =   "不孕不育原因模板"
               Top             =   1200
               Width           =   1935
            End
            Begin VB.TextBox ctxt不孕不育 
               Height          =   270
               Index           =   1
               Left            =   960
               MultiLine       =   -1  'True
               TabIndex        =   53
               Text            =   "frmCareerHstRegister.frx":019E
               Top             =   1200
               Width           =   2535
            End
            Begin VB.TextBox ctxt死产 
               Height          =   270
               Index           =   1
               Left            =   2640
               TabIndex        =   48
               Text            =   "0"
               Top             =   600
               Width           =   735
            End
            Begin VB.TextBox ctxt早产 
               Height          =   270
               Index           =   1
               Left            =   1800
               TabIndex        =   47
               Text            =   "0"
               Top             =   600
               Width           =   735
            End
            Begin VB.TextBox ctxt活产 
               Height          =   270
               Index           =   1
               Left            =   960
               TabIndex        =   46
               Text            =   " "
               Top             =   600
               Width           =   735
            End
            Begin VB.TextBox ctxt孕次 
               Height          =   270
               Index           =   1
               Left            =   120
               TabIndex        =   45
               Text            =   "0"
               Top             =   600
               Width           =   735
            End
            Begin VB.TextBox ctxt多胎 
               Height          =   270
               Index           =   1
               Left            =   3600
               TabIndex        =   49
               Text            =   "0"
               Top             =   600
               Width           =   975
            End
            Begin VB.TextBox ctxt畸胎 
               Height          =   270
               Index           =   1
               Left            =   4560
               TabIndex        =   50
               Text            =   "0"
               Top             =   600
               Width           =   1095
            End
            Begin VB.TextBox ctxt子女健康 
               Height          =   270
               Index           =   1
               Left            =   3600
               TabIndex        =   52
               Top             =   1920
               Width           =   975
            End
            Begin VB.Label Label95 
               Caption         =   "流产："
               Height          =   255
               Left            =   120
               TabIndex        =   304
               Top             =   960
               Width           =   855
            End
            Begin VB.Label Label78 
               Caption         =   "出生日期："
               Height          =   225
               Left            =   1440
               TabIndex        =   295
               Top             =   2040
               Width           =   975
            End
            Begin VB.Label Label77 
               Caption         =   "出生日期："
               Height          =   225
               Left            =   1440
               TabIndex        =   294
               Top             =   1680
               Width           =   975
            End
            Begin VB.Label Label76 
               Caption         =   "现有女孩："
               Height          =   255
               Left            =   120
               TabIndex        =   291
               Top             =   2040
               Width           =   975
            End
            Begin VB.Label Label75 
               Caption         =   "现有男孩："
               Height          =   255
               Left            =   120
               TabIndex        =   290
               Top             =   1680
               Width           =   975
            End
            Begin VB.Label Label79 
               AutoSize        =   -1  'True
               Caption         =   "不孕不育原因："
               Height          =   180
               Left            =   960
               TabIndex        =   184
               Top             =   960
               Width           =   1260
            End
            Begin VB.Label Label80 
               AutoSize        =   -1  'True
               Caption         =   "死产："
               Height          =   180
               Left            =   2640
               TabIndex        =   183
               Top             =   360
               Width           =   540
            End
            Begin VB.Label Label81 
               AutoSize        =   -1  'True
               Caption         =   "早产："
               Height          =   180
               Left            =   1800
               TabIndex        =   182
               Top             =   360
               Width           =   540
            End
            Begin VB.Label Label82 
               AutoSize        =   -1  'True
               Caption         =   "活产："
               Height          =   180
               Left            =   960
               TabIndex        =   181
               Top             =   360
               Width           =   540
            End
            Begin VB.Label Label84 
               AutoSize        =   -1  'True
               Caption         =   "孕次："
               Height          =   180
               Left            =   120
               TabIndex        =   180
               Top             =   360
               Width           =   540
            End
            Begin VB.Label Label85 
               AutoSize        =   -1  'True
               Caption         =   "多胎："
               Height          =   180
               Left            =   3480
               TabIndex        =   179
               Top             =   360
               Width           =   540
            End
            Begin VB.Label Label86 
               AutoSize        =   -1  'True
               Caption         =   "畸胎："
               Height          =   180
               Left            =   4560
               TabIndex        =   178
               Top             =   360
               Width           =   540
            End
            Begin VB.Label Label89 
               AutoSize        =   -1  'True
               Caption         =   "子女健康状况："
               Height          =   300
               Left            =   3600
               TabIndex        =   177
               Top             =   1680
               Width           =   1260
            End
         End
         Begin VB.Frame Frame17 
            Caption         =   "烟酒史"
            ForeColor       =   &H000080FF&
            Height          =   2295
            Left            =   5880
            TabIndex        =   174
            Top             =   1320
            Width           =   5055
            Begin VB.TextBox ctxt戒烟 
               Height          =   270
               Index           =   1
               Left            =   960
               TabIndex        =   313
               Top             =   2040
               Width           =   855
            End
            Begin VB.ComboBox ccmb饮酒 
               Height          =   300
               Index           =   1
               Left            =   3360
               TabIndex        =   312
               Top             =   960
               Width           =   1455
            End
            Begin VB.ComboBox ccmb吸烟 
               Height          =   300
               Index           =   1
               Left            =   960
               TabIndex        =   311
               Top             =   960
               Width           =   1335
            End
            Begin VB.TextBox ctxt烟龄 
               Height          =   270
               Index           =   1
               Left            =   840
               TabIndex        =   302
               Top             =   1680
               Width           =   975
            End
            Begin VB.TextBox ctxt酒龄 
               Height          =   270
               Index           =   1
               Left            =   3240
               TabIndex        =   299
               Top             =   1680
               Width           =   975
            End
            Begin VB.TextBox ctxtMore 
               Height          =   375
               Index           =   1
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   54
               Top             =   480
               Width           =   4695
            End
            Begin VB.TextBox ctxt饮酒量 
               Height          =   270
               Index           =   1
               Left            =   3360
               TabIndex        =   55
               Top             =   1320
               Width           =   975
            End
            Begin VB.TextBox ctxt吸烟量 
               Height          =   270
               Index           =   1
               Left            =   840
               TabIndex        =   56
               Top             =   1320
               Width           =   975
            End
            Begin VB.Label Label111 
               Caption         =   "年"
               Height          =   255
               Left            =   1920
               TabIndex        =   314
               Top             =   2040
               Width           =   375
            End
            Begin VB.Label Label98 
               Caption         =   "戒烟时长："
               Height          =   255
               Left            =   120
               TabIndex        =   310
               Top             =   2040
               Width           =   1215
            End
            Begin VB.Label Label97 
               Caption         =   "饮酒程度："
               Height          =   255
               Left            =   2520
               TabIndex        =   309
               Top             =   960
               Width           =   1095
            End
            Begin VB.Label Label96 
               Caption         =   "吸烟程度："
               Height          =   255
               Left            =   120
               TabIndex        =   308
               Top             =   960
               Width           =   975
            End
            Begin VB.Label Label94 
               Caption         =   "年"
               Height          =   255
               Left            =   4320
               TabIndex        =   303
               Top             =   1680
               Width           =   495
            End
            Begin VB.Label Label93 
               Caption         =   "烟龄："
               Height          =   255
               Left            =   120
               TabIndex        =   301
               Top             =   1680
               Width           =   735
            End
            Begin VB.Label Label92 
               Caption         =   "年"
               Height          =   255
               Left            =   1920
               TabIndex        =   300
               Top             =   1680
               Width           =   495
            End
            Begin VB.Label Label88 
               Caption         =   "酒龄："
               Height          =   255
               Left            =   2520
               TabIndex        =   298
               Top             =   1680
               Width           =   720
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               Caption         =   "多年居住地区、饮食习惯、烟酒嗜好用量："
               Height          =   180
               Index           =   0
               Left            =   120
               TabIndex        =   200
               Top             =   240
               Width           =   3420
            End
            Begin VB.Label Label87 
               AutoSize        =   -1  'True
               Caption         =   "ML/日"
               Height          =   180
               Left            =   4440
               TabIndex        =   186
               Top             =   1320
               Width           =   450
            End
            Begin VB.Label Label83 
               AutoSize        =   -1  'True
               Caption         =   "支/天"
               Height          =   180
               Left            =   1920
               TabIndex        =   185
               Top             =   1320
               Width           =   450
            End
            Begin VB.Label Label90 
               AutoSize        =   -1  'True
               Caption         =   "饮酒量："
               Height          =   180
               Left            =   2520
               TabIndex        =   176
               Top             =   1320
               Width           =   720
            End
            Begin VB.Label Label91 
               AutoSize        =   -1  'True
               Caption         =   "吸烟量："
               Height          =   180
               Left            =   120
               TabIndex        =   175
               Top             =   1320
               Width           =   720
            End
         End
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H0000C000&
         Caption         =   "确  定"
         Height          =   375
         Left            =   -64800
         Style           =   1  'Graphical
         TabIndex        =   272
         Top             =   4320
         Width           =   1095
      End
      Begin VB.Frame Frame5 
         Caption         =   "心率录入"
         ForeColor       =   &H000080FF&
         Height          =   855
         Left            =   720
         TabIndex        =   255
         Top             =   4080
         Width           =   5295
         Begin VB.CommandButton Command3 
            BackColor       =   &H8000000B&
            Caption         =   "确定"
            Height          =   375
            Left            =   4200
            TabIndex        =   259
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox ctxtxinlv 
            Height          =   270
            Left            =   960
            TabIndex        =   256
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label66 
            Height          =   255
            Left            =   3120
            TabIndex        =   266
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label65 
            Caption         =   "心率"
            Height          =   375
            Left            =   240
            TabIndex        =   258
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label64 
            Caption         =   "次/分"
            Height          =   255
            Left            =   2400
            TabIndex        =   257
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00008080&
         Caption         =   "修改"
         Height          =   375
         Left            =   -73200
         TabIndex        =   246
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000C000&
         Caption         =   "新增"
         Height          =   375
         Left            =   -74760
         TabIndex        =   245
         Top             =   480
         Width           =   1215
      End
      Begin VB.Frame Frame6 
         Caption         =   "家族史"
         ForeColor       =   &H000080FF&
         Height          =   975
         Left            =   -68880
         TabIndex        =   232
         Top             =   4320
         Width           =   5175
         Begin VB.TextBox ctxt家族 
            Height          =   615
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   0
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label Label27 
            Caption         =   "提示:家族中有无遗传性疾病、血液病、糖尿病、高血压病、神经精神性疾病、肿瘤、结核病等"
            Height          =   615
            Left            =   2520
            TabIndex        =   250
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "月经史"
         ForeColor       =   &H000080FF&
         Height          =   975
         Left            =   -74640
         TabIndex        =   233
         Top             =   4320
         Width           =   5775
         Begin VB.TextBox ctxt停经 
            Height          =   270
            Left            =   3360
            TabIndex        =   98
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox ctxt末次月经 
            Height          =   270
            Left            =   960
            TabIndex        =   97
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox ctxt周期 
            Height          =   270
            Left            =   4080
            TabIndex        =   96
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox ctxt经期 
            Height          =   270
            Left            =   2400
            TabIndex        =   95
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox ctxt初潮 
            Height          =   270
            Left            =   600
            TabIndex        =   94
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "停经年龄："
            Height          =   180
            Index           =   2
            Left            =   2400
            TabIndex        =   239
            Top             =   600
            Width           =   900
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "末次月经："
            Height          =   180
            Index           =   2
            Left            =   120
            TabIndex        =   238
            Top             =   600
            Width           =   900
         End
         Begin VB.Label Label4 
            Caption         =   "Label4"
            Height          =   15
            Index           =   2
            Left            =   720
            TabIndex        =   237
            Top             =   720
            Width           =   135
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "周期："
            Height          =   180
            Index           =   2
            Left            =   3600
            TabIndex        =   236
            Top             =   240
            Width           =   540
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "经期："
            Height          =   180
            Index           =   2
            Left            =   1920
            TabIndex        =   235
            Top             =   240
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "初潮："
            Height          =   180
            Index           =   2
            Left            =   120
            TabIndex        =   234
            Top             =   240
            Width           =   540
         End
      End
      Begin VB.TextBox ctxtOther 
         Height          =   495
         Left            =   -74640
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   100
         Top             =   5400
         Width           =   10935
      End
      Begin VB.Frame Frame11 
         Caption         =   "放射性工作史  "
         ForeColor       =   &H000080FF&
         Height          =   1455
         Left            =   -74760
         TabIndex        =   107
         Top             =   2880
         Width           =   11055
         Begin VB.CommandButton ccmdok 
            BackColor       =   &H0000C000&
            Caption         =   "确  定"
            Height          =   375
            Left            =   9960
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   960
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox ctxtfangshe 
            Height          =   270
            Left            =   5400
            TabIndex        =   263
            Top             =   480
            Width           =   1575
         End
         Begin VB.ComboBox Combo1 
            Height          =   300
            ItemData        =   "frmCareerHstRegister.frx":01A2
            Left            =   8520
            List            =   "frmCareerHstRegister.frx":01AC
            TabIndex        =   251
            Text            =   "请选择"
            Top             =   480
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CheckBox Chkokclear 
            Caption         =   "确定后清空"
            Height          =   255
            Left            =   8400
            TabIndex        =   35
            Top             =   1080
            Width           =   1335
         End
         Begin VB.ComboBox ctxt放射种类 
            Height          =   300
            Left            =   6960
            TabIndex        =   32
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox ctxt工作量 
            Height          =   270
            Left            =   5400
            TabIndex        =   34
            Top             =   1080
            Width           =   1815
         End
         Begin VB.TextBox ctxt照射量 
            Height          =   270
            Left            =   9480
            TabIndex        =   33
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox ctxt过量照射 
            Height          =   855
            Left            =   240
            TabIndex        =   31
            Top             =   480
            Width           =   4935
         End
         Begin VB.Label Label41 
            Caption         =   "放射线种类："
            Height          =   255
            Left            =   5400
            TabIndex        =   111
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label40 
            Caption         =   "每日工作时数或工作量："
            Height          =   255
            Left            =   5400
            TabIndex        =   110
            Top             =   840
            Width           =   2055
         End
         Begin VB.Label Label39 
            Caption         =   "累积受照射量："
            Height          =   255
            Left            =   8520
            TabIndex        =   109
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label38 
            Caption         =   "过量照射史："
            Height          =   255
            Left            =   240
            TabIndex        =   108
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "体格一般情况录入"
         ForeColor       =   &H000080FF&
         Height          =   3135
         Left            =   720
         TabIndex        =   143
         Top             =   720
         Width           =   6135
         Begin VB.CommandButton Comd填写 
            Caption         =   "填写身高体重"
            Height          =   495
            Left            =   3720
            TabIndex        =   322
            Top             =   240
            Width           =   1815
         End
         Begin VB.TextBox ctxt体重指数 
            Height          =   270
            Left            =   3960
            TabIndex        =   320
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox ctxt体形 
            Height          =   270
            Left            =   3960
            TabIndex        =   319
            Top             =   840
            Width           =   1335
         End
         Begin VB.ComboBox comb发育 
            Height          =   300
            Left            =   960
            TabIndex        =   316
            Top             =   2760
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.ComboBox comb营养 
            Height          =   300
            Left            =   960
            TabIndex        =   1
            Text            =   "Combo1"
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox ctxt身高 
            Height          =   270
            Left            =   960
            TabIndex        =   2
            Top             =   840
            Width           =   1335
         End
         Begin VB.TextBox ctxt体重 
            Height          =   270
            Left            =   960
            TabIndex        =   3
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox ctxt收缩压 
            Height          =   270
            Left            =   960
            TabIndex        =   4
            Top             =   1800
            Width           =   1335
         End
         Begin VB.TextBox ctxt舒张压 
            Height          =   270
            Left            =   960
            TabIndex        =   5
            Top             =   2280
            Width           =   1335
         End
         Begin VB.Label Lab体格 
            Caption         =   "体重指数"
            Height          =   255
            Index           =   7
            Left            =   3120
            TabIndex        =   318
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Lab体格 
            Caption         =   "体形"
            Height          =   255
            Index           =   6
            Left            =   3120
            TabIndex        =   317
            Top             =   840
            Width           =   495
         End
         Begin VB.Label Lab体格 
            Caption         =   "发育"
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   315
            Top             =   2760
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label Lab体格 
            Caption         =   "舒张压"
            Height          =   375
            Index           =   4
            Left            =   240
            TabIndex        =   153
            Top             =   2280
            Width           =   615
         End
         Begin VB.Label Label56 
            Caption         =   "mmHg"
            Height          =   255
            Left            =   2400
            TabIndex        =   152
            Top             =   2280
            Width           =   375
         End
         Begin VB.Label Lab体格 
            Caption         =   "营养"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   150
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Lab体格 
            Caption         =   "身高"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   149
            Top             =   840
            Width           =   495
         End
         Begin VB.Label Label35 
            Caption         =   "cm"
            Height          =   255
            Index           =   0
            Left            =   2400
            TabIndex        =   148
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Lab体格 
            Caption         =   "体重"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   147
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label Label51 
            Caption         =   "kg"
            Height          =   255
            Left            =   2400
            TabIndex        =   146
            Top             =   1320
            Width           =   375
         End
         Begin VB.Label Lab体格 
            Caption         =   "收缩压"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   145
            Top             =   1800
            Width           =   615
         End
         Begin VB.Label Label55 
            Caption         =   "mmHg"
            Height          =   255
            Left            =   2400
            TabIndex        =   144
            Top             =   1800
            Width           =   375
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid cgrd症状 
         Height          =   2535
         Left            =   -74760
         TabIndex        =   142
         Top             =   3300
         Width           =   10935
         _cx             =   2088782680
         _cy             =   2088767863
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmCareerHstRegister.frx":01C0
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VSFlex8Ctl.VSFlexGrid cgrd病史 
         Height          =   2415
         Left            =   -74760
         TabIndex        =   141
         Top             =   3300
         Width           =   9735
         _cx             =   2088780563
         _cy             =   2088767652
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   "系统编号|编号|疾病名称|诊断日期|诊断单位|治疗经过|转归"
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.CommandButton Ccmd删除 
         BackColor       =   &H008080FF&
         Caption         =   "删  除"
         Height          =   375
         Left            =   -64800
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   5280
         Width           =   1095
      End
      Begin VB.CommandButton Ccmd修改 
         BackColor       =   &H0000C0C0&
         Caption         =   "修  改"
         Height          =   375
         Left            =   -64800
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   4800
         Width           =   1095
      End
      Begin VB.CommandButton Ccmdcancel病史 
         BackColor       =   &H008080FF&
         Caption         =   "删  除"
         Height          =   375
         Left            =   -64920
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   4860
         Width           =   1215
      End
      Begin VB.CommandButton Cmmdmody病史 
         BackColor       =   &H0000C0C0&
         Caption         =   "修  改"
         Height          =   375
         Left            =   -64920
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   4020
         Width           =   1215
      End
      Begin VB.Frame Frame15 
         Caption         =   "自觉症状情况录入   "
         ForeColor       =   &H000080FF&
         Height          =   2535
         Left            =   -74760
         TabIndex        =   127
         Top             =   780
         Width           =   10935
         Begin VB.TextBox ctxt项目 
            Height          =   1935
            Left            =   5040
            MultiLine       =   -1  'True
            TabIndex        =   8
            Top             =   360
            Width           =   2535
         End
         Begin VB.CheckBox Chkokclear症状 
            Caption         =   "保存后清空"
            Height          =   255
            Left            =   1680
            TabIndex        =   10
            Top             =   1860
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            ForeColor       =   &H000000FF&
            Height          =   1935
            Left            =   7920
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   244
            Text            =   "frmCareerHstRegister.frx":0257
            Top             =   360
            Width           =   2535
         End
         Begin VB.CommandButton ccmdcancel症状 
            BackColor       =   &H008080FF&
            Caption         =   "删  除"
            Height          =   375
            Left            =   3000
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   1800
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CommandButton ccmbmody症状 
            BackColor       =   &H0000C0C0&
            Caption         =   "修  改"
            Height          =   375
            Left            =   1680
            Style           =   1  'Graphical
            TabIndex        =   243
            Top             =   1800
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.ComboBox ccmb分类 
            Height          =   300
            Left            =   1320
            TabIndex        =   6
            Top             =   360
            Width           =   2535
         End
         Begin VB.ListBox clst项目 
            Height          =   1950
            Left            =   5040
            Style           =   1  'Checkbox
            TabIndex        =   7
            Top             =   360
            Width           =   2535
         End
         Begin VB.ComboBox ctxt程度 
            Height          =   300
            ItemData        =   "frmCareerHstRegister.frx":02FA
            Left            =   1320
            List            =   "frmCareerHstRegister.frx":02FC
            TabIndex        =   9
            Top             =   840
            Width           =   2535
         End
         Begin VB.CommandButton Ccmdok症状 
            BackColor       =   &H0000C000&
            Caption         =   "确定"
            Height          =   375
            Left            =   360
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   1800
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label25 
            Caption         =   "症状部位："
            Height          =   255
            Left            =   120
            TabIndex        =   242
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label67 
            Caption         =   "症状项目："
            Height          =   255
            Left            =   4080
            TabIndex        =   129
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label68 
            AutoSize        =   -1  'True
            Caption         =   "程度："
            Height          =   180
            Left            =   360
            TabIndex        =   128
            Top             =   840
            Width           =   540
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "工作史  "
         ForeColor       =   &H000080FF&
         Height          =   2055
         Left            =   -74760
         TabIndex        =   112
         Top             =   720
         Width           =   11055
         Begin VB.TextBox ctxt接触 
            Height          =   270
            Left            =   1920
            TabIndex        =   288
            Top             =   1680
            Width           =   1935
         End
         Begin VB.TextBox ctxt结束 
            Height          =   300
            Left            =   8880
            TabIndex        =   278
            Text            =   "年月"
            Top             =   1200
            Width           =   2055
         End
         Begin VB.TextBox ctxt起始 
            Height          =   300
            Left            =   8880
            TabIndex        =   277
            Text            =   "年月"
            Top             =   480
            Width           =   2055
         End
         Begin VB.ComboBox Combo4 
            Height          =   300
            Left            =   240
            TabIndex        =   261
            Text            =   "Combo4"
            Top             =   1200
            Width           =   1575
         End
         Begin VB.TextBox ctxtweihai 
            Height          =   270
            Left            =   6360
            TabIndex        =   260
            Top             =   1200
            Width           =   1815
         End
         Begin MSComCtl2.DTPicker ctxt结束1 
            Height          =   300
            Left            =   5160
            TabIndex        =   30
            Top             =   1680
            Visible         =   0   'False
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            Format          =   60030976
            CurrentDate     =   41013
         End
         Begin MSComCtl2.DTPicker ctxt起始1 
            Height          =   300
            Left            =   4320
            TabIndex        =   26
            Top             =   1560
            Visible         =   0   'False
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            Format          =   60030976
            CurrentDate     =   41013
         End
         Begin MSComCtl2.DTPicker ctxt接触时间 
            Height          =   300
            Left            =   9000
            TabIndex        =   29
            Top             =   1680
            Visible         =   0   'False
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   529
            _Version        =   393216
            Format          =   60030976
            CurrentDate     =   41013
         End
         Begin VB.ComboBox ctxt危害种类 
            Height          =   300
            Left            =   4320
            TabIndex        =   28
            Top             =   1200
            Width           =   1815
         End
         Begin VB.ComboBox ctxt工种 
            Height          =   300
            Left            =   4680
            TabIndex        =   24
            Top             =   480
            Width           =   1575
         End
         Begin VB.ComboBox ctxt部门 
            Height          =   300
            Left            =   2640
            TabIndex        =   23
            Top             =   480
            Width           =   1575
         End
         Begin VB.CheckBox Chk放射 
            Caption         =   "是否放射工作人员"
            Height          =   255
            Left            =   7200
            TabIndex        =   130
            Top             =   1680
            Width           =   1815
         End
         Begin VB.TextBox ctxt备注 
            Height          =   300
            Left            =   6600
            TabIndex        =   25
            Top             =   480
            Width           =   2055
         End
         Begin VB.TextBox ctxt防护 
            Height          =   360
            Left            =   1800
            MultiLine       =   -1  'True
            TabIndex        =   27
            Top             =   1200
            Width           =   2415
         End
         Begin VB.TextBox ctxt单位 
            Height          =   300
            Left            =   240
            TabIndex        =   22
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label Label74 
            Caption         =   "接触时间(小时/周)"
            Height          =   255
            Left            =   240
            TabIndex        =   287
            Top             =   1680
            Width           =   1695
         End
         Begin VB.Label Label28 
            Caption         =   "请选择防护措施"
            Height          =   255
            Left            =   240
            TabIndex        =   262
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label36 
            Caption         =   "接触危害因素结束时间："
            Height          =   255
            Left            =   8880
            TabIndex        =   121
            Top             =   960
            Width           =   2175
         End
         Begin VB.Label Label37 
            Caption         =   "接触危害因素开始时间："
            Height          =   255
            Left            =   8880
            TabIndex        =   120
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label Label43 
            Caption         =   "备注："
            Height          =   255
            Left            =   6600
            TabIndex        =   119
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label44 
            Caption         =   "防护措施："
            Height          =   255
            Left            =   1800
            TabIndex        =   118
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label45 
            Caption         =   "接触危害因素开始时间："
            Height          =   255
            Left            =   9480
            TabIndex        =   117
            Top             =   1560
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Label Label46 
            Caption         =   "接触职业病危害种类："
            Height          =   255
            Left            =   4320
            TabIndex        =   116
            Top             =   960
            Width           =   2055
         End
         Begin VB.Label Label47 
            Caption         =   "工种："
            Height          =   255
            Left            =   4680
            TabIndex        =   115
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label48 
            Caption         =   "部门："
            Height          =   255
            Left            =   2760
            TabIndex        =   114
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label49 
            Caption         =   "工作单位："
            Height          =   255
            Left            =   240
            TabIndex        =   113
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "既往病史情况录入   "
         ForeColor       =   &H000080FF&
         Height          =   2535
         Left            =   -74760
         TabIndex        =   101
         Top             =   720
         Width           =   11055
         Begin VB.TextBox ctxt诊断日期 
            Height          =   330
            Left            =   120
            TabIndex        =   279
            Top             =   1440
            Width           =   2535
         End
         Begin VB.ComboBox Combo6 
            Height          =   300
            Left            =   5040
            TabIndex        =   267
            Text            =   "Combo6"
            Top             =   1440
            Width           =   975
         End
         Begin VB.ComboBox Combo5 
            Height          =   300
            Left            =   6240
            TabIndex        =   265
            Text            =   "Combo5"
            Top             =   720
            Width           =   4095
         End
         Begin VB.ComboBox Combo3 
            Height          =   300
            ItemData        =   "frmCareerHstRegister.frx":02FE
            Left            =   120
            List            =   "frmCareerHstRegister.frx":0300
            TabIndex        =   254
            Text            =   "Combo3"
            Top             =   720
            Width           =   1575
         End
         Begin MSComCtl2.DTPicker ctxt诊断日期1 
            Height          =   330
            Left            =   120
            TabIndex        =   15
            Top             =   2040
            Visible         =   0   'False
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   582
            _Version        =   393216
            Format          =   60030976
            CurrentDate     =   41013
         End
         Begin VB.CheckBox Chkokclear病史 
            Caption         =   "确定后清空"
            Height          =   255
            Left            =   7680
            TabIndex        =   18
            Top             =   2040
            Width           =   1215
         End
         Begin VB.CommandButton ccmdok病史 
            BackColor       =   &H0000C000&
            Caption         =   "确定"
            Height          =   375
            Left            =   9000
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   1920
            Width           =   1335
         End
         Begin VB.TextBox ctxt转归 
            Height          =   330
            Left            =   3480
            TabIndex        =   16
            Top             =   1440
            Width           =   1335
         End
         Begin VB.TextBox ctxt诊断单位 
            Height          =   330
            Left            =   3480
            TabIndex        =   14
            Top             =   720
            Width           =   2535
         End
         Begin VB.TextBox ctxt治疗经过 
            Height          =   615
            Left            =   6240
            TabIndex        =   17
            Top             =   1080
            Width           =   4095
         End
         Begin VB.TextBox ctxt疾病名称 
            Height          =   330
            Left            =   1680
            TabIndex        =   13
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Label54 
            Caption         =   "请选择疾病名称:"
            Height          =   255
            Left            =   120
            TabIndex        =   264
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label58 
            Caption         =   "诊断单位："
            Height          =   255
            Left            =   3480
            TabIndex        =   106
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label59 
            Caption         =   "治疗经过/曾患病描述："
            Height          =   255
            Left            =   6240
            TabIndex        =   105
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label Label60 
            Caption         =   "疾病名称："
            Height          =   255
            Left            =   1680
            TabIndex        =   104
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label61 
            Caption         =   "诊断日期："
            Height          =   255
            Left            =   120
            TabIndex        =   103
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label Label62 
            Caption         =   "转归："
            Height          =   255
            Left            =   3480
            TabIndex        =   102
            Top             =   1200
            Width           =   1095
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid cgrdzzxw 
         Height          =   4335
         Left            =   -74760
         TabIndex        =   247
         Top             =   960
         Width           =   11535
         _cx             =   20346
         _cy             =   7646
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmCareerHstRegister.frx":0302
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VSFlex8Ctl.VSFlexGrid cgrd职业史 
         Height          =   1455
         Left            =   -74760
         TabIndex        =   289
         Top             =   4320
         Width           =   9855
         _cx             =   2088780775
         _cy             =   2088765958
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   16
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmCareerHstRegister.frx":0382
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "其他："
         Height          =   180
         Index           =   0
         Left            =   -74640
         TabIndex        =   240
         Top             =   5460
         Width           =   540
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "个人信息"
      Height          =   1935
      Left            =   120
      TabIndex        =   122
      Top             =   840
      Width           =   13455
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   4320
         ScaleHeight     =   1785
         ScaleWidth      =   1545
         TabIndex        =   123
         Top             =   120
         Width           =   1575
      End
      Begin VB.TextBox ctxtsysno 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   99
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label体检类型 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7200
         TabIndex        =   274
         Top             =   1680
         Width           =   90
      End
      Begin VB.Label Label70 
         Caption         =   "体检  类型："
         Height          =   255
         Left            =   6000
         TabIndex        =   273
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label危害因素 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7200
         TabIndex        =   249
         Top             =   1320
         Width           =   90
      End
      Begin VB.Label Label26 
         Caption         =   "危害  因素："
         Height          =   255
         Left            =   6000
         TabIndex        =   248
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "年龄:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2760
         TabIndex        =   140
         Top             =   1560
         Width           =   570
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "性别:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   360
         TabIndex        =   139
         Top             =   1560
         Width           =   570
      End
      Begin VB.Label Lab年龄 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   138
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Lab性别 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   137
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Lab现职务 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7200
         TabIndex        =   136
         Top             =   960
         Width           =   90
      End
      Begin VB.Label Lab现工种 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7200
         TabIndex        =   135
         Top             =   600
         Width           =   90
      End
      Begin VB.Label Label30 
         Caption         =   "现  职  务："
         Height          =   255
         Index           =   0
         Left            =   6000
         TabIndex        =   134
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label29 
         Caption         =   "现  工  种："
         Height          =   255
         Left            =   6000
         TabIndex        =   133
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lab单位 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7200
         TabIndex        =   132
         Top             =   240
         Width           =   90
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "现工作单位："
         Height          =   255
         Left            =   6000
         TabIndex        =   131
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label Label52 
         AutoSize        =   -1  'True
         Caption         =   "编号:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   360
         TabIndex        =   126
         Top             =   360
         Width           =   570
      End
      Begin VB.Label Lab姓名 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1080
         TabIndex        =   125
         Top             =   840
         Width           =   3015
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         Caption         =   "姓名:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   360
         TabIndex        =   124
         Top             =   960
         Width           =   570
      End
   End
   Begin MSComDlg.CommonDialog ccmdFile 
      Left            =   0
      Top             =   500
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar ctbMain 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   230
      Top             =   0
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   1111
      ButtonWidth     =   820
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "cimg按钮图标"
      _Version        =   393216
      Begin MSCommLib.MSComm MSComm1 
         Left            =   2880
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
      End
      Begin VB.CheckBox ChkClear 
         Caption         =   "保存后清空"
         Height          =   255
         Left            =   9840
         TabIndex        =   231
         Top             =   240
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   3
         Left            =   7440
         Top             =   240
      End
      Begin MSComctlLib.ImageList cimg按钮图标 
         Left            =   0
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "frmCareerHstRegt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'名称：职业病史(受检者个人信息)录入
'函数：Private Sub ctxtsysno_LostFocus()  当系统编号文本框失去焦点，填充
'       体检人员基本信息，包括姓名、性别、照片等
'功能：职业病史(受检者个人信息)录入上信息录入，修改，删除
'作者：Yunle Liu
'时间：2012.03
'********************************************************************

Option Explicit

Private WithEvents mobjGUI As cls界面通用对象
Attribute mobjGUI.VB_VarHelpID = -1
Private mblninuse As Boolean
Private mblnSys As Boolean
Private lobjPsLifeHst As Object
Private lobjPsWorkHst As Object
Private mcolIndex As Collection    '职业史
Private mcolIndexwkdis As Collection   '既往病史
Private mcolindexzz As Collection       '自觉症状
Private mintrow As Integer     '当前修改的行号，处理职业史   modify by lanchao 2015-9-16
Private jmintrow As Integer     '当前修改的行号,处理既往病史 modify by lanchao 2015-9-16
Private lobjInDtBase As Object   '保存到数据库
Private mobj体检 As Object
Private mcol体检项目 As New Collection
Private mIndex As String
Public sysno As String
Public selectsysno As String
Public selectzz As String

Public selectcd As String

Public selectcxrq As String

Private Sub ccmb分类_Change()
    Dim lobjRec As Object
    Dim i As Integer
    If ccmb分类.Text = "" Then clst项目.Clear: Exit Sub
    '获取自觉症状项目
    dasubSetQueryTimeout 6000
    Set lobjRec = dafuncGetData("select 名称 from 系统管理_字典_字典内容表 where Parent in (select InnerID from 系统管理_字典_字典内容表 where 名称 = '" & ccmb分类.Text & " ')")
    clst项目.Clear
'    ccmb分类.AddItem ""
    If ccmb分类.Text = "其他" Or ccmb分类.Text = "" Then
        clst项目.Visible = False
        ctxt项目.Visible = True
    Else
        clst项目.Visible = True
        ctxt项目.Visible = False
    End If
    If (lobjRec.EOF Or lobjRec.BOF) Then
        clst项目.Clear
        clst项目.Visible = False
        ctxt项目.Visible = True
        Exit Sub
    End If
    For i = 1 To lobjRec.RecordCount
        clst项目.AddItem lobjRec("名称")
        clst项目.ItemData(clst项目.NewIndex) = i
        lobjRec.MoveNext
    Next
    clst项目.ListIndex = 0
    
End Sub

Private Sub ccmb分类_Click()
    Dim lobjRec As Object
    Dim i As Integer
    If ccmb分类.Text = "" Then clst项目.Clear: Exit Sub
    '获取自觉症状项目
    dasubSetQueryTimeout 6000
    Set lobjRec = dafuncGetData("select 名称 from 系统管理_字典_字典内容表 where Parent in (select InnerID from 系统管理_字典_字典内容表 where 名称 = '" & ccmb分类.Text & " ')")
    clst项目.Clear
'    ccmb分类.AddItem ""
    If ccmb分类.Text = "其他" Or ccmb分类.Text = "" Then
        clst项目.Visible = False
        ctxt项目.Visible = True
    Else
        clst项目.Visible = True
        ctxt项目.Visible = False
    End If
    If (lobjRec.EOF Or lobjRec.BOF) Then
        clst项目.Clear
        clst项目.Visible = False
        ctxt项目.Visible = True
        Exit Sub
    End If
    For i = 1 To lobjRec.RecordCount
        clst项目.AddItem lobjRec("名称")
        clst项目.ItemData(clst项目.NewIndex) = i
        lobjRec.MoveNext
    Next
    clst项目.ListIndex = 0
    
End Sub

'针对未婚体检人员，一些信息不用输入
Private Sub Ccmb婚否_Click(Index As Integer)
    On Error GoTo errHandler
    '修改人：张令 2012.12.06
    'bug号：0000045
    '说明：修改if条件，加上为空时。添加ctxtmatehelh和ctxtmatejob清空。 ↓↓
    If Trim(Ccmb婚否(Index).Text) = "未婚" Or Trim(Ccmb婚否(Index).Text) = "" Then
        If Index <> 2 Then
            ctxtmatehelh(Index).Enabled = False
            ctxtmarrydate(Index).Enabled = False
            ctxtmatejob(Index).Enabled = False
            ctxtmateradioac(Index).Enabled = False
            ctxtmatehelh(Index).Text = ""
            ctxtmatejob(Index).Text = ""
            ctxtmateradioac(Index).Text = ""
            ctxtmarrydate(Index).Text = ""
            subclear生育史
        Else
            ctxt现有子女(Index).Text = ""
            ctxt早产(Index).Text = ""
            ctxt流产(Index).Text = ""
            ctxt死产(Index).Text = ""
            ctxt异常胎.Text = ""
        End If
        Frame3(Index).Enabled = False
    '2012.12.06    ↑↑
    Else
        If Index <> 2 Then
            ctxtmatehelh(Index).Enabled = True
            ctxtmarrydate(Index).Enabled = True
            ctxtmatejob(Index).Enabled = True
            ctxtmateradioac(Index).Enabled = True
        End If
        Frame3(Index).Enabled = True
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "职业病史录入", "frmcareerhstregt", "ccmb婚否_click", Err.Number, Err.Description, True
End Sub




'删除   职业史
Private Sub Ccmd删除_Click()
    Dim introw As String
    Dim i As Integer
    On Error GoTo errHandler
    If cgrd职业史.Row = 0 Or cgrd职业史.Row > cgrd职业史.Rows - 1 Then
        MsgBox "请选择要删除的信息！", vbInformation, "系统提示"
        Exit Sub
    End If
    introw = cgrd职业史.Row
    If introw < cgrd职业史.Rows - 1 Then
        For i = introw + 1 To cgrd职业史.Rows - 1
        cgrd职业史.Cell(flexcpText, i, mcolIndex("编号")) = i - 1
        Next
    End If
    cgrd职业史.RemoveItem cgrd职业史.Row
    cgrd职业史.AutoSize 0, cgrd职业史.Cols - 1
    mintrow = 0
    Exit Sub
errHandler:
    sfsub错误处理 "职业病史录入", "frmcareerhstregt", "ccmb删除_click", Err.Number, Err.Description, True
End Sub
'删除  职业病史
Private Sub Ccmdcancel病史_Click()
    Dim zyintrow As String
    Dim i As Integer
    On Error GoTo errHandler
    If cgrd病史.Row = 0 Or cgrd病史.Row > cgrd病史.Rows - 1 Then
        MsgBox "请选择要删除的信息！", vbInformation, "系统提示"
        Exit Sub
    End If
    zyintrow = cgrd病史.Row
    If zyintrow < cgrd病史.Rows - 1 Then
        For i = zyintrow + 1 To cgrd病史.Rows - 1
        cgrd病史.Cell(flexcpText, i, mcolIndex("编号")) = i - 1
        Next
    End If
    cgrd病史.RemoveItem cgrd病史.Row
    cgrd病史.AutoSize 0, cgrd病史.Cols - 1
    '将mintrow替换成jmintrow 既往病史单独处理  modify by lanchao 2015-9-16
'    mintrow = 0
     jmintrow = 0
    Exit Sub
errHandler:
    sfsub错误处理 "职业病史录入", "frmcareerhstregt", "ccmdcancle病史_click", Err.Number, Err.Description, True
End Sub

'删除  自觉症状
Private Sub ccmdcancel症状_Click()
    Dim zjintrow As String
    Dim i As Integer
    On Error GoTo errHandler
    If cgrd症状.Row = 0 Or cgrd症状.Row > cgrd症状.Rows - 1 Then
        MsgBox "请选择要删除的信息！", vbInformation, "系统提示"
        Exit Sub
    End If
    zjintrow = cgrd症状.Row
    If zjintrow < cgrd症状.Rows - 1 Then
        For i = zjintrow + 1 To cgrd症状.Rows - 1
        cgrd症状.Cell(flexcpText, i, mcolIndex("编号")) = i - 1
        Next
    End If
    cgrd症状.RemoveItem cgrd症状.Row
    cgrd症状.AutoSize 0, cgrd症状.Cols - 1
    zjintrow = 0
    Exit Sub
errHandler:
    sfsub错误处理 "职业病史录入", "frmcareerhstregt", "ccmdcancel症状_click", Err.Number, Err.Description, True
End Sub

'系统编号|职业史编号|工作单位|部门|工种|职业危害种类|接触时间|防护措施|备注|放射线种类|每日工作量|累积照射量|过量放射史|起始时间｜结束时间|是否放射性
'确定  职业史
Private Sub ccmdOk_Click()
    Dim i As Integer
    On Error GoTo errHandler
    '判断是否是修改信息，因为修改不会增加新记录
    If mintrow = 0 Then
        cgrd职业史.Rows = cgrd职业史.Rows + 1
        i = cgrd职业史.Rows - 1
    Else
        i = mintrow
    End If
    cgrd职业史.Cell(flexcpText, i, mcolIndex("编号")) = i
    cgrd职业史.Cell(flexcpText, i, mcolIndex("工作单位")) = Trim(ctxt单位.Text)
    cgrd职业史.Cell(flexcpText, i, mcolIndex("部门")) = Trim(ctxt部门.Text)
    cgrd职业史.Cell(flexcpText, i, mcolIndex("工种")) = Trim(ctxt工种.Text)
    '将备注修改为接触时间显示 modify by lanchao 2015-9-6
    cgrd职业史.Cell(flexcpText, i, mcolIndex("备注")) = Trim(ctxt备注.Text)
    '日期格式不要 2015-6-26 by lanchao
    cgrd职业史.Cell(flexcpText, i, mcolIndex("起始时间")) = Trim(ctxt起始.Text)
    cgrd职业史.Cell(flexcpText, i, mcolIndex("结束时间")) = Trim(ctxt结束.Text)
    cgrd职业史.Cell(flexcpText, i, mcolIndex("危害种类")) = Trim(ctxtweihai.Text)
    cgrd职业史.Cell(flexcpText, i, mcolIndex("防护措施")) = Trim(ctxt防护.Text)

    cgrd职业史.Cell(flexcpText, i, mcolIndex("放射种类")) = Trim(ctxtfangshe.Text)
    cgrd职业史.Cell(flexcpText, i, mcolIndex("每日工作量")) = Trim(ctxt工作量.Text)
    cgrd职业史.Cell(flexcpText, i, mcolIndex("累积照射量")) = Trim(ctxt照射量.Text)
    cgrd职业史.Cell(flexcpText, i, mcolIndex("过量照射史")) = Trim(ctxt过量照射.Text)
    '接触时间重新打开 2015-9-6 by lanchao
    cgrd职业史.Cell(flexcpText, i, mcolIndex("接触时间")) = Trim(ctxt接触.Text)
    '判断是否放射类
    If Chk放射.Value = 1 Then
        cgrd职业史.Cell(flexcpText, i, mcolIndex("是否放射性")) = "是"
    Else
        cgrd职业史.Cell(flexcpText, i, mcolIndex("是否放射性")) = "否"
    End If
    '判断确定后是否清空
    If Chkokclear.Value = 1 Then
        Call subokclear
    End If
    cgrd职业史.AutoSize 0, cgrd职业史.Cols - 1
    mintrow = 0
    Exit Sub
errHandler:
    sfsub错误处理 "职业病史录入", "frmcareerhstregt", "ccmdok_click", Err.Number, Err.Description, True
End Sub

'确定   既往病史
Private Sub ccmdok病史_Click()
    Dim i As Integer
    On Error GoTo errHandler
    '判断是否是修改信息，因为修改不会增加新记录
    '将mintrow修改为jmintrow进行单独处理 modify by lanchao 2015-9-16
    If jmintrow = 0 Then
        cgrd病史.Rows = cgrd病史.Rows + 1
        i = cgrd病史.Rows - 1
    Else
        i = jmintrow
    End If
    cgrd病史.Cell(flexcpText, i, mcolIndexwkdis("编号")) = i
    cgrd病史.Cell(flexcpText, i, mcolIndexwkdis("疾病名称")) = Trim(ctxt疾病名称.Text)
    cgrd病史.Cell(flexcpText, i, mcolIndexwkdis("诊断单位")) = Trim(ctxt诊断单位.Text)
    cgrd病史.Cell(flexcpText, i, mcolIndexwkdis("诊断日期")) = Trim(ctxt诊断日期.Text)
    cgrd病史.Cell(flexcpText, i, mcolIndexwkdis("转归")) = Trim(ctxt转归.Text)
    cgrd病史.Cell(flexcpText, i, mcolIndexwkdis("治疗经过")) = Trim(ctxt治疗经过.Text)
    If Chkokclear病史.Value = 1 Then
        subokclear病史
    End If
    jmintrow = 0
    cgrd病史.AutoSize 0, cgrd病史.Cols - 1
    Exit Sub
errHandler:
    sfsub错误处理 "职业病史录入", "frmcareerhstregt", "ccmdok病史_click", Err.Number, Err.Description, True
End Sub

'确定   自觉症状
Private Sub Ccmdok症状_Click()
    Dim i As Integer, j As Integer
    On Error GoTo errHandler
    '判断是否是修改信息，因为修改不会增加新记录
    Dim zjmintrow As Integer
    For j = 0 To clst项目.ListCount - 1
        If clst项目.Selected(j) = True Then
            If zjmintrow = 0 Then
                cgrd症状.Rows = cgrd症状.Rows + 1
                i = cgrd症状.Rows - 1
            Else
                i = zjmintrow
            End If
            cgrd症状.Cell(flexcpText, i, mcolindexzz("编号")) = i
            cgrd症状.Cell(flexcpText, i, mcolindexzz("症状")) = Trim(clst项目.List(j))
'            clst项目.RemoveItem (j)
'            cgrd症状.Cell(flexcpText, i, mcolindexzz("出现时间")) = Trim(ctxt出现时间.Text)   ' 2015-11-27 by 牟俊
            cgrd症状.Cell(flexcpText, i, mcolindexzz("程度")) = Trim(ctxt程度.Text)
        End If
    Next
    
    If ccmb分类.Text = "其他" Or clst项目.ListCount - 1 < 0 Then
        If zjmintrow = 0 Then
            cgrd症状.Rows = cgrd症状.Rows + 1
            i = cgrd症状.Rows - 1
        Else
            i = zjmintrow
        End If
        cgrd症状.Cell(flexcpText, i, mcolindexzz("编号")) = i
        cgrd症状.Cell(flexcpText, i, mcolindexzz("症状")) = Trim(ctxt项目.Text)
'        cgrd症状.Cell(flexcpText, i, mcolindexzz("出现时间")) = Trim(ctxt出现时间.Text)        ' 2015-11-27 by 牟俊
        cgrd症状.Cell(flexcpText, i, mcolindexzz("程度")) = Trim(ctxt程度.Text)
    End If
    
    ccmb分类_Click
    If Chkokclear症状.Value = 1 Then
        subokclear症状
    End If
    cgrd症状.AutoSize 0, cgrd症状.Cols - 1
    zjmintrow = 0
    Exit Sub
errHandler:
    sfsub错误处理 "职业病史录入", "frmcareerhstregt", "ccmdok症状_click", Err.Number, Err.Description, True
End Sub

'修改   职业史
Private Sub Ccmd修改_Click()
    cgrd职业史_DblClick
End Sub

Private Sub cgrdzzxw_Click()
' MsgBox (cgrdzzxw.TextMatrix(cgrdzzxw.RowSel, 0))
selectsysno = cgrdzzxw.TextMatrix(cgrdzzxw.RowSel, 0)
selectzz = cgrdzzxw.TextMatrix(cgrdzzxw.RowSel, 1)
selectcd = cgrdzzxw.TextMatrix(cgrdzzxw.RowSel, 2)
selectcxrq = cgrdzzxw.TextMatrix(cgrdzzxw.RowSel, 3)
End Sub




Private Sub Chk放射_Click()
    If Chk放射.Value = 1 Then
        Frame11.Enabled = True
    Else
        Frame11.Enabled = False
    End If
End Sub

'修改  职业病史
Private Sub Cmmdmody病史_Click()
    cgrd病史_DblClick
End Sub

'修改   自觉症状
'Private Sub ccmbmody症状_Click()
'    cgrd症状_DblClick
'End Sub

'双击grid 修改  职业病史
Private Sub cgrd病史_DblClick()
    On Error GoTo errHandler
    If cgrd病史.Row = 0 Or cgrd病史.Row > cgrd病史.Rows - 1 Then
        MsgBox "请选择要修改的信息！", vbInformation, "系统提示"
        Exit Sub
    End If
    '将mintrow修改为jmintrow进行单独处理 modify by lanchao 2015-9-16
    jmintrow = cgrd病史.Cell(flexcpText, cgrd病史.Row, mcolIndex("编号"))
    ctxt疾病名称.Text = cgrd病史.Cell(flexcpText, cgrd病史.Row, mcolIndexwkdis("疾病名称"))
    ctxt诊断单位.Text = cgrd病史.Cell(flexcpText, cgrd病史.Row, mcolIndexwkdis("诊断单位"))
    ctxt诊断日期.Text = cgrd病史.Cell(flexcpText, cgrd病史.Row, mcolIndexwkdis("诊断日期"))
    ctxt转归.Text = cgrd病史.Cell(flexcpText, cgrd病史.Row, mcolIndexwkdis("转归"))
    ctxt治疗经过.Text = cgrd病史.Cell(flexcpText, cgrd病史.Row, mcolIndexwkdis("治疗经过"))
    Exit Sub
errHandler:
    sfsub错误处理 "职业病史录入", "frmcareerhstregt", "cgrd病史_dbclick", Err.Number, Err.Description, True
End Sub

'双击修改,填充文本框内信息   职业史
Private Sub cgrd职业史_DblClick()
    On Error GoTo errHandler
    If cgrd职业史.Row = 0 Or cgrd职业史.Row > cgrd职业史.Rows - 1 Then
        MsgBox "请选择要修改的信息！", vbInformation, "系统提示"
        Exit Sub
    End If
    mintrow = cgrd职业史.Cell(flexcpText, cgrd职业史.Row, mcolIndex("编号"))
    ctxt单位.Text = cgrd职业史.Cell(flexcpText, cgrd职业史.Row, mcolIndex("工作单位"))
    ctxt部门.Text = cgrd职业史.Cell(flexcpText, cgrd职业史.Row, mcolIndex("部门"))
    ctxt工种.Text = cgrd职业史.Cell(flexcpText, cgrd职业史.Row, mcolIndex("工种"))
    '将备注修改为接触时间显示 modify by lanchao 2015-9-6
    ctxt备注.Text = cgrd职业史.Cell(flexcpText, cgrd职业史.Row, mcolIndex("备注"))
    ctxt起始.Text = cgrd职业史.Cell(flexcpText, cgrd职业史.Row, mcolIndex("起始时间"))
    ctxt结束.Text = cgrd职业史.Cell(flexcpText, cgrd职业史.Row, mcolIndex("结束时间"))
    ctxt危害种类.Text = cgrd职业史.Cell(flexcpText, cgrd职业史.Row, mcolIndex("危害种类"))
    ctxt防护.Text = cgrd职业史.Cell(flexcpText, cgrd职业史.Row, mcolIndex("防护措施"))
    
    ctxtfangshe.Text = cgrd职业史.Cell(flexcpText, cgrd职业史.Row, mcolIndex("放射种类"))
    ctxt工作量.Text = cgrd职业史.Cell(flexcpText, cgrd职业史.Row, mcolIndex("每日工作量"))
    ctxt照射量.Text = cgrd职业史.Cell(flexcpText, cgrd职业史.Row, mcolIndex("累积照射量"))
    ctxt过量照射.Text = cgrd职业史.Cell(flexcpText, cgrd职业史.Row, mcolIndex("过量照射史"))
    '接触时间重新打开 modify by lanchao 2015-9-6
    ctxt接触.Text = cgrd职业史.Cell(flexcpText, cgrd职业史.Row, mcolIndex("接触时间"))
    Dim tjblable As String
    tjblable = Label体检类型.Caption
    
    Dim tmpstr As String
    tmpstr = cgrd职业史.Cell(flexcpText, cgrd职业史.Row, mcolIndex("是否放射性"))
    '初始化界面时，放射工作信息录入框disable modify by lanchao 2015-8-17
    If tmpstr = "是" Then
        Chk放射.Value = 1
        Frame11.Enabled = True
    Else
        Chk放射.Value = 0
        Frame11.Enabled = False
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "职业病史录入", "frmcareerhstregt", "cgrd职业史_dblclick", Err.Number, Err.Description, True
End Sub

Private Sub Combo1_Click()
ctxt照射量.Text = Combo1.Text
Combo1.Visible = False
Combo1.Text = "请选择"
End Sub

Private Sub Combo10_Click()
ctxt子女健康(1).Text = Combo10.Text
End Sub

Private Sub Combo2_Click()
If ctxt过敏史(2).Text = "" Then
ctxt过敏史(2).Text = Combo2.Text
Else
ctxt过敏史(2).Text = ctxt过敏史(2).Text + "，" + Combo2.Text
End If
End Sub


Private Sub Combo3_Click()
Dim i As Integer
ctxt疾病名称.Text = Combo3.Text
'2015-4-8 刘伟 添加
'If Label体检类型 = "8023部队" And Combo3.Text <> "" Then
' Dim lobjRec As Object
'    Set lobjRec = pobjDict.FetchEx("职业病" + Combo3.Text)
'    Combo5.Clear
'    Combo5.AddItem ""
'    For i = 1 To lobjRec.RecordCount
'        Combo5.AddItem lobjRec("名称")
'        Combo5.ItemData(Combo5.NewIndex) = lobjRec("编号")
'        lobjRec.MoveNext
'    Next
'   Combo5.ListIndex = 0
'   End If
End Sub

Private Sub Combo4_Click()
If ctxt防护.Text = "" Then
ctxt防护.Text = Combo4.Text
Else
ctxt防护.Text = ctxt防护.Text + "，" + Combo4.Text
End If
End Sub

Private Sub Combo5_Click()
If Combo5.ListIndex = 0 Then
ctxt治疗经过.Text = ""
Else
ctxt治疗经过.Text = Combo5.Text
ctxt治疗经过.SetFocus
ctxt治疗经过.SelStart = Len(ctxt治疗经过.Text)


End If
End Sub

Private Sub Combo6_Click()
ctxt转归.Text = Combo6.Text

End Sub

Private Sub Combo7_Click()
ctxt不孕不育(0).Text = Combo7.Text
End Sub

Private Sub Combo8_Click()
ctxt子女健康(0).Text = Combo8.Text
End Sub
'配偶健康情况可选 2015-6-26 by lanchao
Private Sub Combo11_Click()
ctxtmatehelh(0).Text = Combo11.Text
End Sub

Private Sub Combo9_Click()
ctxt不孕不育(1).Text = Combo9.Text
End Sub


Private Sub Comd填写_Click()
    Dim A, B, C As String
    A = Trim(Text临时.Text)
    '体重
    B = Trim(Split(A, "H")(0))
    C = Trim(Mid(B, 3, Len(B) - 2))
    If Left(C, 1) = "0" Then
        C = Right(C, Len(C) - 1)
    End If
    ctxt体重.Text = C
    '身高
    B = Trim(Split(A, "H")(1))
    C = Trim(Mid(B, 2, Len(B) - 1))
    ctxt身高.Text = C
    '体重指数
    Dim W, H, X As Single
    Dim Z As String
    W = Val(ctxt体重.Text)
    H = Val(ctxt身高.Text)  '单位cm
    H = H / 100           '单位转成m
    X = W / (H * H)
    X = Format(X, "0.0")
    Z = Str(X)
    ctxt体重指数.Text = Z
    '体形
    If X < 18.5 Then
        ctxt体形.Text = "过轻"
    ElseIf X >= 18.5 And X < 24 Then
        ctxt体形.Text = "正常"
    ElseIf X >= 24 And X < 27.5 Then
        ctxt体形.Text = "过重"
    ElseIf X >= 27.5 And X < 30 Then
        ctxt体形.Text = "轻度肥胖"
    ElseIf X >= 30 And X < 35 Then
        ctxt体形.Text = "中度肥胖"
    ElseIf X >= 35 Then
        ctxt体形.Text = "重度肥胖"
    End If
End Sub

Private Sub Command1_Click()

If cgrdzzxw.Row > 67 Then

MsgBox "已经添加保存成功，只能修改项目！"

Else
frm症状询问.Show vbModal
End If

End Sub

Private Sub Command2_Click()
If cgrdzzxw.Row > 0 Then

frm症状修改.Show vbModal


Else
 MsgBox "请选择需要修改的症状项目！"
End If
End Sub

Private Sub Command3_Click()
Dim conclusion As String
conclusion = "正常"
If ctxtxinlv.Text <> "" And IsNumeric(ctxtxinlv.Text) Then

If CInt(ctxtxinlv.Text) > 100 Then
conclusion = "过速"
End If
If CInt(ctxtxinlv.Text) < 60 Then
conclusion = "过缓"
End If

dafuncGetData ("update 职业病体检_结果信息_内科 set 体检结果='" & ctxtxinlv.Text & "'  , 体检医师='" & um用户编号 & "',单项结论='" & conclusion & "' where 体检项目='02002' and 系统编号='" & Trim(ctxtsysno.Text) & "'")
'dafuncGetData ("update 职业病体检_结果信息_内科 set 体检结果='" & ctxtxinlv.Text & "'  , 体检医师='" & um用户编号 & "' ,填写时间='" & Now & "',单项结论='" & conclusion & "' where 体检项目='02002' and 系统编号='" & Trim(ctxtsysno.Text) & "'")

Label66.Caption = "已保存。"
ctxtxinlv.Text = ""

End If


End Sub

Private Sub Command4_Click()
Dim i As Integer
    On Error GoTo errHandler
    '判断是否是修改信息，因为修改不会增加新记录
    If mintrow = 0 Then
        cgrd职业史.Rows = cgrd职业史.Rows + 1
        i = cgrd职业史.Rows - 1
    Else
        i = mintrow
    End If
    cgrd职业史.Cell(flexcpText, i, mcolIndex("编号")) = i
    cgrd职业史.Cell(flexcpText, i, mcolIndex("工作单位")) = Trim(ctxt单位.Text)
    cgrd职业史.Cell(flexcpText, i, mcolIndex("部门")) = Trim(ctxt部门.Text)
    cgrd职业史.Cell(flexcpText, i, mcolIndex("工种")) = Trim(ctxt工种.Text)
    cgrd职业史.Cell(flexcpText, i, mcolIndex("备注")) = Trim(ctxt备注.Text)
    '日期格式不要 2015-6-26 by lanchao 不要接触时间
    cgrd职业史.Cell(flexcpText, i, mcolIndex("起始时间")) = Trim(ctxt起始.Text)
    cgrd职业史.Cell(flexcpText, i, mcolIndex("结束时间")) = Trim(ctxt结束.Text)
    cgrd职业史.Cell(flexcpText, i, mcolIndex("危害种类")) = Trim(ctxtweihai.Text)
    cgrd职业史.Cell(flexcpText, i, mcolIndex("防护措施")) = Trim(ctxt防护.Text)
    '接触时间重新显示 2015-9-6 by lanchao
    cgrd职业史.Cell(flexcpText, i, mcolIndex("接触时间")) = Trim(ctxt接触.Text)
    cgrd职业史.Cell(flexcpText, i, mcolIndex("放射种类")) = Trim(ctxtfangshe.Text)
    cgrd职业史.Cell(flexcpText, i, mcolIndex("每日工作量")) = Trim(ctxt工作量.Text)
    cgrd职业史.Cell(flexcpText, i, mcolIndex("累积照射量")) = Trim(ctxt照射量.Text)
    cgrd职业史.Cell(flexcpText, i, mcolIndex("过量照射史")) = Trim(ctxt过量照射.Text)
    '判断是否放射类
    If Chk放射.Value = 1 Then
        cgrd职业史.Cell(flexcpText, i, mcolIndex("是否放射性")) = "是"
    Else
        cgrd职业史.Cell(flexcpText, i, mcolIndex("是否放射性")) = "否"
    End If
    '判断确定后是否清空
    If Chkokclear.Value = 1 Then
        Call subokclear
    End If
    cgrd职业史.AutoSize 0, cgrd职业史.Cols - 1
    mintrow = 0
    Exit Sub
errHandler:
    sfsub错误处理 "职业病史录入", "frmcareerhstregt", "ccmdok_click", Err.Number, Err.Description, True
   
End Sub

'双击grid 修改  自觉症状
'Private Sub cgrd症状_DblClick()
'    On Error GoTo errHandler
'    If cgrd症状.Row = 0 Or cgrd症状.Row > cgrd症状.Rows - 1 Then
'        MsgBox "请选择要修改的信息！", vbInformation, "系统提示"
'        Exit Sub
'    End If
'    mintrow = cgrd症状.Cell(flexcpText, cgrd症状.Row, mcolIndex("编号"))
''    clst项目.AddItem cgrd症状.Cell(flexcpText, cgrd症状.Row, mcolindexzz("症状"))
'    ctxt出现时间.Text = cgrd症状.Cell(flexcpText, cgrd症状.Row, mcolindexzz("出现时间"))
'    ctxt程度.Text = cgrd症状.Cell(flexcpText, cgrd症状.Row, mcolindexzz("程度"))
'    Exit Sub
'errHandler:
'    sfsub错误处理 "职业病史录入", "frmcareerhstregt", "cgrd症状_dblclick", Err.Number, Err.Description, True
'End Sub
'不结婚日期不做格式判断 2015-6-26 by lanchao
Private Sub ctxtmarrydate_LostFocus(Index As Integer)
'    If ctxtmarrydate(mIndex).Enabled Then
'        If DateDiff("d", ctxtmarrydate(mIndex).Value, Date) < 0 Then
'            MsgBox "结婚日期有误！"
'            ctxtmarrydate(mIndex).Value = Date
'        End If
'    End If
End Sub

'当检测到有回车键按下后，移出焦点
Private Sub ctxtsysno_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        Ccmb婚否(mIndex).SetFocus
    End If
End Sub

'当系统编号文本框移走焦点后，填充体检人员个人基本信息
Private Sub ctxtsysno_LostFocus()
    Dim lobjRec As Object
    Dim lobjRegt As Object
    Dim lobjlife As Object
    Dim registeragn As String
    On Error GoTo errHandler
    If Len(RTrim(ctxtsysno.Text)) < 5 Then
        MsgBox "系统编号错误，请检查！"
        ctxtsysno.Text = ""
        gatherclear   '出现人员错误或者找不到或者已经录入情况时清空界面  2016-3-2 by 牟俊
        ctxtsysno.SetFocus
        Exit Sub
    End If
' 将清空屏蔽  2016-3-2 by 牟俊
'    '2012-04-14 于登淼 ↓
'    '可能会填入新的体检人员信息，此时需将4个病史的窗体内容全部清空
'    subclear
'    subokclear
'    subokclear病史
'    subokclear症状
'    subClear体格一般情况
'    '2012-04-14 于登淼 ↑
    
    '创建生活史对象
    Set lobjPsLifeHst = CreateObject("职业病史录入.clslifehstregt")
    lobjPsLifeHst.系统编号 = Trim(ctxtsysno.Text)
    If Not lobjPsLifeHst.tmp已登记 Then
        ctxtsysno.SetFocus
        Exit Sub
    End If
    
    Lab姓名.Caption = lobjPsLifeHst.姓名
    Lab性别.Caption = lobjPsLifeHst.性别
    Lab年龄.Caption = lobjPsLifeHst.年龄
    lab单位.Caption = lobjPsLifeHst.单位名称
    Lab现工种.Caption = lobjPsLifeHst.现工种
    Lab现职务.Caption = lobjPsLifeHst.职务
    Label危害因素.Caption = lobjPsLifeHst.危害因素
    '2012-07-11 于登淼 ↓
    '将已知信息全部添加完整
    ctxt单位.Text = lobjPsLifeHst.单位名称
    ctxt工种.Text = lobjPsLifeHst.现工种
    '2012-07-11 于登淼 ↑
    
    '获取像片
    Set lobjRec = CreateObject("职业病对象.clspersonexamed")
    lobjRec.系统编号 = Trim(ctxtsysno.Text)
    Picture2.Picture = lobjRec.像片
    Picture2.Visible = True
    
    '2012-06-14 于登淼 ↓
    subClear体格一般情况
    subLoad体格一般情况 Trim(ctxtsysno.Text)
    '2012-06-14 于登淼 ↑
    
    '判断是否已进行过 职业病史录入 操作
    Set lobjRegt = CreateObject("职业病史录入.clscareerhstmage")
    lobjRegt.系统编号 = Trim(ctxtsysno.Text)
    registeragn = lobjRegt.已记录标志
    If registeragn = "100" Or registeragn = "101" Then
        MsgBox "该体检人员不存在或已体检完毕！", vbExclamation, "系统提示"
        subclearps
        gatherclear  '出现人员错误或者找不到或者已经录入情况时清空界面  2016-3-2 by 牟俊
        ctxtsysno.SetFocus
        Exit Sub
    End If
    If registeragn = "2" And 访问记号 = 0 Then
        If MsgBox("该条码已进行过受检者个人信息登记，要修改它？", vbYesNo + vbQuestion + vbDefaultButton2, "系统提示") = vbNo Then
            subclearps
             gatherclear  '出现人员错误或者找不到或者已经录入情况时清空界面  2016-3-2 by 牟俊
            ctxtsysno.SetFocus
            Exit Sub
        End If
    End If
    If 访问记号 = 1 Or registeragn = "2" Then
        访问记号 = 1
        '修改生活史
        Set lobjlife = lobjRegt.个人生活史
        sub修改生活史 lobjlife
        
        'sub修改职业史
        Set lobjlife = lobjRegt.职业史
'        cgrd职业史.Clear
        '判断数据库里是否有对应的数据，如果有才显示(下面的病史和自觉症状一样的处理)  2016-3-7 by 牟俊
        If lobjlife.RecordCount > 0 Then
        Set cgrd职业史.DataSource = lobjlife
'        Set cgrd职业史.DataSource = lobjlife
        'cgrd职业史.ColHidden(mcolIndex("系统编号")) = True
        cgrd职业史.ColHidden(0) = True
        '职业健康才显示接触时间
         If Label体检类型.Caption <> "职业健康" Then
         cgrd职业史.ColHidden(10) = True
         cgrd职业史.AutoSize 0, cgrd职业史.Cols - 1
         End If
        End If
        'sub修改病史
        Set lobjlife = lobjRegt.既往病史
'        Set cgrd病史.DataSource = lobjlife
        If lobjlife.RecordCount > 0 Then
         Set cgrd病史.DataSource = lobjlife
        'cgrd病史.ColHidden(mcolIndexwkdis("系统编号")) = True
        cgrd病史.ColHidden(0) = True
        cgrd病史.AutoSize 0, cgrd病史.Cols - 1
        End If
        'sub修改自觉症状
        Set lobjlife = lobjRegt.自觉症状
'        Set cgrd症状.DataSource = lobjlife
        If lobjlife.RecordCount > 0 Then
        Set cgrd症状.DataSource = lobjlife
        'cgrd症状.ColHidden(mcolindexzz("系统编号")) = True
        cgrd症状.ColHidden(0) = True
        cgrd症状.AutoSize 0, cgrd症状.Cols - 1
        End If
    Else
        访问记号 = 0
    End If
'    If Ccmb婚否(mIndex).Text = "已婚" Or Ccmb婚否(mIndex).Text = "已婚" Then
'        If mIndex <> 2 Then
'            ctxtmatehelh(mIndex).Enabled = True
'            ctxtmarrydate(mIndex).Enabled = True
'            ctxtmatejob(mIndex).Enabled = True
'            ctxtmateradioac(mIndex).Enabled = True
'        End If
'        Frame3(mIndex).Enabled = True
'    End If
    '判断性别
    If Trim(Lab性别.Caption) = "男" Then
        Frame1.Enabled = False
    Else
        Frame1.Enabled = True
    End If
    Set lobjRec = Nothing
    
'    MsgBox "移走焦点过程" '2016-3-1 by 牟俊
    Exit Sub
errHandler:
    sfsub错误处理 "职业病史录入", "frmcareerhstregt", "ctxtsysno_lostfocus", Err.Number, Err.Description, True
End Sub

'2015-11-27 by 牟俊

'Private Sub ctxt出现时间_LostFocus()
'
'    '2015-03-30 liuwei
'    'If DateDiff("d", ctxt出现时间.Text, Date) < 0 Then
'     '   MsgBox "出现时间错误！"
'     '   ctxt出现时间.Text = Date
'    'End If
'    '注释下面代码 刘伟2015-4-7
'    'If ctxt出现时间.Text = "" Then
'     'ctxt出现时间.Text = Format(Now, "yyyy年mm月")
'    'End If
'End Sub



Private Sub ctxt放射种类_Click()
If ctxtfangshe.Text = "" Then
ctxtfangshe.Text = ctxt放射种类.Text
Else
ctxtfangshe.Text = ctxtfangshe.Text + "，" + ctxt放射种类.Text
End If
End Sub





Private Sub ctxt接触时间_LostFocus()
    If DateDiff("d", ctxt接触时间.Value, Date) < 0 Then
        MsgBox "接触时间错误！"
        ctxt接触时间.Value = Date
    End If
End Sub

Private Sub ctxt结束_LostFocus()
'    If DateDiff("d", ctxt结束.Value, ctxt起始.Value) > 0 Then
'        MsgBox "结束时间小于起始时间！"
'        ctxt结束.Value = Date
'    End If
End Sub

Private Sub ctxt戒烟_LostFocus(Index As Integer)
    If Trim(ctxt戒烟(Index).Text) = "" Then Exit Sub
    If Not IsNumeric(ctxt戒烟(Index).Text) Then
        MsgBox "戒烟值不对！"
        ctxt戒烟(Index).Text = ""
        ctxt戒烟(Index).SetFocus
    Else
        If Int(Val(ctxt戒烟(Index).Text)) > Int(Val(Lab年龄)) Then
            MsgBox "戒烟值不对！"
            ctxt戒烟(Index).Text = ""
            ctxt戒烟(Index).SetFocus
        End If
    End If
End Sub

Private Sub ctxt酒龄_LostFocus(Index As Integer)
    If Trim(ctxt酒龄(Index).Text) = "" Then Exit Sub
    If Not IsNumeric(ctxt酒龄(Index).Text) Then
        MsgBox "酒龄值不对！"
        ctxt酒龄(Index).Text = ""
        ctxt酒龄(Index).SetFocus
    Else
        If Int(Val(ctxt酒龄(Index).Text)) > Int(Val(Lab年龄)) Then
            MsgBox "酒龄值不对！"
            ctxt酒龄(Index).Text = ""
            ctxt酒龄(Index).SetFocus
        End If
    End If
End Sub

Private Sub ctxt起始_LostFocus()
'日期不做判断 6-26 by lanchao
'    If DateDiff("d", ctxt起始.Value, Date) < 0 Then
'        MsgBox "起始时间错误！"
'        '修改人：罗李奎 2012-12-4 ↓
'        '说明：当起始时间大于了当前时间时提示后，时间还原到当前时间
'        'bug号：0000081
'        ctxt起始.Value = Date
'        '罗李奎 2012-12-4 ↑
'    End If
End Sub

Private Sub ctxt身高_LostFocus()
    If Trim(ctxt身高.Text) = "" Then Exit Sub
    If Not IsNumeric(ctxt身高.Text) Then
        MsgBox ("身高须输入数字")
        ctxt身高.Text = ""
        ctxt身高.SetFocus
        Exit Sub
    Else
        If Int(Val(ctxt身高.Text)) > 280 Then
            MsgBox ("身高值过大")
            ctxt身高.Text = ""
            ctxt身高.SetFocus
            Exit Sub
        ElseIf Int(Val(ctxt身高.Text)) < 60 Then
            MsgBox ("身高值过小")
            ctxt身高.Text = ""
            ctxt身高.SetFocus
            Exit Sub
        End If
    End If
    
End Sub

Private Sub ctxt收缩压_LostFocus()
    If Trim(ctxt收缩压.Text) = "" Then Exit Sub
    If Not IsNumeric(ctxt收缩压.Text) Then
        MsgBox ("收缩压须输入数字")
        ctxt收缩压.Text = ""
        ctxt收缩压.SetFocus
        Exit Sub
    Else
        If Int(Val(ctxt收缩压.Text)) > 230 Then
            MsgBox ("收缩压值过大")
            ctxt收缩压.Text = ""
            ctxt收缩压.SetFocus
            Exit Sub
        ElseIf Int(Val(ctxt收缩压.Text)) < 60 Then
            MsgBox ("收缩压值过小")
            ctxt收缩压.Text = ""
            ctxt收缩压.SetFocus
            Exit Sub
        End If
    End If
End Sub

Private Sub ctxt舒张压_LostFocus()
    If Len(ctxtsysno.Text) = 0 Then Exit Sub    '默认舒张压是最后一个被清空内容的。
    If Trim(ctxt舒张压.Text) = "" Then Exit Sub
    If Not IsNumeric(ctxt舒张压.Text) Then
        MsgBox ("舒张压须输入数字")
        ctxt舒张压.Text = ""
        ctxt舒张压.SetFocus
        Exit Sub
    Else
        If Int(Val(ctxt舒张压.Text)) > 180 Then
            MsgBox ("舒张压值过大")
            ctxt舒张压.Text = ""
            ctxt舒张压.SetFocus
            Exit Sub
        ElseIf Int(Val(ctxt舒张压.Text)) < 30 Then
            MsgBox ("舒张压值过小")
            ctxt舒张压.Text = ""
            ctxt舒张压.SetFocus
            Exit Sub
        End If
    End If
End Sub

Private Sub ctxt体重_LostFocus()
    If Trim(ctxt体重.Text) = "" Then Exit Sub
    If Not IsNumeric(ctxt体重.Text) Then
        MsgBox ("体重须输入数字")
        ctxt体重.Text = ""
        ctxt体重.SetFocus
        Exit Sub
    Else
        If Int(Val(ctxt体重.Text)) > 500 Then
            MsgBox ("体重值过大")
            ctxt体重.Text = ""
            ctxt体重.SetFocus
            Exit Sub
        ElseIf Int(Val(ctxt体重.Text)) < 20 Then
            MsgBox ("体重值过小")
            ctxt体重.Text = ""
            ctxt体重.SetFocus
            Exit Sub
        End If
    End If
End Sub

Private Sub ctxt危害种类_Click()
If ctxtweihai.Text = "" Then
ctxtweihai.Text = ctxt危害种类.Text
Else
ctxtweihai.Text = ctxtweihai.Text + "，" + ctxt危害种类.Text
End If
End Sub

Private Sub ctxt烟龄_LostFocus(Index As Integer)
    If Trim(ctxt烟龄(mIndex).Text) = "" Then Exit Sub
    If Not IsNumeric(ctxt烟龄(mIndex).Text) Then
        MsgBox "烟龄值不对！"
        ctxt烟龄(mIndex).Text = ""
        ctxt烟龄(mIndex).SetFocus
    Else
        If Int(Val(ctxt烟龄(mIndex).Text)) > Int(Val(Lab年龄)) Then
            MsgBox "烟龄值不对！"
            ctxt烟龄(mIndex).Text = ""
            ctxt烟龄(mIndex).SetFocus
        End If
    End If
End Sub


Private Sub ctxt照射量_Click()
Combo1.Visible = True

End Sub

Private Sub ctxt诊断日期_LostFocus()
'    If DateDiff("d", ctxt诊断日期.Value, Date) < 0 Then
'        MsgBox "诊断时间错误！"
'        ctxt诊断日期.Value = Date
'    End If
End Sub

'窗体加载
Private Sub Form_Load()
    Dim lcol工具栏按钮 As New Collection '工具栏上的按钮初始化集合。
    Dim lcolInfo As Collection
    Dim lobj婚否 As Object
    Dim i As Integer
    Dim lstrSysno As String
    On Error GoTo errHandler
    
    '如果窗体已经初始化过则不再进行初始化。
    If mblninuse Then Exit Sub
    
    '设置窗体正在使用的标志。
    mblninuse = True
    bolenProject = False
    Set mobj体检 = CreateObject("职业病对象.clsMedicalExam")
     
    '初始化界面通用对象（每个界面对应一个界面通用对象实例，不可混用，切记切记）。
    Set mobjGUI = New cls界面通用对象
    
    '设置工具栏上所需要的各种按钮。
    With lcol工具栏按钮
        '.Add "去掉人员(&D)129"
        '.Add "清空人员(&R)106"
        .Add "清空(&Cl)110"
        .Add "|"
        .Add "体检项目(&T)102"
        .Add "|"
        .Add "保存"
        .Add "修改"
        .Add "|"
        '.Add "导出(&O)111"
        .Add "保存并打印(&P)107"
        .Add "|"
        '.Add "保存"
        .Add "退出"
    End With
    
    '为需要通过界面通用对象控制的各种控件赋初始值。
    With mobjGUI
        Set .Form = Me
        Set .c工具栏 = ctbMain
'        Set .c状态栏 = csbMain
    End With
    
    '调用界面通用对象提供的方法，对界面控件进行初始化。
    mobjGUI.subInitialize lcol工具栏按钮, ""
   ' If Right(pub系统编号, 1) = "F" Then
    '    lstrSysno = Right(pub系统编号, 2)
     '   lstrSysno = Left(lstrSysno, 1)
    'Else
    '    lstrSysno = Right(pub系统编号, 1)
    'End If
    '2015-03-30 刘伟  6-26 by lanchao "放射健康"<>"8023部队"
    Dim resql As Object
    Set resql = dafuncGetData("select 体检表类型 From 职业病体检_体检人员基本信息表 where 系统编号='" & pub系统编号 & "'")
    If resql("体检表类型") = "普通体检" Or resql("体检表类型") = "职业健康" Then
        freOrdinary.Visible = True
        freNuclear.Visible = False
        freRadiation.Visible = False
        '只有职业健康体检才显示 modify by lanchao 2015-9-6
        Label74.Visible = True
        ctxt接触.Visible = True
        mIndex = 2
     ElseIf resql("体检表类型") = "8023部队" Or resql("体检表类型") = "涉核部队" Or resql("体检表类型") = "涉核部队YK" Then   '增加涉核部队YK  2015-11-24 by 牟俊
'    ElseIf resql("体检表类型") = "8023部队" Or resql("体检表类型") = "涉核部队" Then
        freNuclear.Visible = True
        freOrdinary.Visible = False
        freRadiation.Visible = False
        '只有职业健康体检才显示 modify by lanchao 2015-9-6
        Label74.Visible = False
        ctxt接触.Visible = False
        mIndex = 1
    ElseIf resql("体检表类型") = "放射健康" Then
        freRadiation.Visible = True
        freOrdinary.Visible = False
        freNuclear.Visible = False
        comb发育.Visible = True         '只有放射健康才显示发育  2015-11-30 by 牟俊
        Lab体格(5).Visible = True
        '只有职业健康体检才显示 modify by lanchao 2015-9-6
        Label74.Visible = False
        ctxt接触.Visible = False
        mIndex = 0
    End If
    Label体检类型.Caption = resql("体检表类型")
    
        
    
    '加载时，显示第一个窗口
    SSTab1.Tab = 0
    
    '加载婚否组合框
    'Ccmb婚否.Clear
    'Ccmb婚否.AddItem ""
    'Set lobj婚否 = CreateObject("职业病史录入.Clslifehstregt")
   '' Set lcolInfo = lobj婚否.婚姻元素集
    'For i = 1 To lcolInfo.Count
        'Ccmb婚否.AddItem lcolInfo(i), i
        'Ccmb婚否.ItemData(Ccmb婚否.NewIndex) = i
    'Next
    'Ccmb婚否.Text = Ccmb婚否.List(0)
    
    '2012-04-14 于登淼 ↓
    '“修改”按钮没有功能，故隐藏
    ctbMain.Buttons(6).Visible = False
    '2012-04-14 于登淼 ↑
    
    '2012-06-13 于登淼 ↓
    '省疾控新要求，去掉体检项目调整和打印功能
    ctbMain.Buttons(3).Visible = False
    ctbMain.Buttons(4).Visible = False
    ctbMain.Buttons(8).Visible = False
    ctbMain.Buttons(9).Visible = False
    '2012-06-13 于登淼 ↑
    
    '2012-06-13 于登淼 ↓
    '初始化体格一般情况界面部分
    subInit体格一般情况
    '2012-06-13 于登淼 ↑
    
    Set lobj婚否 = pobjDict.FetchEx("婚姻字典")
    Ccmb婚否(mIndex).Clear
    Ccmb婚否(mIndex).AddItem ""
    For i = 1 To lobj婚否.RecordCount
        Ccmb婚否(mIndex).AddItem lobj婚否("名称")
        Ccmb婚否(mIndex).ItemData(Ccmb婚否(mIndex).NewIndex) = lobj婚否("编号")
        lobj婚否.MoveNext
    Next
    Ccmb婚否(mIndex).ListIndex = 0
    
    '职业史GRID
    Set mcolIndex = New Collection
    For i = 0 To cgrd职业史.Cols - 1
        mcolIndex.Add i, cgrd职业史.TextMatrix(0, i)
    Next
    '初始化界面时，放射工作信息录入框disable modify by lanchao 2015-8-17
    If Chk放射.Value = 1 Then
        Frame11.Enabled = True
    Else
        Frame11.Enabled = False
    End If
    '设置婚姻史界面不可用
'    If Not mIndex = 2 Then
'        ctxtmatehelh(mIndex).Enabled = False
'        ctxtmarrydate(mIndex).Enabled = False
'        ctxtmatejob(mIndex).Enabled = False
'        ctxtmateradioac(mIndex).Enabled = False
'    End If
'    Frame3(mIndex).Enabled = False
    '保存信息至数据库
    Set lobjInDtBase = CreateObject("职业病史录入.clsCareerhstregt")
    '职业病史GRID
    Set mcolIndexwkdis = New Collection
    For i = 0 To cgrd病史.Cols - 1
        mcolIndexwkdis.Add i, cgrd病史.TextMatrix(0, i)
    Next
    '自觉症状 GRID
    Set mcolindexzz = New Collection
    For i = 0 To cgrd症状.Cols - 1
        mcolindexzz.Add i, cgrd症状.TextMatrix(0, i)
    Next
    If 访问记号 = 1 Then
        'ctxtsysno.SetFocus
        ctxtsysno.Text = pub系统编号
        'Ccmb婚否.SetFocus
    End If
    
    
'    访问记号 = 0
'      MsgBox "窗体加载"   '2016-3-1 by 牟俊
    Timer1.Enabled = True
    sub连接终端
    Exit Sub
errHandler:
    sfsub错误处理 "职业病史录入", "frmcareerhstregt", "form_load", Err.Number, Err.Description, True
    Exit Sub
    Resume
End Sub


Public Sub sub查询填充表格()
Dim i As Integer
 '重新定义一个对象变量用来填充病状询问（原来是用的职业史的对象变量mcolIndex），并将下面的mcolIndex换成xmcolIndex  2015-12-11 by 牟俊
Dim xmcolIndex As Object
Dim lobjRec As Object
        dasubSetQueryTimeout 600
        Dim lstrsql As String
        lstrsql = "select 系统编号,症状,程度,出现时间 from 职业病体检_自觉症状表 where 系统编号='" & ctxtsysno.Text & "'"
        
        Set lobjRec = dafuncGetData(lstrsql)
        cgrdzzxw.Rows = 1
        
        If Not lobjRec.EOF Then
            With cgrdzzxw
                Set .DataSource = lobjRec
                If cgrdzzxw.Rows > 1 Then
                    Set xmcolIndex = New Collection
                    For i = 0 To cgrdzzxw.Cols - 1
                        xmcolIndex.Add i, cgrdzzxw.TextMatrix(0, i)
                    Next
                End If
              '  clblInfo = .Rows - 1
                .Col = 0
'                .Sort = flexSortGenericDescending
                .AutoSize 0, .Cols - 1, 0, 0
                .ExplorerBar = flexExSort
'                .DataMode = flexDMFree
             '   clblInfo = .Rows - 1
            End With
            
            Exit Sub
        Else
            cgrdzzxw.Rows = 1
          '  clblInfo = cgrdzzxw.Rows - 1
            Exit Sub
        End If

End Sub


'窗体取消
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    '设置标志pblnInUse。
    mblninuse = False
    '释放模块级对象。
    Set mobjGUI = Nothing
    Set lobjInDtBase = Nothing
    Unload frmCareerHstRegt
End Sub




'点击工具栏按钮后响应操作
Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    Dim i As Long
    Dim lobjRec As Object, lobj体检类型 As Object
    Dim str体检类型 As String
    Dim totalPay As Double
    Dim bol填充 As Boolean
    Dim lcol原体检项目 As Collection
    Dim lcol编号 As Collection
    Dim lstr收费批号 As String
    'Set lobjRec = CreateObject("职业病史录入.clscareerhstregt")
    Dim lstrError As String
    On Error GoTo errHandler
    '2012-05-22 于登淼 ↓
    '不想让“没有可以保存的内容”窗体弹出
    Cancel = True
    '2012-05-22 于登淼 ↑
    Select Case Operate
    Case "清空"
        subclearps
        subclear
        subokclear
        subokclear病史
        subokclear症状
        subClear体格一般情况
        cgrd职业史.Clear
        cgrd病史.Clear
        cgrd症状.Clear
        subclearps
        '表格内容清除
        'cgrd职业史.Clear
        cgrd职业史.Rows = 1
        'cgrd病史.Clear
        cgrd病史.Rows = 1
        'cgrd症状.Clear
        cgrd症状.Rows = 1
        ctxtsysno.SetFocus
    Case "去掉人员"
    Case "清空人员"
    '2012-06-14 于登淼 ↓
    '取消打印功能，case “保存并打印”全部注释
'''    Case "保存并打印"
'''
'''        If bolenProject = False Then
'''            MsgBox "还没确定体检项目，请点击体检项目，确定后才能保存！"
'''            Exit Sub
'''        End If
'''        '开始事务。
'''        dasubBeginTran
'''        '个人生活史保存
'''        With lobjPsLifeHst
'''            .mstr婚否 = Ccmb婚否.Text
'''            .mstrmatehelh = Trim(ctxtmatehelh.Text)
'''            .mstrmarrydate = Trim(ctxtmarrydate.Value)
'''            .mstrmatejob = Trim(ctxtmatejob.Text)
'''            .mstrmateradioac = Trim(ctxtmateradioac.Text)
'''            .mstr异位妊娠 = Trim(ctxt异位妊娠.Text)
'''            .mstr孕次 = Trim(ctxt孕次.Text)
'''            .mstr活产 = Trim(ctxt活产.Text)
'''            .mstr早产 = Trim(ctxt早产.Text)
'''            .mstr死产 = Trim(ctxt死产.Text)
'''            .mstr现有子女 = Trim(ctxt现有子女.Text)
'''            .mstr流产 = Trim(ctxt流产.Text)
'''            .mstr畸胎 = Trim(ctxt畸胎.Text)
'''            .mstr多胎 = Trim(ctxt多胎.Text)
'''            .mstr子女健康 = Trim(ctxt子女健康.Text)
'''            .mstr不孕不育 = Trim(ctxt不孕不育.Text)
'''            .mstr饮酒 = ccmb饮酒.Text
'''            .mstr吸烟 = ccmb吸烟.Text
'''            .mstr酒龄 = Trim(ctxt酒龄.Text)
'''            .mstr烟龄 = Trim(ctxt烟龄.Text)
'''            .mstr戒烟时长 = Trim(ctxt戒烟.Text)
'''            .mstr过敏史 = Trim(ctxt过敏史.Text)
'''            .mstr家族史 = Trim(ctxt家族.Text)
'''            .mstr初潮 = Trim(ctxt初潮.Text)
'''            .mstr经期 = Trim(ctxt经期.Text)
'''            .mstr周期 = Trim(ctxt周期.Text)
'''            .mstr末次月经 = Trim(ctxt末次月经.Text)
'''            .mstr停经 = Trim(ctxt停经.Text)
'''            If 访问记号 = 1 Then
'''                .subDelLifeHst   '删除个人生活史
'''            End If
'''
'''            '张元娅2012.04.10
'''            lobjInDtBase.系统编号 = Trim(ctxtsysno.Text)
'''            '张元娅2012.04.10
'''
'''            .subSaveLifeHst   '保存个人生活史
'''        End With
'''
'''        '保存工作史
'''        saveworkhst
'''
'''        '既往病史保存
'''        SavePastMedcHst
'''
'''        '自觉症状保存
'''        SaveSymptom
'''
'''        '修改体检状态
'''        lobjInDtBase.sub修改体检状态
'''
'''        '保存体检项目  职业病史录入后
'''        Set lobjRec = CreateObject("职业病史录入.clscareerhstregt")
'''        lobjRec.系统编号 = Trim(ctxtsysno.Text)
'''        Set lobjRec.col体检项目 = mcol体检项目
'''        lobjRec.save优化的体检项目
'''
'''        lstrError = lobjRec.func收费(lstr收费批号)
'''        If lstrError <> "" And lstrError <> "Cancel" Then
'''            MsgBox lstrError, vbOKOnly + vbExclamation, "系统提示"
'''        End If
'''        '结束事务。
'''        dasubCommitTran
'''
'''        '打印
''''        Set lcol编号 = New Collection
''''        lcol编号.Add Trim(ctxtsysno.Text)
''''
''''        Set lobj体检类型 = CreateObject("职业病对象.clsmedicalexam")
''''        lobj体检类型.系统编号 = Trim(ctxtsysno.Text)
''''        str体检类型 = lobj体检类型.体检类型
''''            '打印
''''            pobj业务对象.Sub打印文书 str体检类型 & "体检登记单", lcol编号, False
''''            'Cancel = True
''''        Set lobj体检类型 = Nothing
'''        subPrint Trim(ctxtsysno.Text)
'''
'''        访问记号 = 0
'''        Set lobjInDtBase.mobjWorkHst = New Collection
'''        Set lobjInDtBase.mobjPastHst = New Collection
'''        Set lobjInDtBase.mobjSymptom = New Collection
'''        If ChkClear Then
'''            subclearps
'''            subclear
'''            subokclear
'''            subokclear病史
'''            subokclear症状
'''            cgrd职业史.Rows = 1
'''            cgrd病史.Rows = 1
'''            cgrd症状.Rows = 1
'''        End If
'''        subclearps
'''        '表格内容清除
'''        cgrd职业史.Rows = 1
'''        cgrd病史.Rows = 1
'''        cgrd症状.Rows = 1
'''        ctxtsysno.SetFocus
'''        Cancel = True
'''
'''        MsgBox "打印成功！"
    '2012-06-14 于登淼 ↓
    '取消体检项目调整功能，全部注释 case "体检项目"
'''    Case "体检项目"
'''        Dim lobj体检模板 As Object
'''
'''        If ctxtsysno.Text = "" Then
'''            MsgBox "系统编号不能为空！"
'''            Exit Sub
'''        End If
'''        '获取体检表上已有的体检项目。
'''        mobj体检.系统编号 = Trim(ctxtsysno.Text)
'''        Set lcol原体检项目 = mobj体检.体检表.体检项目集("")
'''
'''        '设置选择项目界面的属性。
'''        frmSelectItem.pstr体检表名称 = "从业人员初检表"
'''        Set frmSelectItem.pcol复查项目 = lcol原体检项目
'''        'Set frmSelectItem.pcol收费项目 = mcol收费项目
'''        '启动选择项目界面。
'''        frmSelectItem.Show 1
'''        If frmSelectItem.pblnOk Then
'''            '获取选中的复查项目。
'''            Set mcol体检项目 = frmSelectItem.pcol复查项目
'''            '获取设置的收费项目。
'''            'Set mcol收费项目 = frmSelectItem.pcol收费项目
'''
'''            '显示收费金额。
'''            Dim ldblTotal As Double
'''            'For i = 1 To mcol收费项目.Count
'''            '    ldblTotal = Format(ldblTotal + mcol收费项目(i)("单价"), "0.00")
'''            'Next
'''            'On Error Resume Next
'''            'If sffunc判断集合键值是否存在(mobj体检.体检表.附加信息, "体检金额") Then
'''                'ciptBase.Box1("体检金额").Text = ldblTotal
'''                'mobj体检.体检表.Sub填附加信息值 "体检金额", ldblTotal
'''                mobj体检.体检表.Sub填附加信息值 "体检金额", 100
'''            'End If
'''            bolenProject = True '已确定体检项目
'''        End If
    Case "导出"
    
    Case "保存"
        '2012-06-14 于登淼
        'case "保存" 部分，关于项目体检部分全部注释
        
'''        '判断是否已确定体检项目
'''
'''        '功能：记录人员的体检项目
'''        '时间：2012-06-05
'''        '作者：翁乔
'''        Dim lobj体检项目 As Object
'''        Set lobj体检项目 = CreateObject("职业病史录入.ClsCareerHstRegt")
'''        Set mcol体检项目 = lobj体检项目.func获取体检人员的体检项目(ctxtsysno.Text)
'''
'''        '显示收费金额。
'''        mobj体检.体检表.Sub填附加信息值 "体检金额", 100
'''        bolenProject = True
'''        '时间：2012-06-05
'''
'''        If bolenProject = False Then
'''            MsgBox "还没确定体检项目，请点击体检项目，确定后才能保存！"
'''            Exit Sub
'''        End If
        '开始事务。
        dasubBeginTran
        '个人生活史保存
        With lobjPsLifeHst
            .mstr婚否 = Ccmb婚否(mIndex).Text
            .mstr现有子女 = Trim(ctxt现有子女(mIndex).Text)

            If mIndex <> 2 Then
                If Ccmb婚否(mIndex).Text = "已婚" Or Ccmb婚否(mIndex).Text = "离异" Then
                    '.mstrmarrydate = Trim(ctxtmarrydate(mIndex).Value)
                    .mstrmarrydate = Trim(ctxtmarrydate(mIndex).Text)
                    .mstrmatehelh = Trim(ctxtmatehelh(mIndex).Text)
                    .mstrmatejob = Trim(ctxtmatejob(mIndex).Text)
                    .mstrmateradioac = Trim(ctxtmateradioac(mIndex).Text)
                End If
                If mIndex <> 1 Then
                    .mstr异位妊娠 = Trim(ctxt异位妊娠(mIndex).Text)
                End If
                .mstr孕次 = Trim(ctxt孕次(mIndex).Text)
                .mstr活产 = Trim(ctxt活产(mIndex).Text)
                .mstr畸胎 = Trim(ctxt畸胎(mIndex).Text)
                .mstr多胎 = Trim(ctxt多胎(mIndex).Text)
                .mstr子女健康 = Trim(ctxt子女健康(mIndex).Text)
                .mstr不孕不育 = Trim(ctxt不孕不育(mIndex).Text)
                .mstrMore = Trim(ctxtMore(mIndex).Text)
            End If
            .mstr异常胎 = Trim(ctxt异常胎.Text)
            .mstr早产 = Trim(ctxt早产(mIndex).Text)
            .mstr死产 = Trim(ctxt死产(mIndex).Text)
            If mIndex <> 1 Then
            '8023部队增加流产和烟龄酒龄，所以放到下面一起 2015-9-29
'                .mstr现有子女 = Trim(ctxt现有子女(mIndex).Text)
'                .mstr流产 = Trim(ctxt流产(mIndex).Text)
'                .mstr饮酒 = ccmb饮酒(mIndex).Text
'                .mstr吸烟 = ccmb吸烟(mIndex).Text
'                .mstr酒龄 = Trim(ctxt酒龄(mIndex).Text)
'                .mstr烟龄 = Trim(ctxt烟龄(mIndex).Text)
'                .mstr戒烟时长 = Trim(ctxt戒烟(mIndex).Text)
                If mIndex <> 0 Then
                    .mstr过敏史 = Trim(ctxt过敏史(mIndex).Text)
                End If
            End If
            '增加吸烟饮酒程度，戒烟时长  2015-11-10 by 牟俊
            .mstr饮酒 = ccmb饮酒(mIndex).Text
            .mstr吸烟 = ccmb吸烟(mIndex).Text
            .mstr戒烟时长 = Trim(ctxt戒烟(mIndex).Text)
            
            .mstr吸烟量 = Trim(ctxt吸烟量(mIndex).Text)
            .mstr饮酒量 = Trim(ctxt饮酒量(mIndex).Text)
            .mstr烟龄 = Trim(ctxt烟龄(mIndex).Text)
            .mstr酒龄 = Trim(ctxt酒龄(mIndex).Text)
            .mstr流产 = Trim(ctxt流产(mIndex).Text)
            
            .mstr家族史 = Trim(ctxt家族.Text)
            .mstr初潮 = Trim(ctxt初潮.Text)
            .mstr经期 = Trim(ctxt经期.Text)
            .mstr周期 = Trim(ctxt周期.Text)
            .mstr末次月经 = Trim(ctxt末次月经.Text)
            .mstr停经 = Trim(ctxt停经.Text)
            .mstrOther = Trim(ctxtOther.Text)
            '放射健康体检处理女孩数量、男孩、女孩出生日期 2015-7-1 by lanchao
            If mIndex = 0 Then
            .mstr现有女孩 = Trim(ctxt现有女孩(mIndex).Text)
            .mstr男孩出生日期 = Trim(ctxt男孩出生日期(mIndex).Text)
            .mstr女孩出生日期 = Trim(ctxt女孩出生日期(mIndex).Text)
            End If
              '8023部队体检处理男孩女孩数量和出生日期 2015 - 9 - 28 by 牟俊
             If mIndex = 1 Then
            .mstr现有女孩 = Trim(ctxt现有女孩(mIndex).Text)
            .mstr男孩出生日期 = Trim(ctxt男孩出生日期(mIndex).Text)
            .mstr女孩出生日期 = Trim(ctxt女孩出生日期(mIndex).Text)
            End If
            
            '职业健康增加出生地    2015-11-10  by 牟俊
            If mIndex = 2 Then
            dafuncGetData ("update 职业病体检_体检人员基本信息表 set 出生地='" & ctxt出生地(mIndex).Text & "' where  系统编号='" & Trim(ctxtsysno.Text) & "'")
            End If
            
            If 访问记号 = 1 Then
                .subDelLifeHst   '删除个人生活史
            End If
            
            '张元娅2012.04.10
            lobjInDtBase.系统编号 = Trim(ctxtsysno.Text)
            '张元娅2012.04.10
            
            .subSaveLifeHst   '保存个人生活史
        End With
        
        '2012-06-14 于登淼 ↓
        subSave体格一般情况 ctxtsysno.Text
        '2012-06-14 于登淼 ↑
    
        '保存工作史
        saveworkhst
        
        '既往病史保存
        SavePastMedcHst
        
        '自觉症状保存 2015-8-20 by lanchao 整体保存的时候不操作
        'SaveSymptom
        
        '修改体检状态
        lobjInDtBase.sub修改体检状态
        
'''        '保存体检项目  职业病史录入后
'''        Set lobjRec = CreateObject("职业病史录入.clscareerhstregt")
'''        lobjRec.系统编号 = Trim(ctxtsysno.Text)
'''        Set lobjRec.col体检项目 = mcol体检项目
'''        lobjRec.save优化的体检项目
'''
'''
'''        lstrError = lobjRec.func收费(lstr收费批号)
'''        If lstrError <> "" And lstrError <> "Cancel" Then
'''            MsgBox lstrError, vbOKOnly + vbExclamation, "系统提示"
'''        End If
        
        '2012-06-15 于登淼 ↓
        '更改体检状态，由"未录入受检者个人信息"2，变为"体检中"3
        pobj业务对象.func写入单人当前体检状态 ctxtsysno.Text, 3
        '2012-06-15 于登淼 ↑
        
        '2012-07-04 于登淼 ↓
        '更新各科体检状态，和总的体检状态
        pobj业务对象.sub修改结果录入状态 ctxtsysno.Text, "13", "2"  '13代表个人信息录入科室，且表示在各科体检状态字符串的位置（非科室编号）
        pobj业务对象.sub结果录入修改体检状态 Trim(ctxtsysno.Text), "4"
        '2012-07-04 于登淼 ↑
        
        '2012-08-22 于登淼 ↓
        '填入该科室体检结论（填完）
        pobj业务对象.sub单个填写体检结论 Trim(ctxtsysno.Text), "受检者个人信息录入科", "填完", um用户编号
        '2012-08-22 于登淼 ↑
        
'        '将清空放到事务里执行   2016-2-24 by 牟俊
'        If ChkClear Then
'            subclearps
'            subclear
'            subokclear
'            subokclear病史
'            subokclear症状
'            subClear体格一般情况
'            cgrd职业史.Rows = 1
'            cgrd病史.Rows = 1
'            cgrd症状.Rows = 1
'        End If
        
        
        '结束事务。
        dasubCommitTran
        If MSComm1.PortOpen = True Then
            MSComm1.PortOpen = False    '保存后关掉终端连接
        End If
        MsgBox "保存成功！"
        访问记号 = 0
        Set lobjInDtBase.mobjWorkHst = New Collection
        Set lobjInDtBase.mobjPastHst = New Collection
        Set lobjInDtBase.mobjSymptom = New Collection
        
'        '保存成功之后不清空界面 2016-1-7 by 牟俊 ↓  暂时现不屏蔽
'        If ChkClear Then
'            subclearps
'            subclear
'            subokclear
'            subokclear病史
'            subokclear症状
'            subClear体格一般情况
'            cgrd职业史.Rows = 1
'            cgrd病史.Rows = 1
'            cgrd症状.Rows = 1
'        End If
'    '保存成功之后不清空界面 2016-1-7 by 牟俊 ↑   暂时现不屏蔽

'        MsgBox "保存过程" '2016-3-1 by 牟俊
        '关闭当前窗口 2015-8-17 modify by lanchao
        Unload frmCareerHstRegt
        '修改人：罗李奎 2012-12-11 ↓
        '说明：选择保存后不清空
        'bug号：0000059
'        subclearps
     
        
        '表格内容清除
'        cgrd职业史.Rows = 1
'        cgrd病史.Rows = 1
'        cgrd症状.Rows = 1
'        ctxtsysno.SetFocus
'        Cancel = True
   '修改人：罗李奎 2012-12-11 ↑
    Case "退出"
        If Len(ctxtsysno.Text) > 0 Then
            i = MsgBox("是否保存当前录入信息？", vbYesNo, "系统提示")
            If i = vbYes Then mobjGUI_BeforeOperate "保存", True
        End If
        '2012-05-22 于登淼 ↓
        '如果中间操作失误，可能导致退出按钮不能用。加上这行，弥补这个功能
        Unload frmCareerHstRegt
        '2012-05-22 于登淼 ↑
    End Select
    
    Exit Sub
errHandler:
    sfsub错误处理 "职业病史录入", "frmcareerhstregt", "mobjGUI_BeforeOperate", Err.Number, Err.Description, True
    MousePointer = 0
'    csbMain.Panels(1) = ""
    Exit Sub
    Resume
End Sub

'将表格中内容保存至数据库  自觉症状
Private Sub SaveSymptom()
    Dim i As Integer
    Dim lobjdetail As Object
    On Error GoTo errHandler
    For i = 1 To cgrd症状.Rows - 1
    '创建症状对象
    Set lobjdetail = CreateObject("职业病史录入.clssymptomdetl")
    '张元娅2012.04.10
    '为系统编号赋值
    lobjInDtBase.系统编号 = Trim(ctxtsysno.Text)
    '张元娅2012.04.10
        lobjdetail.mstr编号 = cgrd症状.Cell(flexcpText, i, mcolindexzz("编号"))
        lobjdetail.mstr症状 = cgrd症状.Cell(flexcpText, i, mcolindexzz("症状"))
        lobjdetail.mstr出现时间 = cgrd症状.Cell(flexcpText, i, mcolindexzz("病程时间"))
        lobjdetail.mstr程度 = cgrd症状.Cell(flexcpText, i, mcolindexzz("程度"))
        lobjInDtBase.mobjSymptom.Add lobjdetail
    Next
    If 访问记号 = 1 Then
        lobjInDtBase.subDelSymptom
    End If
    lobjInDtBase.SubSaveSymptom
    Exit Sub
errHandler:
    sfsub错误处理 "职业病史录入", "frmcareerhstregt", "savesymptom", Err.Number, Err.Description, True
End Sub

'将表格中内容保存至数据库  既往病史
Private Sub SavePastMedcHst()
    Dim i As Integer
    Dim lobjdetail As Object
    On Error GoTo errHandler
    For i = 1 To cgrd病史.Rows - 1
    '创建病史对象
    Set lobjdetail = CreateObject("职业病史录入.clsPastMedcHstdetl")
    '张元娅2012.04.10
    '为系统编号赋值
    lobjInDtBase.系统编号 = Trim(ctxtsysno.Text)
    '张元娅2012.04.10
        lobjdetail.mstr编号 = cgrd病史.Cell(flexcpText, i, mcolIndexwkdis("编号"))
        lobjdetail.mstr疾病名称 = cgrd病史.Cell(flexcpText, i, mcolIndexwkdis("疾病名称"))
        lobjdetail.mstr诊断单位 = cgrd病史.Cell(flexcpText, i, mcolIndexwkdis("诊断单位"))
        lobjdetail.mstr诊断日期 = cgrd病史.Cell(flexcpText, i, mcolIndexwkdis("诊断日期"))
        lobjdetail.mstr转归 = cgrd病史.Cell(flexcpText, i, mcolIndexwkdis("转归"))
        lobjdetail.mstr治疗经过 = cgrd病史.Cell(flexcpText, i, mcolIndexwkdis("治疗经过"))
        lobjInDtBase.mobjPastHst.Add lobjdetail
    Next
    If 访问记号 = 1 Then
        lobjInDtBase.subDelPastMedcHst
    End If
    lobjInDtBase.SubSavePastMedcHst
    Exit Sub
errHandler:
    sfsub错误处理 "职业病史录入", "frmcareerhstregt", "savepastmedchst", Err.Number, Err.Description, True
End Sub

'将表格中内容保存至数据库  职业史
Private Sub saveworkhst()
    Dim i As Integer
    Dim lobjdetail As Object
    On Error GoTo errHandler
    
    Set mcolIndex = New Collection
    For i = 0 To cgrd职业史.Cols - 1
        mcolIndex.Add i, cgrd职业史.TextMatrix(0, i)
    Next
    For i = 1 To cgrd职业史.Rows - 1
        '创建职业史对象
        Set lobjdetail = CreateObject("职业病史录入.clscareerhstDetl")
        lobjInDtBase.系统编号 = Trim(ctxtsysno.Text)
    
        lobjdetail.mstr编号 = cgrd职业史.Cell(flexcpText, i, mcolIndex("编号"))
        lobjdetail.mstr单位 = cgrd职业史.Cell(flexcpText, i, mcolIndex("工作单位"))
        lobjdetail.mstr部门 = cgrd职业史.Cell(flexcpText, i, mcolIndex("部门"))
        lobjdetail.mstr工种 = cgrd职业史.Cell(flexcpText, i, mcolIndex("工种"))
        lobjdetail.mstr危害种类 = cgrd职业史.Cell(flexcpText, i, mcolIndex("危害种类"))
        lobjdetail.mstr接触时间 = cgrd职业史.Cell(flexcpText, i, mcolIndex("接触时间"))
        lobjdetail.mstr措施 = cgrd职业史.Cell(flexcpText, i, mcolIndex("防护措施"))
        lobjdetail.mstr备注 = cgrd职业史.Cell(flexcpText, i, mcolIndex("备注"))
        lobjdetail.mstr起始时间 = cgrd职业史.Cell(flexcpText, i, mcolIndex("起始时间"))
        lobjdetail.mstr结束时间 = cgrd职业史.Cell(flexcpText, i, mcolIndex("结束时间"))
       
        lobjdetail.mstr放射种类 = cgrd职业史.Cell(flexcpText, i, mcolIndex("放射种类"))
        lobjdetail.mstr工作量 = cgrd职业史.Cell(flexcpText, i, mcolIndex("每日工作量"))
        lobjdetail.mstr照射量 = cgrd职业史.Cell(flexcpText, i, mcolIndex("累积照射量"))
        lobjdetail.mstr过量照射史 = cgrd职业史.Cell(flexcpText, i, mcolIndex("过量照射史"))
        lobjdetail.mstr是否放射性 = cgrd职业史.Cell(flexcpText, i, mcolIndex("是否放射性"))
        lobjInDtBase.mobjWorkHst.Add lobjdetail
    Next
    If 访问记号 = 1 Then
        lobjInDtBase.subDelWorkHst
    End If
    lobjInDtBase.subSaveWorkHst
    Exit Sub
errHandler:
    sfsub错误处理 "职业病史录入", "frmcareerhstregt", "saveworkhst", Err.Number, Err.Description, True
End Sub

'清空个人信息
Private Sub subclearps()
    ctxtsysno.Text = ""
    Lab姓名.Caption = ""
    Lab性别.Caption = ""
    Lab年龄.Caption = ""
    lab单位.Caption = ""
    Lab现工种.Caption = ""
    Lab现职务.Caption = ""
    Label危害因素.Caption = ""
    Set Picture2.Picture = Nothing
    
    '2012-04-14 于登淼 ↓
    '清空、初始化窗体内全局变量
    mintrow = 0
    jmintrow = 0
    Set mobj体检 = CreateObject("职业病对象.clsMedicalExam")
    Set lobjInDtBase = CreateObject("职业病史录入.clsCareerhstregt")
    Set mcol体检项目 = frmSelectItem.pcol复查项目
    Set mcolindexzz = New Collection
    Set mcolIndexwkdis = New Collection
    Set mcolIndex = New Collection
    '2012-04-14 于登淼 ↑
     
End Sub
'清空界面信息  生活史
Private Sub subclear()
    Ccmb婚否(mIndex).Text = ""
    '将结婚日期，男孩女孩出生日期清空  2016-3-2 by 牟俊
    ctxtmarrydate(mIndex).Text = ""
    ctxt现有子女(mIndex).Text = "0"
    ctxt现有女孩(mIndex).Text = "0"
    ctxt男孩出生日期(mIndex).Text = ""
    ctxt女孩出生日期(mIndex).Text = ""
    ctxt酒龄(mIndex).Text = ""
    ctxt烟龄(mIndex).Text = ""
    
    If mIndex <> 2 Then
        'ctxtmatehelh(mIndex).Text = ""
        'modify by lanchao 2015-7.20,配偶健康默认为健康，不清空
        ctxtmatejob(mIndex).Text = ""
        ctxtmateradioac(mIndex).Text = ""
        If mIndex <> 1 Then
            ctxt异位妊娠(mIndex).Text = "0"
        End If
        ctxt孕次(mIndex).Text = "0"
         'modify by lanchao 2015-7.20,活产数量不清空，显示为X男X女
        If mIndex <> 1 Then
           ctxt活产(mIndex).Text = "0"
        End If
        ctxt畸胎(mIndex).Text = "0"
        ctxt多胎(mIndex).Text = "0"
        ctxt子女健康(mIndex).Text = ""
        ctxt不孕不育(mIndex).Text = ""
        ctxtMore(mIndex).Text = ""
    End If
    ctxt早产(mIndex).Text = "0"
    ctxt死产(mIndex).Text = "0"
    If mIndex <> 1 Then
        ctxt现有子女(mIndex).Text = "0"
        ctxt流产(mIndex).Text = "0"
'        ccmb饮酒(mIndex).Text = ""
'        ccmb吸烟(mIndex).Text = ""
        ctxt酒龄(mIndex).Text = ""
        ctxt烟龄(mIndex).Text = ""
        ctxt戒烟(mIndex).Text = ""
        If mIndex <> 0 Then
            ctxt过敏史(mIndex).Text = ""
        End If
    End If
    
    ccmb饮酒(mIndex).Text = ""
    ccmb吸烟(mIndex).Text = ""
        
    ctxt吸烟量(mIndex).Text = ""
    ctxt饮酒量(mIndex).Text = ""
    ctxt异常胎.Text = "0"
    ctxt家族.Text = ""
    ctxt初潮.Text = ""
    ctxt经期.Text = ""
    ctxt周期.Text = ""
    ctxt末次月经.Text = ""
    ctxt停经.Text = ""
End Sub

'清空  职业史
Private Sub subokclear()
     ctxt单位.Text = ""
     ctxt部门.Text = ""
     ctxt工种.Text = ""
     ctxt备注.Text = ""

     ctxt危害种类.Text = ""
     ctxt防护.Text = ""
     ctxt放射种类.Text = ""
     ctxtfangshe.Text = ""   '放射线种类清空  2015-12-2 by 牟俊
     ctxt工作量.Text = ""
     ctxt照射量.Text = ""
     ctxt过量照射.Text = ""
     ctxt起始.Text = "年月"
     ctxt结束.Text = "年月"
     ctxt接触.Text = ""        '不要接触时间  2015-12-11 by 牟俊
     ctxtweihai.Text = ""
          
     '修改人：罗李奎 2012-12-28   ↓
     '说明：还原时间
     'bug号：0000122 2015-6-26 不需要日期格式 by lanchao
     'ctxt起始.Value = Date
     'ctxt接触时间.Value = Date
     'ctxt结束.Value = Date
     '修改人：罗李奎 2012-12-28   ↑
End Sub

'清空  职业病史
Private Sub subokclear病史()
    ctxt疾病名称.Text = ""
    ctxt诊断单位.Text = ""
    ctxt转归.Text = ""
    ctxt治疗经过.Text = ""
    '修改：罗李奎 2012-12-7 ↓
    '恢复默认诊断时间
    'Bug号：0000057
    'ctxt诊断日期.Value = Date
    '修改：罗李奎 2012-12-7 ↑
End Sub

'2012.12.11 张令
'说明：清空生育史  ↓
Public Sub subclear生育史()
'    ctxt异位妊娠(Index).Text = ""
    ctxt孕次(mIndex).Text = "0"
    ctxt活产(mIndex).Text = "0"
    ctxt早产(mIndex).Text = "0"
    ctxt死产(mIndex).Text = "0"
    ctxt流产(mIndex).Text = "0"
    
    '未婚将男女及出生日期都赋空
    ctxt现有子女(mIndex).Text = "0"
    If mIndex <> 2 Then
    ctxt现有女孩(mIndex).Text = "0"
    ctxt男孩出生日期(mIndex).Text = ""
    ctxt女孩出生日期(mIndex).Text = ""
    End If
    '8023部队已经增加了这几项内容  2015-9-29
'    If mIndex <> 1 Then
''        ctxt现有子女(mIndex).Text = "0"
''        ctxt流产(mIndex).Text = "0"
'    End If
    ctxt畸胎(mIndex).Text = "0"
    ctxt多胎(mIndex).Text = "0"
    ctxt不孕不育(mIndex).Text = ""
    ctxt子女健康(mIndex).Text = ""
End Sub

'清空  自觉症状
Private Sub subokclear症状()
'    ctxt症状.Text = ""
'    ctxt程度.Text = ""
    ccmb分类.Text = ""
    ctxt程度.Text = ""
'    ctxt出现时间.Text = "年月"
End Sub

'情况受检者个人信息集合  2016-3-2 by 牟俊
Private Sub gatherclear()
    subclear
    subokclear
    subokclear病史
    subokclear症状
    subClear体格一般情况
    cgrd职业史.Rows = 1
    cgrd病史.Rows = 1
    cgrd症状.Rows = 1
End Sub

'修改 生活史
Private Sub sub修改生活史(ByVal lobjlife As Object)
    On Error GoTo errHandler
    '2012-04-14 于登淼 ↓
    '当没有信息时，直接退出
    If lobjlife.RecordCount = 0 Then Exit Sub
    '2012-04-14 于登淼 ↑
        Ccmb婚否(mIndex).Text = IIf(IsNull(lobjlife!是否结婚), "", lobjlife!是否结婚)
        Ccmb婚否_Click (mIndex)
        ctxt现有子女(mIndex).Text = IIf(IsNull(lobjlife!现有子女数目), "", lobjlife!现有子女数目)
        If mIndex <> 2 Then
            ctxtmatehelh(mIndex).Text = IIf(IsNull(lobjlife!配偶健康状况), "良好", lobjlife!配偶健康状况)
            '将日期格式修改为文本框  2015-6-26 by lanchao
            'ctxtmarrydate(mIndex).Value = Format(IIf(lobjlife!结婚日期 = "", Date, lobjlife!结婚日期), "yyyy/mm/dd")
            ctxtmarrydate(mIndex).Text = IIf(IsNull(lobjlife!结婚日期), "", lobjlife!结婚日期)
            ctxtmatejob(mIndex).Text = IIf(IsNull(lobjlife!配偶职业), "", lobjlife!配偶职业)
            ctxtmateradioac(mIndex).Text = IIf(IsNull(lobjlife!配偶接触放射), "", lobjlife!配偶接触放射)
            If mIndex <> 1 Then
                ctxt异位妊娠(mIndex).Text = IIf(IsNull(lobjlife!异位妊娠), "", lobjlife!异位妊娠)
            End If
            ctxt孕次(mIndex).Text = IIf(IsNull(lobjlife!孕次), "", lobjlife!孕次)
            ctxt活产(mIndex).Text = IIf(IsNull(lobjlife!活产), "", lobjlife!活产)
            ctxt畸胎(mIndex).Text = IIf(IsNull(lobjlife!畸胎), "", lobjlife!畸胎)
            ctxt多胎(mIndex).Text = IIf(IsNull(lobjlife!多胎), "", lobjlife!多胎)
            ctxt子女健康(mIndex).Text = IIf(IsNull(lobjlife!子女健康状况), "", lobjlife!子女健康状况)
            ctxt不孕不育(mIndex).Text = IIf(IsNull(lobjlife!不孕不育原因), "", lobjlife!不孕不育原因)
            ctxtMore(mIndex).Text = IIf(IsNull(lobjlife!生活更多), "", lobjlife!生活更多)
        End If
        ctxt异常胎.Text = IIf(IsNull(lobjlife!异常胎), "", lobjlife!异常胎)
        ctxt早产(mIndex).Text = IIf(IsNull(lobjlife!早产), "", lobjlife!早产)
        ctxt死产(mIndex).Text = IIf(IsNull(lobjlife!死产), "", lobjlife!死产)
        If mIndex <> 1 Then
            ctxt现有子女(mIndex).Text = IIf(IsNull(lobjlife!现有子女数目), "", lobjlife!现有子女数目)
'            ctxt流产(mIndex).Text = IIf(IsNull(lobjlife!自然流产), "", lobjlife!自然流产)

'            ccmb饮酒(mIndex).Text = IIf(IsNull(lobjlife!饮酒程度), "", lobjlife!饮酒程度)
'            ccmb吸烟(mIndex).Text = IIf(IsNull(lobjlife!吸烟程度), "", lobjlife!吸烟程度)

'            ctxt酒龄(mIndex).Text = IIf(IsNull(lobjlife!酒龄), "", lobjlife!酒龄)
'            ctxt烟龄(mIndex).Text = IIf(IsNull(lobjlife!烟龄), "", lobjlife!烟龄)
'            ctxt戒烟(mIndex).Text = IIf(IsNull(lobjlife!戒烟时长), "", lobjlife!戒烟时长)
            If mIndex <> 0 Then
                ctxt过敏史(mIndex).Text = IIf(IsNull(lobjlife!过敏史), "", lobjlife!过敏史)
            End If
        End If
        ccmb饮酒(mIndex).Text = IIf(IsNull(lobjlife!饮酒程度), "", lobjlife!饮酒程度)
        ccmb吸烟(mIndex).Text = IIf(IsNull(lobjlife!吸烟程度), "", lobjlife!吸烟程度)
        ctxt戒烟(mIndex).Text = IIf(IsNull(lobjlife!戒烟时长), "", lobjlife!戒烟时长)
        
        ctxt流产(mIndex).Text = IIf(IsNull(lobjlife!自然流产), "", lobjlife!自然流产)
        ctxt吸烟量(mIndex).Text = IIf(IsNull(lobjlife!吸烟量), "", lobjlife!吸烟量)
        ctxt饮酒量(mIndex).Text = IIf(IsNull(lobjlife!饮酒量), "", lobjlife!饮酒量)
        ctxt烟龄(mIndex).Text = IIf(IsNull(lobjlife!烟龄), "", lobjlife!烟龄)
        ctxt酒龄(mIndex).Text = IIf(IsNull(lobjlife!酒龄), "", lobjlife!酒龄)
        ctxt家族.Text = IIf(IsNull(lobjlife!家族史), "", lobjlife!家族史)
        ctxt初潮.Text = IIf(IsNull(lobjlife!初潮), "", lobjlife!初潮)
        ctxt经期.Text = IIf(IsNull(lobjlife!经期), "", lobjlife!经期)
        ctxt周期.Text = IIf(IsNull(lobjlife!周期), "", lobjlife!周期)
        ctxt末次月经.Text = IIf(IsNull(lobjlife!末次月经), "", lobjlife!末次月经)
        ctxt停经.Text = IIf(IsNull(lobjlife!停经年龄), "", lobjlife!停经年龄)
        ctxtOther.Text = IIf(IsNull(lobjlife!其他), "", lobjlife!其他)
        '处理放射健康体检男孩、女孩出生日期问题 2015-7-1 by lanchao
         If mIndex = 0 Then
            ctxt现有女孩(mIndex).Text = IIf(IsNull(lobjlife!现有女孩), "", lobjlife!现有女孩)
            ctxt男孩出生日期(mIndex).Text = IIf(IsNull(lobjlife!男孩出生日期), "", lobjlife!男孩出生日期)
            ctxt女孩出生日期(mIndex).Text = IIf(IsNull(lobjlife!女孩出生日期), "", lobjlife!女孩出生日期)
        End If
        '处理8023部队体检男孩、女孩出生日期问题 2015-9-28
         If mIndex = 1 Then
'            ctxt现有子女(mIndex).Text = IIf(IsNull(lobjlife!现有子女), "", lobjlife!现有子女)
            ctxt现有女孩(mIndex).Text = IIf(IsNull(lobjlife!现有女孩), "", lobjlife!现有女孩)
            ctxt男孩出生日期(mIndex).Text = IIf(IsNull(lobjlife!男孩出生日期), "", lobjlife!男孩出生日期)
            ctxt女孩出生日期(mIndex).Text = IIf(IsNull(lobjlife!女孩出生日期), "", lobjlife!女孩出生日期)
        End If
        
        '修改时显示上次写入的出生地信息  2015-11-10 by 牟俊
        If mIndex = 2 Then
        Dim csql As Object
        Set csql = dafuncGetData("select * From 职业病体检_体检人员基本信息表 where 系统编号='" & Trim(ctxtsysno.Text) & "'")
        ctxt出生地(mIndex).Text = csql("出生地")
        End If
        Exit Sub
errHandler:
    sfsub错误处理 "职业病史录入", "frmcareerhstregt", "sub修改生活史", Err.Number, Err.Description, True
End Sub


Private Sub MSComm1_OnComm()
'Dim S() As Byte
'    Dim SS(1024) As Byte
'    Static N As Long
'    Static T As Variant
'     Dim intInputLen, i As Integer
'Dim instring As String
'    If (MSComm1.CommEvent = comEvReceive) Then
'        S = MSComm1.Input                      '只要有数据就收进来，哪怕只是一个
'
'        T = Timer
'        For i = 0 To UBound(S)
'        '一个数据包可能产生若干个oncomm事件
'        instring = StrConv(S, vbUnicode)
'                         Text临时.Text = Text临时.Text & instring
'            SS(N + i) = S(i)                 '接收数据包缓存于SS()
'            N = N + UBound(S)
'        Next i
'       ' MSComm1.InBufferCount = 0
'    End If

    If MSComm1.InBufferCount Then
        ' 通讯埠中假如有资料的话, 则读取进来
           Dim InStringB() As Byte
           Dim instring As String
          InStringB = MSComm1.Input
          instring = StrConv(InStringB, vbUnicode)
          Text临时.Text = Text临时.Text & instring
          InStringB = ""
          If Len(Text临时.Text) < 13 Then
                MSComm1_OnComm
          End If
    End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
'去掉下面内容测试  2015-12-11 by 牟俊
'If PreviousTab = 5 Then
'sub查询填充表格
'End If
'改成只要点病症询问就显示  '2015-12-11 by 牟俊
If SSTab1.Tab = 5 Then
sub查询填充表格
End If

End Sub

'为了让系统编号获取焦点，再释放，以触发lostfocus事件
Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Dim i As Integer
    Dim lobjRec As Object
    'Dim lobjDetl As Object
    On Error GoTo errHandler
    'Set lobjRec = CreateObject("职业病史录入.clscareerhstregt")
    
    '获取部门
    Set lobjRec = pobjDict.FetchEx("部门字典")
    ctxt部门.Clear
    ctxt部门.AddItem ""
    For i = 1 To lobjRec.RecordCount
        ctxt部门.AddItem lobjRec("名称")
        ctxt部门.ItemData(ctxt部门.NewIndex) = lobjRec("编号")
        lobjRec.MoveNext
    Next
    ctxt部门.ListIndex = 0
    
    '获取工种
    Set lobjRec = pobjDict.FetchEx("工种字典")
    ctxt工种.Clear
    ctxt工种.AddItem ""
    For i = 1 To lobjRec.RecordCount
        ctxt工种.AddItem lobjRec("名称")
        ctxt工种.ItemData(ctxt工种.NewIndex) = lobjRec("编号")
        lobjRec.MoveNext
    Next
    ctxt工种.ListIndex = 0
    
    '获取放射线种类
    Set lobjRec = pobjDict.FetchEx("放射线种类字典")
    ctxt放射种类.Clear
    ctxt放射种类.AddItem ""
    For i = 1 To lobjRec.RecordCount
        ctxt放射种类.AddItem lobjRec("名称")
        ctxt放射种类.ItemData(ctxt放射种类.NewIndex) = lobjRec("编号")
        lobjRec.MoveNext
    Next
    ctxt放射种类.ListIndex = 0
    
    '获取职业危害种类
    Set lobjRec = pobjDict.FetchEx("危害种类字典")
    ctxt危害种类.Clear
    ctxt危害种类.AddItem ""
    For i = 1 To lobjRec.RecordCount
        ctxt危害种类.AddItem lobjRec("名称")
        ctxt危害种类.ItemData(ctxt危害种类.NewIndex) = lobjRec("编号")
        lobjRec.MoveNext
    Next
    ctxt危害种类.ListIndex = 0
    
    '获取自觉症状程度
    Set lobjRec = pobjDict.FetchEx("病情程度字典")
    ctxt程度.Clear
    ctxt程度.AddItem ""
    For i = 1 To lobjRec.RecordCount
        ctxt程度.AddItem lobjRec("名称")
        ctxt程度.ItemData(ctxt程度.NewIndex) = lobjRec("编号")
        lobjRec.MoveNext
    Next
    ctxt程度.ListIndex = 0
    
    
    Set lobjRec = pobjDict.FetchEx("职业病过敏源字典")
    Combo2.Clear
    Combo2.AddItem ""
    For i = 1 To lobjRec.RecordCount
       Combo2.AddItem lobjRec("名称")
        Combo2.ItemData(Combo2.NewIndex) = lobjRec("编号")
        lobjRec.MoveNext
    Next
    Combo2.ListIndex = 0
    
    
    Set lobjRec = pobjDict.FetchEx("职业病防护措施字典")
    Combo4.Clear
    Combo4.AddItem ""
    For i = 1 To lobjRec.RecordCount
    Combo4.AddItem lobjRec("名称")
    Combo4.ItemData(Combo4.NewIndex) = lobjRec("编号")
    lobjRec.MoveNext
    Next
    Combo4.ListIndex = 0
    
'   8023不单独处理   2015-11-26 by 牟俊
'      If Label体检类型.Caption = "8023部队" Then
'    Set lobjRec = pobjDict.FetchEx("职业病8023疾病字典")
'    Combo3.Clear
'    Combo3.AddItem ""
'    For i = 1 To lobjRec.RecordCount
'    Combo3.AddItem lobjRec("名称")
'    Combo3.ItemData(Combo3.NewIndex) = lobjRec("编号")
'    lobjRec.MoveNext
'    Next
'    Combo3.ListIndex = 0
'    Else
    Set lobjRec = pobjDict.FetchEx("职业病疾病名称字典")
    Combo3.Clear
    Combo3.AddItem ""
    For i = 1 To lobjRec.RecordCount
    Combo3.AddItem lobjRec("名称")
    Combo3.ItemData(Combo3.NewIndex) = lobjRec("编号")
    lobjRec.MoveNext
    Next
    Combo3.ListIndex = 0
'    End If
    
    
    
      Set lobjRec = pobjDict.FetchEx("职业病常见治疗经过字典")
    Combo5.Clear
    Combo5.AddItem "请选择治疗经过常见模板。"
    For i = 1 To lobjRec.RecordCount
    Combo5.AddItem lobjRec("名称")
    Combo5.ItemData(Combo5.NewIndex) = lobjRec("编号")
    lobjRec.MoveNext
    Next
    Combo5.ListIndex = 0
    
          Set lobjRec = pobjDict.FetchEx("职业病转归字典")
    Combo6.Clear
    Combo6.AddItem ""
    For i = 1 To lobjRec.RecordCount
    Combo6.AddItem lobjRec("名称")
    Combo6.ItemData(Combo6.NewIndex) = lobjRec("编号")
    lobjRec.MoveNext
    Next
    Combo6.ListIndex = 0
    '获取自觉症状分类
'    Set lobjRec = pobjDict.FetchEx("职业病体检症状字典")
    dasubSetQueryTimeout 6000
    Set lobjRec = dafuncGetData("select 名称,编号 from 系统管理_字典_字典内容表 where ID in " _
                            & "(select ID from 系统管理_字典_字典表列表 where 名称='职业病体检症状字典') and Parent='0'")
    ccmb分类.Clear
    ccmb分类.AddItem ""
    For i = 1 To lobjRec.RecordCount
        ccmb分类.AddItem lobjRec("名称")
        ccmb分类.ItemData(ccmb分类.NewIndex) = lobjRec("编号")
        lobjRec.MoveNext
    Next
    ccmb分类.ListIndex = 0
    
    If ccmb分类.Text <> "" Then
        '获取自觉症状项目
        dasubSetQueryTimeout 6000
        Set lobjRec = dafuncGetData("select 名称,编号 from 系统管理_字典_字典内容表 where parent in " _
                                & "(select innerid from 系统管理_字典_字典内容表 where 名称 = '" & ccmb分类.Text & " ')")
        clst项目.Clear
    '    ccmb分类.AddItem ""
        For i = 1 To lobjRec.RecordCount
            clst项目.AddItem lobjRec("名称")
            clst项目.ItemData(clst项目.NewIndex) = lobjRec("编号")
            lobjRec.MoveNext
        Next
        clst项目.ListIndex = 0
    End If
    
    
    '8023即mIndex=1时也需要程度字典，所有将判断语句去掉   2015-11-11 by 牟俊
'    If mIndex <> 1 Then
        Set lobjRec = pobjDict.FetchEx("程度字典")
        ccmb饮酒(mIndex).Clear
        ccmb饮酒(mIndex).AddItem ""
        For i = 1 To lobjRec.RecordCount
            ccmb饮酒(mIndex).AddItem lobjRec("名称")
            ccmb饮酒(mIndex).ItemData(ccmb饮酒(mIndex).NewIndex) = lobjRec("编号")
            lobjRec.MoveNext
        Next
        ccmb饮酒(mIndex).ListIndex = 0
        Set lobjRec = pobjDict.FetchEx("程度字典")
        ccmb吸烟(mIndex).Clear
        ccmb吸烟(mIndex).AddItem ""
        For i = 1 To lobjRec.RecordCount
            ccmb吸烟(mIndex).AddItem lobjRec("名称")
            ccmb吸烟(mIndex).ItemData(ccmb吸烟(mIndex).NewIndex) = lobjRec("编号")
            lobjRec.MoveNext
        Next
        ccmb吸烟(mIndex).ListIndex = 0
'    End If
    
    Set lobjRec = Nothing
    '注销 刘伟 2015-4-7
    'ctxt出现时间.Text = Date
'    ctxt诊断日期.Value = Date
'    ctxt接触时间.Value = Date
'    ctxt起始.Value = Date
'    ctxt结束.Value = Date
    cgrd职业史.ColHidden(mcolIndex("系统编号")) = True
    '职业健康的时候显示，否则不隐藏
    If Label体检类型.Caption <> "职业健康" Then
       cgrd职业史.ColHidden(mcolIndex("接触时间")) = True
    End If
    cgrd病史.ColHidden(mcolIndexwkdis("系统编号")) = True
    cgrd症状.ColHidden(mcolindexzz("系统编号")) = True
    
    ctxtsysno.SetFocus
    If 访问记号 = 1 Then
        'Ccmb婚否(mIndex).SetFocus
    End If
'    MsgBox "timer1过程" '2016-3-1 by 牟俊
    Exit Sub
errHandler:
    sfsub错误处理 "职业病史录入", "frmcareerhstregt", "sub修改生活史", Err.Number, Err.Description, True
End Sub

Public Sub subPrint(ByVal para系统编号 As String)

'    Set lcol编号 = New Collection
'        lcol编号.Add Trim(ctxtsysno.Text)
'
'        Set lobj体检类型 = CreateObject("职业病对象.clsmedicalexam")
'        lobj体检类型.系统编号 = Trim(ctxtsysno.Text)
'        str体检类型 = lobj体检类型.体检类型
'            '打印
'            pobj业务对象.Sub打印文书 str体检类型 & "体检登记单", lcol编号, False
'            'Cancel = True
'        Set lobj体检类型 = Nothing
    On Error GoTo errHandler
    Dim lobjRec As Object
    Dim lcolInfo As Collection
    Dim lcolItem As Collection
    Dim lstrRptName As String
    Dim mstrPrint As String
    Dim mstrselect As String
    Dim lobj体检类型 As Object
    Dim str体检类型 As String
    Set lobj体检类型 = CreateObject("职业病对象.clsmedicalexam")
    lobj体检类型.系统编号 = Trim(para系统编号)
    str体检类型 = lobj体检类型.体检类型
    Set lobj体检类型 = Nothing
    lstrRptName = str体检类型 & "体检登记单"
    mstrPrint = Trim(para系统编号)

    Set lcolInfo = New Collection
    Set lcolItem = New Collection
    lcolItem.Add "系统编号", "名称"
    lcolItem.Add mstrPrint, "值"
    lcolInfo.Add lcolItem, lcolItem("名称")
'    Set lcolItem = New Collection
'    lcolItem.Add "选择条件", "名称"
'    lcolItem.Add mstrselect, "值"
'    lcolInfo.Add lcolItem, lcolItem("名称")
    '打印   false为不预览
    
     '提取签字图片
    Dim lpicPhoto As StdPicture
    Dim lobjSys As Object
    Set lobjSys = CreateObject("Scripting.FileSystemObject")
    '先拷贝空白签名的图片。
    lobjSys.CopyFile App.Path & "\空白照片.bmp", "c:\体检照片.bmp"
    
    Set lpicPhoto = pmfunc获取图片(mstrPrint, "职业病体检")
    
    'Set lpicPhoto = pmfunc获取图片("0001", "系统管理")
    '数据库里没有图片时，lpicPhoto返回值为0，而不是null
    If Not lpicPhoto = 0 Then
        SavePicture lpicPhoto, "c:\体检照片.bmp"
    End If
    
    Set lobjRec = CreateObject("职业病文书.cls文书")
    lobjRec.funcCHPrintReport lstrRptName, lcolInfo, App.Path, False

    Exit Sub
errHandler:
    Dim llngErr As Long
    Dim lstrError As String
    llngErr = Err.Number
    lstrError = Err.Description

    If llngErr = 20526 Then
        lstrError = "存在妨碍打印的问题。此错误产生的原因及解决方法如下： " & Chr(13) & Chr(10) & Chr(13) & Chr(10) _
                    & "  (1) 没有从 Windows 控制面板中安装打印机？" & Chr(13) & Chr(10) _
                    & "      解决：打开控制面板，双击“打印机”图标，选择“添加打印机”以装入打印机。" & Chr(13) & Chr(10) & Chr(13) & Chr(10) _
                    & "  (2) 打印机没在线？" & Chr(13) & Chr(10) _
                    & "      解决：检查打印机与计算机的连接是否正常。" & Chr(13) & Chr(10) & Chr(13) & Chr(10) _
                    & "  (3)  打印机阻塞或缺纸？" & Chr(13) & Chr(10) _
                    & "      解决：解决这些问题。" & Chr(13) & Chr(10) & Chr(13) & Chr(10) _
                    & "  (4) 试图在只能接受文本的打印机上打印窗体？" & Chr(13) & Chr(10) _
                    & "      解决：切换到一台能打印图形的打印机。"
        llngErr = 6666
    End If
    sfsub错误处理 "物资设备管理界面", "frmBiologMaterialApply", "subPrint", Err.Number, Err.Description, False
End Sub

'2012-06-13 于登淼 ↓
'初始化体格一般情况界面部分
Sub subInit体格一般情况()
    comb营养.Clear
    comb营养.AddItem "良好": comb营养.ItemData(comb营养.NewIndex) = 0
    comb营养.AddItem "中等": comb营养.ItemData(comb营养.NewIndex) = 1
    comb营养.AddItem "不良": comb营养.ItemData(comb营养.NewIndex) = 2
    comb营养.ListIndex = 0
 '增加发育的可选值 2015-11-27 by 牟俊
    comb发育.Clear
    comb发育.AddItem "正力型": comb发育.ItemData(comb发育.NewIndex) = 0
    comb发育.AddItem "无力型": comb发育.ItemData(comb发育.NewIndex) = 1
    comb发育.AddItem "超力型": comb发育.ItemData(comb发育.NewIndex) = 2
    comb发育.ListIndex = 0
End Sub

'2012-06-13 于登淼 ↓
'清空体格一般情况界面部分
Sub subClear体格一般情况()
    comb营养.ListIndex = 0
    ctxt身高.Text = ""
    ctxt体重.Text = ""
    ctxt体形.Text = ""
    ctxt体重指数.Text = ""
    ctxt收缩压.Text = ""
    ctxt舒张压.Text = ""
    comb发育.ListIndex = 0
End Sub

'2012-06-13 于登淼 ↓
'载入体格一般情况界面部分(体格一般情况属于体检的内科)
Sub subLoad体格一般情况(ByVal paraSysNo As String)
    Dim lobjRec As Object
    Dim lobjTemp As Object
    Dim lobjResult As Object
    
    Set lobjTemp = CreateObject("职业病史录入.clsCareerHstRegt")
    
    '营养
    Set lobjRec = lobjTemp.func获取体检项目编号("营养", "13")
    Set lobjResult = lobjTemp.func获取单人单项体检结果(paraSysNo, lobjRec("编码"))
    If Not lobjResult Is Nothing Then
        If lobjResult.RecordCount > 0 Then
            comb营养.Enabled = True
        Else
            comb营养.Enabled = False
        End If
        If IsNull(lobjResult("体检结果")) = False Then comb营养.Text = lobjResult("体检结果")
    Else
        comb营养.Enabled = False
    End If
    
    '身高
    Set lobjRec = lobjTemp.func获取体检项目编号("身高", "13")
    Set lobjResult = lobjTemp.func获取单人单项体检结果(paraSysNo, lobjRec("编码"))
    If Not lobjResult Is Nothing Then
        If lobjResult.RecordCount > 0 Then
            ctxt身高.Enabled = True
        Else
            ctxt身高.Enabled = False
        End If
        If IsNull(lobjResult("体检结果")) = False Then ctxt身高.Text = lobjResult("体检结果")
    Else
        ctxt身高.Enabled = False
    End If
    
    '体重
    Set lobjRec = lobjTemp.func获取体检项目编号("体重", "13")
    Set lobjResult = lobjTemp.func获取单人单项体检结果(paraSysNo, lobjRec("编码"))
    If Not lobjResult Is Nothing Then
        If lobjResult.RecordCount > 0 Then
            ctxt体重.Enabled = True
        Else
            ctxt体重.Enabled = False
        End If
        If IsNull(lobjResult("体检结果")) = False Then ctxt体重.Text = lobjResult("体检结果")
    Else
        ctxt体重.Enabled = False
    End If
    
    '体形
    Set lobjRec = lobjTemp.func获取体检项目编号("体形", "13")
    Set lobjResult = lobjTemp.func获取单人单项体检结果(paraSysNo, lobjRec("编码"))
    If Not lobjResult Is Nothing Then
        If lobjResult.RecordCount > 0 Then
            ctxt体形.Enabled = True
        Else
            ctxt体重.Enabled = False
        End If
        If IsNull(lobjResult("体检结果")) = False Then ctxt体形.Text = lobjResult("体检结果")
    Else
        ctxt体形.Enabled = False
    End If
    
    '体重指数
    Set lobjRec = lobjTemp.func获取体检项目编号("体重指数", "13")
    Set lobjResult = lobjTemp.func获取单人单项体检结果(paraSysNo, lobjRec("编码"))
    If Not lobjResult Is Nothing Then
        If lobjResult.RecordCount > 0 Then
            ctxt体重指数.Enabled = True
        Else
            ctxt体重指数.Enabled = False
        End If
        If IsNull(lobjResult("体检结果")) = False Then ctxt体重指数.Text = lobjResult("体检结果")
    Else
        ctxt体重指数.Enabled = False
    End If
    
    '收缩压
    Set lobjRec = lobjTemp.func获取体检项目编号("收缩压", "13")
    Set lobjResult = lobjTemp.func获取单人单项体检结果(paraSysNo, lobjRec("编码"))
    If Not lobjResult Is Nothing Then
        If lobjResult.RecordCount > 0 Then
            ctxt收缩压.Enabled = True
        Else
            ctxt收缩压.Enabled = False
        End If
        If IsNull(lobjResult("体检结果")) = False Then ctxt收缩压.Text = lobjResult("体检结果")
    Else
        ctxt收缩压.Enabled = False
    End If
    
    '舒张压
    Set lobjRec = lobjTemp.func获取体检项目编号("舒张压", "13")
    Set lobjResult = lobjTemp.func获取单人单项体检结果(paraSysNo, lobjRec("编码"))
    If Not lobjResult Is Nothing Then
        If lobjResult.RecordCount > 0 Then
            ctxt舒张压.Enabled = True
        Else
            ctxt舒张压.Enabled = False
        End If
        If IsNull(lobjResult("体检结果")) = False Then ctxt舒张压.Text = lobjResult("体检结果")
    Else
        ctxt舒张压.Enabled = False
    End If
'    '增加心率显示 2015-7-1 by lanchao
'    Dim xls As Object
'    Set xls = dafuncGetData("select 体检结果 from 职业病体检_结果信息_内科 where 体检项目='02002' and 系统编号='" & paraSysNo & "'")
'    If Not IsNull(xls("体检结果")) Then ctxtxinlv.Text = xls("体检结果")
   
    '修改上面的增加心率显示  2015-12-2 by 牟俊
    Dim xls As Object
    Set xls = dafuncGetData("select 体检结果 from 职业病体检_结果信息_内科 where 体检项目='02002' and 系统编号='" & paraSysNo & "'")
    If xls.RecordCount > 0 And Not IsNull(xls("体检结果")) Then
     ctxtxinlv.Text = xls("体检结果")
     Else
     ctxtxinlv.Text = ""
    End If
'    '增加发育显示 2015-11-26 by 牟俊
    Dim ttype As Object
    Dim fy As Object
    Set ttype = dafuncGetData("select 体检表类型 from 职业病体检_体检人员基本信息表 where 系统编号='" & paraSysNo & "'")
    If ttype.RecordCount > 0 And ttype("体检表类型") = "放射健康" Then
        Set fy = dafuncGetData("select 体检结果 from 职业病体检_结果信息_内科 where 体检项目='02019' and 系统编号='" & paraSysNo & "'")
        If fy.RecordCount > 0 And Not IsNull(fy("体检结果")) Then
        comb发育.Text = fy("体检结果")
        Else
        comb发育.Text = ""
        End If
    End If
End Sub

'2012-06-13 于登淼 ↓
'保存体格一般情况界面部分(体格一般情况属于体检的内科)
Sub subSave体格一般情况(ByVal paraSysNo As String)
    Dim lobjRec As Object
    Dim lobjTemp As Object
    Dim lstr身高 As String
    Dim lstr体重 As String
    Dim lstr收缩压 As String
    Dim lstr舒张压 As String
    Dim lstr发育 As String
    Dim lstr体形 As String
    Dim lstr体重指数 As String
    '2015-7-1 modify by lanchao  一般体检数据不带单位保存
    'lstr身高 = ctxt身高.Text & " cm"
    lstr身高 = ctxt身高.Text
    lstr体重 = ctxt体重.Text
    lstr收缩压 = ctxt收缩压.Text
    lstr舒张压 = ctxt舒张压.Text
    lstr发育 = comb发育.Text   '2015-11-26 by 牟俊
    lstr体形 = ctxt体形.Text
    lstr体重指数 = ctxt体重指数.Text
    Set lobjTemp = CreateObject("职业病史录入.clsCareerHstRegt")
       
    
    '营养
    Set lobjRec = lobjTemp.func获取体检项目编号("营养", "13")
    lobjTemp.func保存单人单项体检结果 paraSysNo, "13", lobjRec("编码"), comb营养.Text
    
    '身高
    Set lobjRec = lobjTemp.func获取体检项目编号("身高", "13")
    lobjTemp.func保存单人单项体检结果 paraSysNo, "13", lobjRec("编码"), lstr身高

    '体重
    Set lobjRec = lobjTemp.func获取体检项目编号("体重", "13")
    lobjTemp.func保存单人单项体检结果 paraSysNo, "13", lobjRec("编码"), lstr体重
    
    '体形
    Set lobjRec = lobjTemp.func获取体检项目编号("体形", "13")
    lobjTemp.func保存单人单项体检结果 paraSysNo, "13", lobjRec("编码"), lstr体形
    
    '体重指数
    Set lobjRec = lobjTemp.func获取体检项目编号("体重指数", "13")
    lobjTemp.func保存单人单项体检结果 paraSysNo, "13", lobjRec("编码"), lstr体重指数
    
    '收缩压
    Set lobjRec = lobjTemp.func获取体检项目编号("收缩压", "13")
    lobjTemp.func保存单人单项体检结果 paraSysNo, "13", lobjRec("编码"), lstr收缩压
    
    '舒张压
    Set lobjRec = lobjTemp.func获取体检项目编号("舒张压", "13")
    lobjTemp.func保存单人单项体检结果 paraSysNo, "13", lobjRec("编码"), lstr舒张压
    
    '发育   2015-11-26 by 牟俊
    Dim teststylelob As Object
    Dim lob As Object
    Dim teststyle As String
    Set teststylelob = dafuncGetData("select 体检类型 from 职业病体检_体检基本信息表 where 系统编号='" & Trim(ctxtsysno.Text) & "'")
    teststyle = teststylelob("体检类型")
    If teststyle = "放射健康" Then
        Set lob = dafuncGetData("select 体检项目 from 职业病体检_结果信息_内科  where 体检项目='02019' and 系统编号='" & Trim(ctxtsysno.Text) & "'")
        If lob.RecordCount < 1 Then
    '    dafuncGetData ("insert into 职业病体检_结果信息_内科 values('" & Trim(ctxtsysno.Text) & "','02019','" & comb发育.Text & "','" & um用户编号 & "' ,'" & Now & "','" & conclusion & "')")
        dafuncGetData ("insert into 职业病体检_结果信息_内科( 系统编号,体检项目,体检结果,体检医师,填写时间) values('" & Trim(ctxtsysno.Text) & "','02019','" & comb发育.Text & "','" & um用户编号 & "','" & Now & "')")
        Else
        dafuncGetData ("update 职业病体检_结果信息_内科 set 体检结果='" & comb发育.Text & "' ,体检医师='" & um用户编号 & "' ,填写时间='" & Now & "' where 体检项目='02019' and 系统编号='" & Trim(ctxtsysno.Text) & "'")
        End If
    End If
End Sub

'2012-06-13 于登淼 ↓
'删除体格一般情况界面部分(体格一般情况属于体检的内科)
Sub subDel体格一般情况(ByVal paraSysNo As String)
    Dim lobjRec As Object
    Dim lobjTemp As Object
    Set lobjTemp = CreateObject("职业病史录入.clsCareerHstRegt")
    
    '营养
    Set lobjRec = lobjTemp.func获取体检项目编号("营养", "13")
    lobjTemp.func删除单人单项体检结果 paraSysNo, "13", lobjRec("编码")
    
    '身高
    Set lobjRec = lobjTemp.func获取体检项目编号("身高", "13")
    lobjTemp.func删除单人单项体检结果 paraSysNo, "13", lobjRec("编码")

    '体重
    Set lobjRec = lobjTemp.func获取体检项目编号("体重", "13")
    lobjTemp.func删除单人单项体检结果 paraSysNo, "13", lobjRec("编码")
    
    '收缩压
    Set lobjRec = lobjTemp.func获取体检项目编号("收缩压", "13")
    lobjTemp.func删除单人单项体检结果 paraSysNo, "13", lobjRec("编码")
    
    '舒张压
    Set lobjRec = lobjTemp.func获取体检项目编号("舒张压", "13")
    lobjTemp.func删除单人单项体检结果 paraSysNo, "13", lobjRec("编码")
    
    subClear体格一般情况
End Sub
Private Sub sub连接终端()
      With MSComm1
       
        .CommPort = 1
        .Settings = "4800,N,8,1"
        .InBufferSize = 1024 '原来为19
        .RThreshold = 1      '接收1字节触发oncomm事件
        .InputMode = comInputModeBinary
        .InputLen = 1 '输入长度为19
        .InBufferCount = 0      '清除接收缓冲区
    End With
        '打开端口
    If MSComm1.PortOpen Then MSComm1.PortOpen = False
                
'                MSComm1.CommPort = 7  '假定是用COM5口
                MSComm1.CommPort = 1
                
                ' 设定传输速率等，可依照您的需求更改
                MSComm1.Settings = "4800,N,8,1"
    
                MSComm1.PortOpen = True
                Text临时.Text = ""
End Sub
