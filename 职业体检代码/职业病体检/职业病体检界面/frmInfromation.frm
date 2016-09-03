VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmInfromation 
   Caption         =   "个人信息"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13320
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   13320
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame8 
      Caption         =   "个人信息"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12375
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   4320
         ScaleHeight     =   1785
         ScaleWidth      =   1545
         TabIndex        =   1
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label电话 
         Height          =   375
         Left            =   9480
         TabIndex        =   235
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label38 
         Caption         =   "电话："
         Height          =   255
         Left            =   9360
         TabIndex        =   234
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Lab编号 
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
         TabIndex        =   141
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label危害因素 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7200
         TabIndex        =   140
         Top             =   1320
         Width           =   90
      End
      Begin VB.Label Label30 
         Caption         =   "危害  因素："
         Height          =   255
         Index           =   1
         Left            =   6000
         TabIndex        =   139
         Top             =   1320
         Width           =   1095
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
         TabIndex        =   16
         Top             =   960
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
         TabIndex        =   15
         Top             =   840
         Width           =   3015
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
         TabIndex        =   14
         Top             =   360
         Width           =   570
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "现工作单位："
         Height          =   255
         Left            =   6000
         TabIndex        =   13
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label lab单位 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   7200
         TabIndex        =   12
         Top             =   240
         Width           =   90
      End
      Begin VB.Label Label29 
         Caption         =   "现  工  种："
         Height          =   255
         Left            =   6000
         TabIndex        =   11
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label30 
         Caption         =   "现  职  务："
         Height          =   255
         Index           =   0
         Left            =   6000
         TabIndex        =   10
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Lab现工种 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7200
         TabIndex        =   9
         Top             =   600
         Width           =   90
      End
      Begin VB.Label Lab现职务 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7200
         TabIndex        =   8
         Top             =   960
         Width           =   90
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
         TabIndex        =   7
         Top             =   1440
         Width           =   615
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
         TabIndex        =   6
         Top             =   1440
         Width           =   615
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
         TabIndex        =   5
         Top             =   1560
         Width           =   570
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
         TabIndex        =   4
         Top             =   1560
         Width           =   570
      End
      Begin VB.Label Label70 
         Caption         =   "体检  类型："
         Height          =   255
         Left            =   6000
         TabIndex        =   3
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label体检类型 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7200
         TabIndex        =   2
         Top             =   1680
         Width           =   90
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5895
      Left            =   120
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2160
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   10398
      _Version        =   393216
      Tabs            =   6
      Tab             =   1
      TabsPerRow      =   6
      TabHeight       =   520
      BackColor       =   16777215
      ForeColor       =   16711680
      TabCaption(0)   =   "个人生活史"
      TabPicture(0)   =   "frmInfromation.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "freNuclear"
      Tab(0).Control(1)=   "freRadiation"
      Tab(0).Control(2)=   "freOrdinary"
      Tab(0).Control(3)=   "ctxtOther"
      Tab(0).Control(4)=   "Frame1"
      Tab(0).Control(5)=   "Frame6"
      Tab(0).Control(6)=   "Label5(0)"
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "职业史"
      TabPicture(1)   =   "frmInfromation.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Lab职业史"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cgrd职业史"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "既往病史(包括职业病史)"
      TabPicture(2)   =   "frmInfromation.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Lab病史"
      Tab(2).Control(1)=   "cgrd病史"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "自觉症状"
      TabPicture(3)   =   "frmInfromation.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cgrd症状"
      Tab(3).Control(1)=   "Lab症状"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "体格一般情况"
      TabPicture(4)   =   "frmInfromation.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame5"
      Tab(4).Control(1)=   "Frame7"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "病状询问"
      TabPicture(5)   =   "frmInfromation.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "cgrdzzxw"
      Tab(5).ControlCount=   1
      Begin VB.Frame freNuclear 
         BorderStyle     =   0  'None
         Height          =   3615
         Left            =   -74760
         TabIndex        =   18
         Top             =   600
         Width           =   10815
         Begin VB.Frame Frame17 
            Caption         =   "烟酒史"
            ForeColor       =   &H000080FF&
            Height          =   2295
            Left            =   5880
            TabIndex        =   39
            Top             =   1320
            Width           =   5055
            Begin VB.Label Label40 
               Caption         =   "年"
               Height          =   255
               Left            =   1800
               TabIndex        =   229
               Top             =   2040
               Width           =   375
            End
            Begin VB.Label Lab戒烟时长 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   960
               TabIndex        =   228
               Top             =   2040
               Width           =   855
            End
            Begin VB.Label Lab吸烟程度 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   3360
               TabIndex        =   227
               Top             =   1680
               Width           =   975
            End
            Begin VB.Label Lab饮酒程度 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   960
               TabIndex        =   226
               Top             =   1680
               Width           =   855
            End
            Begin VB.Label Label28 
               Caption         =   "戒烟时长："
               Height          =   255
               Left            =   120
               TabIndex        =   225
               Top             =   2040
               Width           =   975
            End
            Begin VB.Label Label26 
               Caption         =   "吸烟程度："
               Height          =   255
               Left            =   2520
               TabIndex        =   224
               Top             =   1680
               Width           =   1095
            End
            Begin VB.Label Label25 
               Caption         =   "饮酒程度："
               Height          =   255
               Left            =   120
               TabIndex        =   223
               Top             =   1680
               Width           =   975
            End
            Begin VB.Label Lab烟龄 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   3240
               TabIndex        =   212
               Top             =   1320
               Width           =   975
            End
            Begin VB.Label Lab酒龄 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   840
               TabIndex        =   211
               Top             =   1320
               Width           =   975
            End
            Begin VB.Label Lab吸烟量 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   3240
               TabIndex        =   210
               Top             =   960
               Width           =   975
            End
            Begin VB.Label Lab饮酒量 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   840
               TabIndex        =   209
               Top             =   960
               Width           =   975
            End
            Begin VB.Label Lab饮食习惯 
               BackColor       =   &H00FFFFFF&
               Height          =   375
               Index           =   1
               Left            =   120
               TabIndex        =   208
               Top             =   480
               Width           =   4455
            End
            Begin VB.Label Label91 
               AutoSize        =   -1  'True
               Caption         =   "吸烟量："
               Height          =   180
               Left            =   2520
               TabIndex        =   48
               Top             =   960
               Width           =   720
            End
            Begin VB.Label Label90 
               AutoSize        =   -1  'True
               Caption         =   "饮酒量："
               Height          =   180
               Left            =   120
               TabIndex        =   47
               Top             =   960
               Width           =   720
            End
            Begin VB.Label Label83 
               AutoSize        =   -1  'True
               Caption         =   "支/天"
               Height          =   180
               Left            =   4320
               TabIndex        =   46
               Top             =   960
               Width           =   450
            End
            Begin VB.Label Label87 
               AutoSize        =   -1  'True
               Caption         =   "ML/日"
               Height          =   180
               Left            =   1920
               TabIndex        =   45
               Top             =   960
               Width           =   450
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               Caption         =   "多年居住地区、饮食习惯、烟酒嗜好用量："
               Height          =   180
               Index           =   0
               Left            =   120
               TabIndex        =   44
               Top             =   240
               Width           =   3420
            End
            Begin VB.Label Label88 
               Caption         =   "酒龄："
               Height          =   255
               Left            =   120
               TabIndex        =   43
               Top             =   1320
               Width           =   720
            End
            Begin VB.Label Label92 
               Caption         =   "年"
               Height          =   255
               Left            =   1920
               TabIndex        =   42
               Top             =   1320
               Width           =   495
            End
            Begin VB.Label Label93 
               Caption         =   "烟龄："
               Height          =   255
               Left            =   2520
               TabIndex        =   41
               Top             =   1320
               Width           =   735
            End
            Begin VB.Label Label94 
               Caption         =   "年"
               Height          =   255
               Left            =   4320
               TabIndex        =   40
               Top             =   1320
               Width           =   495
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "生育史(或配偶生育史)"
            ForeColor       =   &H000080FF&
            Height          =   2295
            Index           =   1
            Left            =   120
            TabIndex        =   25
            Top             =   1320
            Width           =   5775
            Begin VB.Label Lab子女健康 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   3720
               TabIndex        =   207
               Top             =   2040
               Width           =   1335
            End
            Begin VB.Label Lab女孩出生日期 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   2280
               TabIndex        =   206
               Top             =   2040
               Width           =   1335
            End
            Begin VB.Label Lab男孩出生日期 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   2280
               TabIndex        =   205
               Top             =   1680
               Width           =   1335
            End
            Begin VB.Label Lab女孩 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   960
               TabIndex        =   204
               Top             =   2040
               Width           =   375
            End
            Begin VB.Label Lab子女数 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   960
               TabIndex        =   203
               Top             =   1680
               Width           =   375
            End
            Begin VB.Label Lab不孕原因 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   960
               TabIndex        =   202
               Top             =   1200
               Width           =   4215
            End
            Begin VB.Label Lab流产 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   201
               Top             =   1200
               Width           =   615
            End
            Begin VB.Label Lab畸胎 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   4560
               TabIndex        =   200
               Top             =   600
               Width           =   615
            End
            Begin VB.Label Lab多胎 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   3480
               TabIndex        =   199
               Top             =   600
               Width           =   615
            End
            Begin VB.Label Lab死产 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   2640
               TabIndex        =   198
               Top             =   600
               Width           =   615
            End
            Begin VB.Label Lab早产 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   1800
               TabIndex        =   197
               Top             =   600
               Width           =   615
            End
            Begin VB.Label Lab活产 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   960
               TabIndex        =   196
               Top             =   600
               Width           =   615
            End
            Begin VB.Label Lab孕次 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   195
               Top             =   600
               Width           =   615
            End
            Begin VB.Label Label89 
               AutoSize        =   -1  'True
               Caption         =   "子女健康状况："
               Height          =   300
               Left            =   3600
               TabIndex        =   38
               Top             =   1680
               Width           =   1260
            End
            Begin VB.Label Label86 
               AutoSize        =   -1  'True
               Caption         =   "畸胎："
               Height          =   180
               Left            =   4560
               TabIndex        =   37
               Top             =   360
               Width           =   540
            End
            Begin VB.Label Label85 
               AutoSize        =   -1  'True
               Caption         =   "多胎："
               Height          =   180
               Left            =   3480
               TabIndex        =   36
               Top             =   360
               Width           =   540
            End
            Begin VB.Label Label84 
               AutoSize        =   -1  'True
               Caption         =   "孕次："
               Height          =   180
               Left            =   120
               TabIndex        =   35
               Top             =   360
               Width           =   540
            End
            Begin VB.Label Label82 
               AutoSize        =   -1  'True
               Caption         =   "活产："
               Height          =   180
               Left            =   960
               TabIndex        =   34
               Top             =   360
               Width           =   540
            End
            Begin VB.Label Label81 
               AutoSize        =   -1  'True
               Caption         =   "早产："
               Height          =   180
               Left            =   1800
               TabIndex        =   33
               Top             =   360
               Width           =   540
            End
            Begin VB.Label Label80 
               AutoSize        =   -1  'True
               Caption         =   "死产："
               Height          =   180
               Left            =   2640
               TabIndex        =   32
               Top             =   360
               Width           =   540
            End
            Begin VB.Label Label79 
               AutoSize        =   -1  'True
               Caption         =   "不孕不育原因："
               Height          =   180
               Left            =   960
               TabIndex        =   31
               Top             =   960
               Width           =   1260
            End
            Begin VB.Label Label75 
               Caption         =   "现有男孩："
               Height          =   255
               Left            =   120
               TabIndex        =   30
               Top             =   1680
               Width           =   975
            End
            Begin VB.Label Label76 
               Caption         =   "现有女孩："
               Height          =   255
               Left            =   120
               TabIndex        =   29
               Top             =   2040
               Width           =   975
            End
            Begin VB.Label Label77 
               Caption         =   "出生日期："
               Height          =   225
               Left            =   1440
               TabIndex        =   28
               Top             =   1680
               Width           =   975
            End
            Begin VB.Label Label78 
               Caption         =   "出生日期："
               Height          =   225
               Left            =   1440
               TabIndex        =   27
               Top             =   2040
               Width           =   975
            End
            Begin VB.Label Label95 
               Caption         =   "流产："
               Height          =   255
               Left            =   120
               TabIndex        =   26
               Top             =   960
               Width           =   855
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "婚姻史"
            ForeColor       =   &H000080FF&
            Height          =   1095
            Index           =   1
            Left            =   120
            TabIndex        =   19
            Top             =   120
            Width           =   10935
            Begin VB.Label Lab配偶放射 
               BackColor       =   &H00FFFFFF&
               Height          =   375
               Index           =   1
               Left            =   5880
               TabIndex        =   194
               Top             =   600
               Width           =   4335
            End
            Begin VB.Label Lab配偶职业 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   3600
               TabIndex        =   193
               Top             =   720
               Width           =   1455
            End
            Begin VB.Label Lab配偶健康 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   3960
               TabIndex        =   192
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Lab婚期 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   960
               TabIndex        =   191
               Top             =   720
               Width           =   1095
            End
            Begin VB.Label Lab婚否 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   960
               TabIndex        =   190
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Lab 
               AutoSize        =   -1  'True
               Caption         =   "配偶健康状况："
               Height          =   180
               Index           =   1
               Left            =   2760
               TabIndex        =   24
               Top             =   300
               Width           =   1260
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "配偶接触放射线情况："
               Height          =   180
               Index           =   1
               Left            =   5880
               TabIndex        =   23
               Top             =   240
               Width           =   1800
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "配偶职业："
               Height          =   180
               Index           =   1
               Left            =   2760
               TabIndex        =   22
               Top             =   760
               Width           =   900
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "结婚日期："
               Height          =   180
               Index           =   1
               Left            =   120
               TabIndex        =   21
               Top             =   760
               Width           =   900
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "是否结婚："
               Height          =   180
               Index           =   1
               Left            =   120
               TabIndex        =   20
               Top             =   300
               Width           =   900
            End
         End
      End
      Begin VB.Frame freRadiation 
         BorderStyle     =   0  'None
         Height          =   3735
         Left            =   -74760
         TabIndex        =   72
         Top             =   600
         Width           =   11895
         Begin VB.Frame Frame3 
            Caption         =   "生育史(或配偶生育史)"
            ForeColor       =   &H000080FF&
            Height          =   2295
            Index           =   0
            Left            =   120
            TabIndex        =   95
            Top             =   1440
            Width           =   5775
            Begin VB.Label Lab子女健康 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   4080
               TabIndex        =   161
               Top             =   1920
               Width           =   1215
            End
            Begin VB.Label Lab女孩出生日期 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   2760
               TabIndex        =   160
               Top             =   1920
               Width           =   1215
            End
            Begin VB.Label Lab男孩出生日期 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   2760
               TabIndex        =   159
               Top             =   1560
               Width           =   1215
            End
            Begin VB.Label Lab女孩 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   1080
               TabIndex        =   158
               Top             =   1920
               Width           =   735
            End
            Begin VB.Label Lab子女数 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   1080
               TabIndex        =   157
               Top             =   1560
               Width           =   735
            End
            Begin VB.Label Lab不孕原因 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   2040
               TabIndex        =   156
               Top             =   1080
               Width           =   3255
            End
            Begin VB.Label Lab异位妊娠 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   1080
               TabIndex        =   155
               Top             =   1080
               Width           =   735
            End
            Begin VB.Label Lab多胎 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   154
               Top             =   1080
               Width           =   735
            End
            Begin VB.Label Lab畸胎 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   4560
               TabIndex        =   153
               Top             =   480
               Width           =   735
            End
            Begin VB.Label Lab流产 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   3720
               TabIndex        =   152
               Top             =   480
               Width           =   735
            End
            Begin VB.Label Lab死产 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   2880
               TabIndex        =   151
               Top             =   480
               Width           =   735
            End
            Begin VB.Label Lab早产 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   1920
               TabIndex        =   150
               Top             =   480
               Width           =   735
            End
            Begin VB.Label Lab活产 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   1080
               TabIndex        =   149
               Top             =   480
               Width           =   735
            End
            Begin VB.Label Lab孕次 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   147
               Top             =   480
               Width           =   735
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               Caption         =   "子女健康状况："
               Height          =   180
               Index           =   0
               Left            =   4080
               TabIndex        =   109
               Top             =   1560
               Width           =   1260
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               Caption         =   "现有男孩："
               Height          =   180
               Index           =   0
               Left            =   240
               TabIndex        =   108
               Top             =   1560
               Width           =   900
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               Caption         =   "流产："
               Height          =   180
               Index           =   0
               Left            =   3720
               TabIndex        =   107
               Top             =   240
               Width           =   540
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               Caption         =   "畸胎："
               Height          =   180
               Index           =   0
               Left            =   4560
               TabIndex        =   106
               Top             =   240
               Width           =   540
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               Caption         =   "多胎："
               Height          =   180
               Index           =   0
               Left            =   240
               TabIndex        =   105
               Top             =   840
               Width           =   540
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "孕次："
               Height          =   180
               Index           =   0
               Left            =   240
               TabIndex        =   104
               Top             =   240
               Width           =   540
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               Caption         =   "异位妊娠："
               Height          =   180
               Index           =   0
               Left            =   1080
               TabIndex        =   103
               Top             =   840
               Width           =   900
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "活产："
               Height          =   180
               Index           =   0
               Left            =   1080
               TabIndex        =   102
               Top             =   240
               Width           =   540
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "早产："
               Height          =   180
               Index           =   0
               Left            =   1920
               TabIndex        =   101
               Top             =   240
               Width           =   540
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "死产："
               Height          =   180
               Index           =   0
               Left            =   2880
               TabIndex        =   100
               Top             =   240
               Width           =   540
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               Caption         =   "不孕不育原因："
               Height          =   180
               Index           =   0
               Left            =   2040
               TabIndex        =   99
               Top             =   840
               Width           =   1500
            End
            Begin VB.Label Label71 
               Caption         =   "出生日期："
               Height          =   255
               Left            =   1920
               TabIndex        =   98
               Top             =   1560
               Width           =   975
            End
            Begin VB.Label Label72 
               Caption         =   "现有女孩："
               Height          =   255
               Left            =   240
               TabIndex        =   97
               Top             =   1920
               Width           =   975
            End
            Begin VB.Label Label73 
               Caption         =   "出生日期："
               Height          =   255
               Left            =   1920
               TabIndex        =   96
               Top             =   1920
               Width           =   975
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "婚姻史"
            ForeColor       =   &H000080FF&
            Height          =   1095
            Index           =   0
            Left            =   120
            TabIndex        =   88
            Top             =   120
            Width           =   11055
            Begin MSComCtl2.DTPicker ctxtmarrydate1 
               CausesValidation=   0   'False
               Height          =   300
               Index           =   0
               Left            =   7560
               TabIndex        =   89
               Top             =   120
               Visible         =   0   'False
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   0   'False
               CustomFormat    =   "yyyy/MM"
               Format          =   60227584
               CurrentDate     =   41013
            End
            Begin VB.Label Lab配偶放射 
               BackColor       =   &H00FFFFFF&
               Height          =   495
               Index           =   0
               Left            =   5880
               TabIndex        =   146
               Top             =   480
               Width           =   4815
            End
            Begin VB.Label Lab配偶职业 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   3960
               TabIndex        =   145
               Top             =   720
               Width           =   1575
            End
            Begin VB.Label Lab配偶健康 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   3960
               TabIndex        =   144
               Top             =   240
               Width           =   1575
            End
            Begin VB.Label Lab婚期 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   960
               TabIndex        =   143
               Top             =   720
               Width           =   1575
            End
            Begin VB.Label Lab婚否 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   960
               TabIndex        =   142
               Top             =   240
               Width           =   1575
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "是否结婚："
               Height          =   180
               Index           =   0
               Left            =   120
               TabIndex        =   94
               Top             =   300
               Width           =   900
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "结婚日期："
               Height          =   180
               Index           =   0
               Left            =   120
               TabIndex        =   93
               Top             =   760
               Width           =   900
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "配偶职业："
               Height          =   180
               Index           =   0
               Left            =   2760
               TabIndex        =   92
               Top             =   760
               Width           =   900
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "配偶接触放射线情况："
               Height          =   180
               Index           =   0
               Left            =   5880
               TabIndex        =   91
               Top             =   240
               Width           =   1800
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "配偶健康状况："
               Height          =   180
               Index           =   0
               Left            =   2760
               TabIndex        =   90
               Top             =   300
               Width           =   1260
            End
         End
         Begin VB.ComboBox Combo11 
            Height          =   300
            ItemData        =   "frmInfromation.frx":00A8
            Left            =   5040
            List            =   "frmInfromation.frx":00B5
            TabIndex        =   87
            Text            =   "健康"
            Top             =   360
            Width           =   855
         End
         Begin VB.Frame Frame19 
            Caption         =   "烟酒史"
            ForeColor       =   &H000080FF&
            Height          =   2175
            Index           =   0
            Left            =   6000
            TabIndex        =   73
            Top             =   1560
            Width           =   5175
            Begin VB.Label Lab饮食习惯 
               BackColor       =   &H00FFFFFF&
               Height          =   375
               Index           =   0
               Left            =   120
               TabIndex        =   175
               Top             =   1680
               Width           =   4215
            End
            Begin VB.Label Lab饮酒量 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   3360
               TabIndex        =   174
               Top             =   840
               Width           =   975
            End
            Begin VB.Label Lab吸烟量 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   960
               TabIndex        =   173
               Top             =   1200
               Width           =   975
            End
            Begin VB.Label Lab戒烟时长 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   960
               TabIndex        =   172
               Top             =   840
               Width           =   975
            End
            Begin VB.Label Lab酒龄 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   3360
               TabIndex        =   171
               Top             =   480
               Width           =   975
            End
            Begin VB.Label Lab烟龄 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   960
               TabIndex        =   169
               Top             =   480
               Width           =   975
            End
            Begin VB.Label Lab饮酒程度 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   3360
               TabIndex        =   168
               Top             =   120
               Width           =   975
            End
            Begin VB.Label Lab吸烟程度 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   960
               TabIndex        =   167
               Top             =   120
               Width           =   975
            End
            Begin VB.Label Label99 
               AutoSize        =   -1  'True
               Caption         =   "吸烟程度："
               Height          =   180
               Index           =   0
               Left            =   120
               TabIndex        =   86
               Top             =   195
               Width           =   900
            End
            Begin VB.Label Label100 
               AutoSize        =   -1  'True
               Caption         =   "饮酒程度："
               Height          =   180
               Index           =   0
               Left            =   2520
               TabIndex        =   85
               Top             =   195
               Width           =   900
            End
            Begin VB.Label Label102 
               AutoSize        =   -1  'True
               Caption         =   "支/天"
               Height          =   180
               Index           =   0
               Left            =   2040
               TabIndex        =   84
               Top             =   1200
               Width           =   450
            End
            Begin VB.Label Label103 
               AutoSize        =   -1  'True
               Caption         =   "饮酒量："
               Height          =   180
               Index           =   0
               Left            =   2640
               TabIndex        =   83
               Top             =   915
               Width           =   720
            End
            Begin VB.Label Label104 
               AutoSize        =   -1  'True
               Caption         =   "吸烟量："
               Height          =   180
               Index           =   0
               Left            =   240
               TabIndex        =   82
               Top             =   1215
               Width           =   720
            End
            Begin VB.Label Label105 
               AutoSize        =   -1  'True
               Caption         =   "酒龄："
               Height          =   180
               Index           =   0
               Left            =   2640
               TabIndex        =   81
               Top             =   540
               Width           =   540
            End
            Begin VB.Label Label106 
               AutoSize        =   -1  'True
               Caption         =   "烟龄："
               Height          =   180
               Index           =   0
               Left            =   360
               TabIndex        =   80
               Top             =   540
               Width           =   540
            End
            Begin VB.Label Label107 
               AutoSize        =   -1  'True
               Caption         =   "戒烟时长："
               Height          =   180
               Index           =   0
               Left            =   120
               TabIndex        =   79
               Top             =   855
               Width           =   900
            End
            Begin VB.Label Label108 
               AutoSize        =   -1  'True
               Caption         =   "年"
               Height          =   180
               Index           =   0
               Left            =   2040
               TabIndex        =   78
               Top             =   480
               Width           =   180
            End
            Begin VB.Label Label109 
               AutoSize        =   -1  'True
               Caption         =   "年"
               Height          =   180
               Index           =   0
               Left            =   4440
               TabIndex        =   77
               Top             =   480
               Width           =   180
            End
            Begin VB.Label Label110 
               AutoSize        =   -1  'True
               Caption         =   "年"
               Height          =   180
               Index           =   0
               Left            =   2040
               TabIndex        =   76
               Top             =   840
               Width           =   180
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               Caption         =   "多年居住地区、饮食习惯、烟酒嗜好用量："
               Height          =   180
               Index           =   1
               Left            =   120
               TabIndex        =   75
               Top             =   1440
               Width           =   3420
            End
            Begin VB.Label Label101 
               AutoSize        =   -1  'True
               Caption         =   "ML/日"
               Height          =   300
               Index           =   0
               Left            =   4440
               TabIndex        =   74
               Top             =   840
               Width           =   810
            End
         End
      End
      Begin VB.Frame freOrdinary 
         BorderStyle     =   0  'None
         Height          =   3735
         Left            =   -74640
         TabIndex        =   49
         Top             =   600
         Width           =   11175
         Begin VB.Frame Frame4 
            Caption         =   "过敏史"
            ForeColor       =   &H000080FF&
            Height          =   1095
            Index           =   1
            Left            =   6000
            TabIndex        =   71
            Top             =   2520
            Width           =   5055
            Begin VB.Label Lab过敏史 
               BackColor       =   &H00FFFFFF&
               Height          =   735
               Index           =   2
               Left            =   120
               TabIndex        =   189
               Top             =   240
               Width           =   4815
            End
         End
         Begin VB.Frame Frame19 
            Caption         =   "烟酒史"
            ForeColor       =   &H000080FF&
            Height          =   1455
            Index           =   1
            Left            =   6000
            TabIndex        =   58
            Top             =   240
            Width           =   5055
            Begin VB.Label Lab吸烟量 
               BackColor       =   &H00FFFFFF&
               Height          =   180
               Index           =   2
               Left            =   960
               TabIndex        =   188
               Top             =   1200
               Width           =   975
            End
            Begin VB.Label Lab饮酒量 
               BackColor       =   &H00FFFFFF&
               Height          =   180
               Index           =   2
               Left            =   3360
               TabIndex        =   187
               Top             =   960
               Width           =   975
            End
            Begin VB.Label Lab戒烟时长 
               BackColor       =   &H00FFFFFF&
               Height          =   180
               Index           =   2
               Left            =   960
               TabIndex        =   186
               Top             =   960
               Width           =   975
            End
            Begin VB.Label Lab酒龄 
               BackColor       =   &H00FFFFFF&
               Height          =   180
               Index           =   2
               Left            =   3120
               TabIndex        =   185
               Top             =   600
               Width           =   1095
            End
            Begin VB.Label Lab烟龄 
               BackColor       =   &H00FFFFFF&
               Height          =   180
               Index           =   2
               Left            =   840
               TabIndex        =   184
               Top             =   600
               Width           =   1095
            End
            Begin VB.Label Lab饮酒程度 
               BackColor       =   &H00FFFFFF&
               Height          =   180
               Index           =   2
               Left            =   3360
               TabIndex        =   183
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label Lab吸烟程度 
               BackColor       =   &H00FFFFFF&
               Height          =   180
               Index           =   2
               Left            =   960
               TabIndex        =   182
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label Label110 
               AutoSize        =   -1  'True
               Caption         =   "年"
               Height          =   180
               Index           =   1
               Left            =   2040
               TabIndex        =   70
               Top             =   960
               Width           =   180
            End
            Begin VB.Label Label109 
               AutoSize        =   -1  'True
               Caption         =   "年"
               Height          =   180
               Index           =   1
               Left            =   4440
               TabIndex        =   69
               Top             =   600
               Width           =   180
            End
            Begin VB.Label Label108 
               AutoSize        =   -1  'True
               Caption         =   "年"
               Height          =   180
               Index           =   1
               Left            =   2040
               TabIndex        =   68
               Top             =   600
               Width           =   180
            End
            Begin VB.Label Label107 
               AutoSize        =   -1  'True
               Caption         =   "戒烟时长："
               Height          =   180
               Index           =   1
               Left            =   120
               TabIndex        =   67
               Top             =   960
               Width           =   900
            End
            Begin VB.Label Label106 
               AutoSize        =   -1  'True
               Caption         =   "烟龄："
               Height          =   180
               Index           =   1
               Left            =   360
               TabIndex        =   66
               Top             =   600
               Width           =   540
            End
            Begin VB.Label Label105 
               AutoSize        =   -1  'True
               Caption         =   "酒龄："
               Height          =   180
               Index           =   1
               Left            =   2640
               TabIndex        =   65
               Top             =   600
               Width           =   540
            End
            Begin VB.Label Label104 
               AutoSize        =   -1  'True
               Caption         =   "吸烟量："
               Height          =   180
               Index           =   1
               Left            =   240
               TabIndex        =   64
               Top             =   1215
               Width           =   720
            End
            Begin VB.Label Label103 
               AutoSize        =   -1  'True
               Caption         =   "饮酒量："
               Height          =   180
               Index           =   1
               Left            =   2640
               TabIndex        =   63
               Top             =   960
               Width           =   720
            End
            Begin VB.Label Label102 
               AutoSize        =   -1  'True
               Caption         =   "支/天"
               Height          =   180
               Index           =   1
               Left            =   2040
               TabIndex        =   62
               Top             =   1200
               Width           =   450
            End
            Begin VB.Label Label101 
               AutoSize        =   -1  'True
               Caption         =   "ML/日"
               Height          =   180
               Index           =   1
               Left            =   4440
               TabIndex        =   61
               Top             =   960
               Width           =   450
            End
            Begin VB.Label Label100 
               AutoSize        =   -1  'True
               Caption         =   "饮酒程度："
               Height          =   180
               Index           =   1
               Left            =   2520
               TabIndex        =   60
               Top             =   240
               Width           =   900
            End
            Begin VB.Label Label99 
               AutoSize        =   -1  'True
               Caption         =   "吸烟程度："
               Height          =   180
               Index           =   1
               Left            =   120
               TabIndex        =   59
               Top             =   240
               Width           =   900
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "婚姻史"
            ForeColor       =   &H000080FF&
            Height          =   615
            Index           =   2
            Left            =   120
            TabIndex        =   56
            Top             =   240
            Width           =   5775
            Begin VB.Label Lab婚否 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   2
               Left            =   1320
               TabIndex        =   213
               Top             =   240
               Width           =   1695
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "是否结婚："
               Height          =   180
               Index           =   2
               Left            =   480
               TabIndex        =   57
               Top             =   300
               Width           =   900
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "生育史(或配偶生育史)"
            ForeColor       =   &H000080FF&
            Height          =   2655
            Index           =   2
            Left            =   120
            TabIndex        =   50
            Top             =   960
            Width           =   5775
            Begin VB.Label Lab异常胎 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   2
               Left            =   1680
               TabIndex        =   181
               Top             =   1080
               Width           =   1335
            End
            Begin VB.Label Lab死产 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   2
               Left            =   3960
               TabIndex        =   180
               Top             =   720
               Width           =   1335
            End
            Begin VB.Label Lab流产 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   2
               Left            =   1680
               TabIndex        =   179
               Top             =   720
               Width           =   1335
            End
            Begin VB.Label Lab早产 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   2
               Left            =   3960
               TabIndex        =   178
               Top             =   360
               Width           =   1335
            End
            Begin VB.Label Lab子女数 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   2
               Left            =   1680
               TabIndex        =   177
               Top             =   360
               Width           =   1335
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "死产："
               Height          =   180
               Index           =   1
               Left            =   3480
               TabIndex        =   55
               Top             =   720
               Width           =   540
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "早产："
               Height          =   180
               Index           =   1
               Left            =   3480
               TabIndex        =   54
               Top             =   360
               Width           =   540
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               Caption         =   "流产："
               Height          =   180
               Index           =   1
               Left            =   1080
               TabIndex        =   53
               Top             =   720
               Width           =   540
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               Caption         =   "现有子女数目："
               Height          =   180
               Index           =   1
               Left            =   480
               TabIndex        =   52
               Top             =   360
               Width           =   1260
            End
            Begin VB.Label Label异常胎 
               AutoSize        =   -1  'True
               Caption         =   "异常胎："
               Height          =   180
               Left            =   840
               TabIndex        =   51
               Top             =   1080
               Width           =   720
            End
         End
         Begin VB.Label Lab出生地 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   2
            Left            =   6120
            TabIndex        =   231
            Top             =   2040
            Width           =   4935
         End
         Begin VB.Label Label36 
            Caption         =   "出生地："
            Height          =   255
            Left            =   6120
            TabIndex        =   230
            Top             =   1800
            Width           =   975
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "体格一般情况录入"
         ForeColor       =   &H000080FF&
         Height          =   3255
         Left            =   -74280
         TabIndex        =   124
         Top             =   1140
         Width           =   3615
         Begin VB.Label Lab发育 
            BackColor       =   &H80000009&
            Height          =   255
            Left            =   960
            TabIndex        =   233
            Top             =   2760
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label Label37 
            Caption         =   "发育"
            Height          =   375
            Left            =   240
            TabIndex        =   232
            Top             =   2760
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label Lab舒张压 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   960
            TabIndex        =   218
            Top             =   2280
            Width           =   1335
         End
         Begin VB.Label Lab收缩压 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   960
            TabIndex        =   217
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label Lab体重 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   960
            TabIndex        =   216
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label Lab身高 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   960
            TabIndex        =   215
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Lab营养 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   960
            TabIndex        =   214
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label55 
            Caption         =   "mmHg"
            Height          =   255
            Left            =   2400
            TabIndex        =   133
            Top             =   1800
            Width           =   375
         End
         Begin VB.Label Label53 
            Caption         =   "收缩压"
            Height          =   255
            Left            =   240
            TabIndex        =   132
            Top             =   1800
            Width           =   615
         End
         Begin VB.Label Label51 
            Caption         =   "kg"
            Height          =   255
            Left            =   2400
            TabIndex        =   131
            Top             =   1320
            Width           =   375
         End
         Begin VB.Label Label42 
            Caption         =   "体重"
            Height          =   255
            Left            =   240
            TabIndex        =   130
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label Label35 
            Caption         =   "cm"
            Height          =   255
            Index           =   0
            Left            =   2400
            TabIndex        =   129
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label34 
            Caption         =   "身高"
            Height          =   255
            Left            =   240
            TabIndex        =   128
            Top             =   840
            Width           =   495
         End
         Begin VB.Label Label33 
            Caption         =   "营养"
            Height          =   255
            Left            =   240
            TabIndex        =   127
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label56 
            Caption         =   "mmHg"
            Height          =   255
            Left            =   2400
            TabIndex        =   126
            Top             =   2280
            Width           =   375
         End
         Begin VB.Label Label57 
            Caption         =   "舒张压"
            Height          =   375
            Left            =   240
            TabIndex        =   125
            Top             =   2280
            Width           =   615
         End
      End
      Begin VB.TextBox ctxtOther 
         Height          =   495
         Left            =   -74640
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   123
         Top             =   5400
         Width           =   10935
      End
      Begin VB.Frame Frame1 
         Caption         =   "月经史"
         ForeColor       =   &H000080FF&
         Height          =   975
         Left            =   -74640
         TabIndex        =   116
         Top             =   4320
         Width           =   5775
         Begin VB.Label Lab停经年龄 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   3240
            TabIndex        =   166
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Lab末次月经 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   960
            TabIndex        =   165
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Lab周期 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   4080
            TabIndex        =   164
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Lab经期 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   2400
            TabIndex        =   163
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Lab初潮 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   600
            TabIndex        =   162
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "初潮："
            Height          =   180
            Index           =   2
            Left            =   120
            TabIndex        =   122
            Top             =   240
            Width           =   540
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "经期："
            Height          =   180
            Index           =   2
            Left            =   1920
            TabIndex        =   121
            Top             =   240
            Width           =   540
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "周期："
            Height          =   180
            Index           =   2
            Left            =   3600
            TabIndex        =   120
            Top             =   240
            Width           =   540
         End
         Begin VB.Label Label4 
            Caption         =   "Label4"
            Height          =   15
            Index           =   2
            Left            =   720
            TabIndex        =   119
            Top             =   720
            Width           =   135
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "末次月经："
            Height          =   180
            Index           =   2
            Left            =   120
            TabIndex        =   118
            Top             =   600
            Width           =   900
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "停经年龄："
            Height          =   180
            Index           =   2
            Left            =   2400
            TabIndex        =   117
            Top             =   600
            Width           =   900
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "家族史"
         ForeColor       =   &H000080FF&
         Height          =   975
         Left            =   -68880
         TabIndex        =   114
         Top             =   4320
         Width           =   5175
         Begin VB.Label Lab家族史 
            BackColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   240
            TabIndex        =   176
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label Label27 
            Caption         =   "提示:家族中有无遗传性疾病、血液病、糖尿病、高血压病、神经精神性疾病、肿瘤、结核病等"
            Height          =   615
            Left            =   2520
            TabIndex        =   115
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "心率录入"
         ForeColor       =   &H000080FF&
         Height          =   855
         Left            =   -74280
         TabIndex        =   110
         Top             =   4800
         Width           =   5295
         Begin VB.Label Lab心率 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   840
            TabIndex        =   219
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label64 
            Caption         =   "次/分"
            Height          =   255
            Left            =   2400
            TabIndex        =   113
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label65 
            Caption         =   "心率"
            Height          =   375
            Left            =   240
            TabIndex        =   112
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label66 
            Height          =   255
            Left            =   3120
            TabIndex        =   111
            Top             =   360
            Width           =   855
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid cgrd症状 
         Height          =   2175
         Left            =   -74880
         TabIndex        =   134
         Top             =   480
         Width           =   12375
         _cx             =   21828
         _cy             =   3836
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
         FormatString    =   $"frmInfromation.frx":00CB
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
         Height          =   4335
         Left            =   -74880
         TabIndex        =   135
         Top             =   480
         Width           =   12375
         _cx             =   21828
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
      Begin VSFlex8Ctl.VSFlexGrid cgrdzzxw 
         Height          =   4335
         Left            =   -74880
         TabIndex        =   136
         Top             =   480
         Width           =   12255
         _cx             =   21616
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
         FormatString    =   $"frmInfromation.frx":0162
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
         Height          =   4815
         Left            =   120
         TabIndex        =   137
         Top             =   480
         Width           =   12375
         _cx             =   21828
         _cy             =   8493
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
         FormatString    =   $"frmInfromation.frx":01E2
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
      Begin VB.Label Lab症状 
         BackColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   -74760
         TabIndex        =   222
         Top             =   480
         Width           =   5055
      End
      Begin VB.Label Lab病史 
         BackColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   -74760
         TabIndex        =   221
         Top             =   600
         Width           =   3975
      End
      Begin VB.Label Lab职业史 
         BackColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   120
         TabIndex        =   220
         Top             =   600
         Width           =   5415
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "其他："
         Height          =   180
         Index           =   0
         Left            =   -74640
         TabIndex        =   138
         Top             =   5460
         Width           =   540
      End
   End
   Begin VB.Label Label97 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label"
      Height          =   255
      Left            =   9720
      TabIndex        =   170
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label Label39 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label"
      Height          =   255
      Left            =   1560
      TabIndex        =   148
      Top             =   4680
      Width           =   735
   End
End
Attribute VB_Name = "frmInfromation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'窗体加载
Private Sub Form_Load()
    
    Dim Index As Integer
    Dim baseresql As Object
    Set baseresql = dafuncGetData("select * From 职业病体检_体检人员基本信息表 where 系统编号='" & mstr系统编号 & "'")
    Lab编号.Caption = baseresql("系统编号")
    Lab姓名.Caption = baseresql("姓名")
    Lab性别.Caption = baseresql("性别")
    Lab年龄.Caption = baseresql("年龄")
    lab单位.Caption = baseresql("单位名称")
    Lab现工种.Caption = baseresql("现工种")
    Lab现职务.Caption = baseresql("职务或职称")
    Label危害因素.Caption = baseresql("危害因素")
    
    
        '获取像片
    Dim lobjRec As Object
    Set lobjRec = CreateObject("职业病对象.clspersonexamed")
    lobjRec.系统编号 = Trim(Lab编号.Caption)
    Picture2.Picture = lobjRec.像片
    Picture2.Visible = True
     
    '初始化界面通用对象（每个界面对应一个界面通用对象实例，不可混用，切记切记）。
'    Set mobjGUI = New cls界面通用对象
      Dim resql As Object
    Set resql = dafuncGetData("select * From 职业病体检_体检人员基本信息表 where 系统编号='" & mstr系统编号 & "'")
    
    If resql("体检表类型") = "普通体检" Or resql("体检表类型") = "职业健康" Then
        freOrdinary.Visible = True
        freNuclear.Visible = False
        freRadiation.Visible = False
        Index = 2
        
    ElseIf resql("体检表类型") = "8023部队" Or resql("体检表类型") = "涉核部队" Or resql("体检表类型") = "涉核部队YK" Then
        freNuclear.Visible = True
        freOrdinary.Visible = False
        freRadiation.Visible = False
        Index = 1
    ElseIf resql("体检表类型") = "放射健康" Then
        freRadiation.Visible = True
        freOrdinary.Visible = False
        freNuclear.Visible = False
        Index = 0
    End If
    Label体检类型.Caption = resql("体检表类型")
    If Len(resql("电话号码")) = 11 Then
        resql("电话号码") = Left(resql("电话号码"), 3) & "-" & Mid(resql("电话号码"), 4, 4) & "-" & Mid(resql("电话号码"), 8, 4)
    End If
    Label电话.Caption = IIf(resql("电话号码") = "", "无", resql("电话号码"))
    Label电话.FontSize = 14
        
    
    '加载时，显示第一个窗口
    SSTab1.Tab = 0
    Dim detsql As Object
    Set detsql = dafuncGetData("select * From 职业病体检_个人生活史表 where 系统编号='" & mstr系统编号 & "'")
    If Index = 0 Then   '放射健康
        
    Lab婚否(Index).Caption = detsql("是否结婚")
    Lab配偶健康(Index).Caption = detsql("配偶健康状况")
    Lab婚期(Index).Caption = detsql("结婚日期")
    Lab配偶职业(Index).Caption = detsql("配偶职业")
    Lab配偶放射(Index).Caption = detsql("配偶接触放射")
    Lab孕次(Index).Caption = detsql("孕次")
    Lab活产(Index).Caption = detsql("活产")
    Lab早产(Index).Caption = detsql("早产")
    Lab死产(Index).Caption = detsql("死产")
    Lab流产(Index).Caption = detsql("自然流产")
    Lab畸胎(Index).Caption = detsql("畸胎")
    Lab多胎(Index).Caption = detsql("多胎")
    Lab异位妊娠(Index).Caption = detsql("异位妊娠")
    Lab不孕原因(Index).Caption = detsql("不孕不育原因")
    Lab子女数(Index).Caption = detsql("现有子女数目")
    Lab女孩(Index).Caption = detsql("现有女孩")
    Lab男孩出生日期(Index).Caption = detsql("男孩出生日期")
    Lab女孩出生日期(Index).Caption = detsql("女孩出生日期")
    Lab子女健康(Index).Caption = detsql("子女健康状况")
    Lab吸烟程度(Index).Caption = detsql("吸烟程度")
    Lab饮酒程度(Index).Caption = detsql("饮酒程度")
    Lab烟龄(Index).Caption = detsql("烟龄")
    Lab酒龄(Index).Caption = detsql("酒龄")
    Lab戒烟时长(Index).Caption = detsql("戒烟时长")
    Lab吸烟量(Index).Caption = detsql("吸烟量")
    Lab饮酒量(Index).Caption = detsql("饮酒量")
    Lab饮食习惯(Index).Caption = detsql("生活更多")
    End If
    
    
    If Index = 1 Then    '核部队
    Lab吸烟程度(Index).Caption = detsql("吸烟程度")
    Lab饮酒程度(Index).Caption = detsql("饮酒程度")
    Lab戒烟时长(Index).Caption = detsql("戒烟时长")
    Lab婚否(Index).Caption = detsql("是否结婚")
    Lab配偶健康(Index).Caption = detsql("配偶健康状况")
    Lab婚期(Index).Caption = detsql("结婚日期")
    Lab配偶职业(Index).Caption = detsql("配偶职业")
    Lab配偶放射(Index).Caption = detsql("配偶接触放射")
    Lab孕次(Index).Caption = detsql("孕次")
    Lab活产(Index).Caption = detsql("活产")
    Lab早产(Index).Caption = detsql("早产")
    Lab死产(Index).Caption = detsql("死产")
    Lab流产(Index).Caption = detsql("自然流产")
    Lab畸胎(Index).Caption = detsql("畸胎")
    Lab多胎(Index).Caption = detsql("多胎")
    Lab不孕原因(Index).Caption = detsql("不孕不育原因")
    Lab子女数(Index).Caption = detsql("现有子女数目")
    Lab女孩(Index).Caption = detsql("现有女孩")
    Lab男孩出生日期(Index).Caption = detsql("男孩出生日期")
    Lab女孩出生日期(Index).Caption = detsql("女孩出生日期")
    Lab子女健康(Index).Caption = detsql("子女健康状况")
    Lab烟龄(Index).Caption = detsql("烟龄")
    Lab酒龄(Index).Caption = detsql("酒龄")
    Lab吸烟量(Index).Caption = detsql("吸烟量")
    Lab饮酒量(Index).Caption = detsql("饮酒量")
        If Not IsNull(detsql("生活更多")) Then      '判断居住地等信息
        Lab饮食习惯(Index).Caption = detsql("生活更多")
        End If
    End If
    
    
    If Index = 2 Then   '职业健康
    Dim sql As Object
    Set sql = dafuncGetData("select * From 职业病体检_体检人员基本信息表 where 系统编号='" & mstr系统编号 & "'")
    Lab出生地(Index).Caption = sql("出生地")
    
    Lab婚否(Index).Caption = detsql("是否结婚")
    Lab子女数(Index).Caption = detsql("现有子女数目")
    Lab早产(Index).Caption = detsql("早产")
    Lab死产(Index).Caption = detsql("死产")
'    Lab流产(Index).Caption = detsql("流产")
    Lab异常胎(Index).Caption = detsql("异常胎")
    Lab吸烟程度(Index).Caption = detsql("吸烟程度")
    Lab饮酒程度(Index).Caption = detsql("饮酒程度")
    Lab烟龄(Index).Caption = detsql("烟龄")
    Lab酒龄(Index).Caption = detsql("酒龄")
    Lab吸烟量(Index).Caption = detsql("吸烟量")
    Lab饮酒量(Index).Caption = detsql("饮酒量")
    Lab戒烟时长(Index).Caption = detsql("戒烟时长")
    Lab过敏史(Index).Caption = detsql("过敏史")
    End If
    
    Lab初潮.Caption = detsql("初潮")
    Lab经期.Caption = detsql("经期")
    Lab周期.Caption = detsql("周期")
    Lab末次月经.Caption = detsql("末次月经")
    Lab停经年龄.Caption = detsql("停经年龄")
    Lab家族史.Caption = detsql("家族史")


'职业史
    SSTab1.Tab = 1
   Dim lstrWhere As String
    Dim lstrSql As String
        
        lstrSql = "select * From 职业病体检_职业史表 where 系统编号='" & mstr系统编号 & "'"
        dasubSetQueryTimeout 6000
        Set lobjRec = dafuncGetData(lstrSql)
        
        If Not lobjRec.EOF Then
            With cgrd职业史
                Set .DataSource = lobjRec
                clblInfo = .rows - 1
                .Col = 0
                .Sort = flexSortGenericDescending
                .AutoSize 0, .cols - 1, 0, 0
                .ExplorerBar = flexExSort
                .DataMode = flexDMFree
            End With
            
            If cgrd职业史.rows > 1 Then
                Dim i As Long
                Set mcolIndex = New Collection
                For i = 0 To cgrd职业史.cols - 1
                    mcolIndex.Add i, cgrd职业史.TextMatrix(0, i)
                Next
            End If

        Else
            '当职业史没有内容时，显示“无职业史”  2015-10-26
            cgrd职业史.Visible = False
            Lab职业史.Caption = "无职业史"
            Lab职业史.FontSize = 22
            cgrd职业史.rows = 1
        End If
    
    
  '既往病史
 SSTab1.Tab = 2
        lstrSql = "select * From 职业病体检_既往病史表 where 系统编号='" & mstr系统编号 & "'"
        dasubSetQueryTimeout 6000
        Set lobjRec = dafuncGetData(lstrSql)
        If Not lobjRec.EOF Then
            With cgrd病史
                Set .DataSource = lobjRec
                clblInfo = .rows - 1
                .Col = 0
                .Sort = flexSortGenericDescending
                .AutoSize 0, .cols - 1, 0, 0
                .ExplorerBar = flexExSort
                .DataMode = flexDMFree
            End With
              
            If cgrd病史.rows > 1 Then
                Set mcolIndex = New Collection
                For i = 0 To cgrd病史.cols - 1
                    mcolIndex.Add i, cgrd病史.TextMatrix(0, i)
                Next
            End If
                
        Else
         '当职业史没有内容时，显示“无既往病史”  2015-10-26
            cgrd病史.Visible = False
            Lab病史.Caption = "无既往病史"
            Lab病史.FontSize = 22
            cgrd病史.rows = 1
        End If
    
    '自觉症状
   SSTab1.Tab = 3
'        lstrSql = "select 系统编号,编号,症状,程度,出现时间 From 职业病体检_自觉症状表 where 系统编号='" & mstr系统编号 & "'"
'       lstrSql = "select * From 职业病体检_自觉症状表 where 系统编号='" & mstr系统编号 & "'"
   Dim symresql As Object
    Set symresql = dafuncGetData("select 体检表类型 From 职业病体检_体检人员基本信息表 where 系统编号='" & mstr系统编号 & "'")
    If symresql("体检表类型") = "8023部队" Or resql("体检表类型") = "涉核部队" Or resql("体检表类型") = "涉核部队YK" Then
      lstrSql = "select 系统编号,编号,症状,程度,出现时间 From 职业病体检_自觉症状表 where 系统编号='" & mstr系统编号 & "'and 出现时间!='' "
      dasubSetQueryTimeout 6000
        Set lobjRec = dafuncGetData(lstrSql)
        If Not lobjRec.EOF Then
            With cgrd症状
                Set .DataSource = lobjRec
                clblInfo = .rows - 1
                .Col = 0
                .Sort = flexSortGenericDescending
                .AutoSize 0, .cols - 1, 0, 0
                .ExplorerBar = flexExSort
                .DataMode = flexDMFree
            End With
            
            If cgrd症状.rows > 1 Then
                Set mcolIndex = New Collection
                For i = 0 To cgrd症状.cols - 1
                    mcolIndex.Add i, cgrd症状.TextMatrix(0, i)
                Next
             End If
                         
        Else
            cgrd症状.Visible = False
            Lab症状.Caption = "无自觉症状"
            Lab症状.FontSize = 22
            cgrd症状.rows = 1
        End If
    Else
    '不要体检医师  2015-10-26
    lstrSql = "select 系统编号,编号,症状,程度,出现时间 From 职业病体检_自觉症状表 where 系统编号='" & mstr系统编号 & "'"
          dasubSetQueryTimeout 6000
        Set lobjRec = dafuncGetData(lstrSql)
        If Not lobjRec.EOF Then
            With cgrd症状
                Set .DataSource = lobjRec
                clblInfo = .rows - 1
                .Col = 0
                .Sort = flexSortGenericDescending
                .AutoSize 0, .cols - 1, 0, 0
                .ExplorerBar = flexExSort
                .DataMode = flexDMFree
            End With
            
            If cgrd症状.rows > 1 Then
                Set mcolIndex = New Collection
                For i = 0 To cgrd症状.cols - 1
                    mcolIndex.Add i, cgrd症状.TextMatrix(0, i)
                Next
             End If
            Else
            cgrd症状.Visible = False
            Lab症状.Caption = "无自觉症状"
            Lab症状.FontSize = 22
            cgrd症状.rows = 1
        End If
    End If
        
        '体格一般情况
    SSTab1.Tab = 4
        Set detsql = dafuncGetData("select 体检结果 From 职业病体检_结果信息_受检者个人信息录入科 where 系统编号='" & mstr系统编号 & "'and 体检项目='13017'")
        Lab营养.Caption = detsql("体检结果")
        Set detsql = dafuncGetData("select 体检结果 From 职业病体检_结果信息_受检者个人信息录入科 where 系统编号='" & mstr系统编号 & "'and 体检项目='13018'")
        Lab身高.Caption = detsql("体检结果")
        Set detsql = dafuncGetData("select 体检结果 From 职业病体检_结果信息_受检者个人信息录入科 where 系统编号='" & mstr系统编号 & "'and 体检项目='13019'")
        Lab体重.Caption = detsql("体检结果")
        Set detsql = dafuncGetData("select 体检结果 From 职业病体检_结果信息_受检者个人信息录入科 where 系统编号='" & mstr系统编号 & "'and 体检项目='13020'")
        Lab收缩压.Caption = detsql("体检结果")
        Set detsql = dafuncGetData("select 体检结果 From 职业病体检_结果信息_受检者个人信息录入科 where 系统编号='" & mstr系统编号 & "'and 体检项目='13021'")
        Lab舒张压.Caption = detsql("体检结果")
        Set detsql = dafuncGetData("select 体检结果 From 职业病体检_结果信息_内科 where 系统编号='" & mstr系统编号 & "'and 体检项目='02002'")
        If IsNull(detsql("体检结果")) Then
        Lab心率.Caption = "未录入"
        Else
        Lab心率.Caption = detsql("体检结果")
        End If
        Set detsql = dafuncGetData("select 体检结果 From 职业病体检_结果信息_内科 where 系统编号='" & mstr系统编号 & "'and 体检项目='02019'")
        If detsql.RecordCount > 0 Then
        Label37.Visible = True
        Lab发育.Visible = True
           If detsql("体检结果") = "" Then
           Lab发育.Caption = "/"
           Else
           Lab发育.Caption = detsql("体检结果")
           End If
        End If
    '隐藏第六个选项卡  2015-10-26
        SSTab1.Tab = 5
        SSTab1.TabVisible(5) = False
        
End Sub


