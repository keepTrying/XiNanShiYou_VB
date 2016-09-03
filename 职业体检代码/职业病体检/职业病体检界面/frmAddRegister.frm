VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{2327075A-FED2-4AEF-82B8-DD0C2B1AC8E1}#2.0#0"; "dyCatchPhoto.ocx"
Object = "{F48D4DEC-1198-11D5-91BE-0050BA06B70C}#5.9#0"; "录入控件.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAddRegister 
   Caption         =   "Form1"
   ClientHeight    =   9465
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12360
   LinkTopic       =   "Form1"
   ScaleHeight     =   9465
   ScaleWidth      =   12360
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   6600
      Top             =   360
   End
   Begin VB.CheckBox cchkClear 
      Caption         =   "保存后清空"
      Height          =   345
      Left            =   8520
      TabIndex        =   82
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.CheckBox Check身份证 
      Caption         =   "刷二代身份证"
      Height          =   255
      Left            =   8520
      TabIndex        =   81
      Top             =   480
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   900
      Left            =   5640
      Top             =   360
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7335
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   12938
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "基本信息录入            "
      TabPicture(0)   =   "frmAddRegister.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "clblHistory"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "clblHintCheck"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label34"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label33"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label30"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label12"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label3"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label2(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label8"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label2(6)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label6"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label5"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label2(2)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label4"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label2(3)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cctlCatchPhoto"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "ctlInputDictGrid1"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cdtpDate"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "cdtp出生"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "cgrdHistory"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "ccmbSex"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Picture2"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Ccmb体检人类别"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "ccmb体检人类型"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "clblsysno"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "ctxt身份证号"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "ctxtAge"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "ctxtName"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "ccmbTemplate"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "ccmb体检类型"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "ccmb体检类别"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).ControlCount=   32
      TabCaption(1)   =   "附加信息录入           "
      TabPicture(1)   =   "frmAddRegister.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(2)=   "Frame3"
      Tab(1).ControlCount=   3
      Begin VB.ComboBox ccmb体检类别 
         Height          =   300
         ItemData        =   "frmAddRegister.frx":0038
         Left            =   6000
         List            =   "frmAddRegister.frx":003F
         Style           =   2  'Dropdown List
         TabIndex        =   62
         Top             =   960
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.ComboBox ccmb体检类型 
         Height          =   300
         Left            =   7080
         TabIndex        =   61
         Top             =   1320
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.ComboBox ccmbTemplate 
         Height          =   300
         ItemData        =   "frmAddRegister.frx":0051
         Left            =   360
         List            =   "frmAddRegister.frx":0053
         TabIndex        =   60
         Top             =   2040
         Width           =   3480
      End
      Begin VB.TextBox ctxtName 
         BackColor       =   &H00C0FFFF&
         Height          =   300
         Left            =   2640
         TabIndex        =   59
         Top             =   3000
         Width           =   1410
      End
      Begin VB.TextBox ctxtAge 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   1440
         TabIndex        =   58
         Top             =   3720
         Width           =   495
      End
      Begin VB.TextBox ctxt身份证号 
         Height          =   300
         Left            =   360
         TabIndex        =   57
         Top             =   3000
         Width           =   2130
      End
      Begin VB.TextBox clblsysno 
         Height          =   270
         Left            =   360
         TabIndex        =   56
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Frame Frame1 
         Caption         =   "  个人附加信息      "
         ForeColor       =   &H000080FF&
         Height          =   2295
         Left            =   -74520
         TabIndex        =   37
         Top             =   480
         Width           =   9735
         Begin VB.TextBox ctxt籍贯 
            Height          =   300
            Left            =   480
            TabIndex        =   46
            Top             =   600
            Width           =   1815
         End
         Begin VB.TextBox ctxt邮编 
            Height          =   300
            Left            =   2640
            TabIndex        =   45
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox ctxt住址 
            Height          =   300
            Left            =   4560
            TabIndex        =   44
            Top             =   600
            Width           =   4215
         End
         Begin VB.ComboBox ccmb文化程度 
            Height          =   300
            ItemData        =   "frmAddRegister.frx":0055
            Left            =   480
            List            =   "frmAddRegister.frx":0074
            TabIndex        =   43
            Text            =   "ccmb文化程度"
            Top             =   1320
            Width           =   1815
         End
         Begin VB.TextBox ctxt工龄 
            Height          =   270
            Left            =   7920
            TabIndex        =   42
            Top             =   1320
            Width           =   975
         End
         Begin VB.TextBox ctxt电话 
            Height          =   300
            Left            =   5400
            TabIndex        =   41
            Top             =   1320
            Width           =   1815
         End
         Begin VB.ComboBox Ccmb婚否 
            Height          =   300
            ItemData        =   "frmAddRegister.frx":00C6
            Left            =   2640
            List            =   "frmAddRegister.frx":00D3
            TabIndex        =   40
            Text            =   "Ccmb婚否"
            Top             =   1320
            Width           =   855
         End
         Begin VB.ComboBox ccmb民族 
            Height          =   300
            ItemData        =   "frmAddRegister.frx":00E9
            Left            =   3840
            List            =   "frmAddRegister.frx":00F0
            TabIndex        =   39
            Text            =   "ccmb民族"
            Top             =   1320
            Width           =   1215
         End
         Begin VB.TextBox ctxt出生地 
            Height          =   300
            Left            =   1200
            TabIndex        =   38
            Top             =   1750
            Width           =   4455
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "籍贯："
            Height          =   180
            Left            =   480
            TabIndex        =   55
            Top             =   360
            Width           =   540
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "邮政编号："
            Height          =   180
            Left            =   2640
            TabIndex        =   54
            Top             =   360
            Width           =   900
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "住址："
            Height          =   180
            Left            =   4560
            TabIndex        =   53
            Top             =   360
            Width           =   540
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "工龄："
            Height          =   180
            Left            =   7920
            TabIndex        =   52
            Top             =   1080
            Width           =   540
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "电话号码："
            Height          =   180
            Left            =   5400
            TabIndex        =   51
            Top             =   1080
            Width           =   900
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "文化程度："
            Height          =   180
            Left            =   480
            TabIndex        =   50
            Top             =   1080
            Width           =   900
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "婚否："
            Height          =   180
            Left            =   2640
            TabIndex        =   49
            Top             =   1080
            Width           =   540
         End
         Begin VB.Label label80 
            AutoSize        =   -1  'True
            Caption         =   "民族："
            Height          =   180
            Left            =   3840
            TabIndex        =   48
            Top             =   1080
            Width           =   540
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "出生地："
            Height          =   180
            Left            =   480
            TabIndex        =   47
            Top             =   1800
            Width           =   720
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "  危害因素信息录入    "
         ForeColor       =   &H000080FF&
         Height          =   1815
         Left            =   -74520
         TabIndex        =   22
         Top             =   2880
         Width           =   9735
         Begin VB.TextBox ctxt危害工龄 
            Height          =   270
            Left            =   4560
            TabIndex        =   29
            Top             =   1200
            Width           =   1815
         End
         Begin VB.TextBox ctxt放射剂量 
            Height          =   270
            Left            =   6720
            TabIndex        =   28
            Top             =   1200
            Width           =   1815
         End
         Begin VB.ComboBox ccmb职务 
            Height          =   300
            Left            =   2640
            TabIndex        =   27
            Top             =   1200
            Width           =   1575
         End
         Begin VB.ComboBox ccmb现工种 
            Height          =   300
            Left            =   480
            TabIndex        =   26
            Top             =   1200
            Width           =   1815
         End
         Begin VB.ComboBox ccmb照射源 
            Height          =   300
            ItemData        =   "frmAddRegister.frx":00F8
            Left            =   480
            List            =   "frmAddRegister.frx":016A
            TabIndex        =   25
            Top             =   480
            Width           =   1815
         End
         Begin VB.ComboBox ccmb职业类别 
            Height          =   300
            ItemData        =   "frmAddRegister.frx":02F4
            Left            =   2640
            List            =   "frmAddRegister.frx":02F6
            TabIndex        =   24
            Top             =   480
            Width           =   1575
         End
         Begin VB.ComboBox ccmb危害因素 
            Height          =   300
            Left            =   4560
            TabIndex        =   23
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "职业危害工龄："
            Height          =   180
            Left            =   4560
            TabIndex        =   36
            Top             =   960
            Width           =   1260
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "放射剂量："
            Height          =   180
            Left            =   6720
            TabIndex        =   35
            Top             =   960
            Width           =   900
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "职业/职称："
            Height          =   180
            Left            =   2640
            TabIndex        =   34
            Top             =   960
            Width           =   990
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "现工种："
            Height          =   180
            Left            =   480
            TabIndex        =   33
            Top             =   960
            Width           =   720
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "照射源："
            Height          =   180
            Left            =   480
            TabIndex        =   32
            Top             =   240
            Width           =   720
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "职业类别："
            Height          =   180
            Left            =   2640
            TabIndex        =   31
            Top             =   240
            Width           =   900
         End
         Begin VB.Label Label17 
            Caption         =   "危害因素："
            Height          =   255
            Left            =   4560
            TabIndex        =   30
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "  现单位信息录入   "
         ForeColor       =   &H000080FF&
         Height          =   2415
         Left            =   -74520
         TabIndex        =   7
         Top             =   4800
         Width           =   9735
         Begin VB.ComboBox ccmbUnit 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   300
            Left            =   600
            TabIndex        =   15
            Top             =   480
            Width           =   3480
         End
         Begin VB.CommandButton ccmd单位定位 
            Caption         =   "定位(&T)"
            Height          =   375
            Left            =   4080
            Style           =   1  'Graphical
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   360
            Width           =   945
         End
         Begin VB.CheckBox cchk录入单位名称 
            Caption         =   "录入单位名称"
            Height          =   255
            Left            =   6600
            TabIndex        =   13
            Top             =   240
            Value           =   1  'Checked
            Width           =   1575
         End
         Begin VB.TextBox ctxt负责人 
            Height          =   300
            Left            =   480
            TabIndex        =   12
            Top             =   1200
            Width           =   1695
         End
         Begin VB.TextBox ctxt联系电话 
            Height          =   300
            Left            =   2520
            TabIndex        =   11
            Top             =   1200
            Width           =   1575
         End
         Begin VB.TextBox ctxt单位地址 
            Height          =   300
            Left            =   480
            TabIndex        =   10
            Top             =   1920
            Width           =   5775
         End
         Begin VB.ComboBox ccmb经济性质 
            Height          =   300
            Left            =   4560
            TabIndex        =   9
            Text            =   "ccmb经济性质"
            Top             =   1200
            Width           =   1695
         End
         Begin VB.ComboBox Ccmb行业类别 
            Height          =   300
            Left            =   6720
            TabIndex        =   8
            Text            =   "Ccmb行业类别"
            Top             =   1200
            Width           =   1815
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "单位名称："
            Height          =   180
            Index           =   5
            Left            =   480
            TabIndex        =   21
            Top             =   240
            Width           =   900
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "负责人："
            Height          =   180
            Left            =   480
            TabIndex        =   20
            Top             =   960
            Width           =   720
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "联系电话："
            Height          =   180
            Left            =   2520
            TabIndex        =   19
            Top             =   960
            Width           =   900
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "单位地址："
            Height          =   180
            Left            =   480
            TabIndex        =   18
            Top             =   1680
            Width           =   900
         End
         Begin VB.Label Label28 
            Caption         =   "经济性质："
            Height          =   255
            Left            =   4560
            TabIndex        =   17
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "行业类别："
            Height          =   180
            Left            =   6720
            TabIndex        =   16
            Top             =   960
            Width           =   900
         End
      End
      Begin VB.ComboBox ccmb体检人类型 
         Height          =   300
         ItemData        =   "frmAddRegister.frx":02F8
         Left            =   2880
         List            =   "frmAddRegister.frx":030E
         TabIndex        =   5
         Top             =   1320
         Width           =   1695
      End
      Begin VB.ComboBox Ccmb体检人类别 
         Height          =   300
         ItemData        =   "frmAddRegister.frx":0352
         Left            =   4800
         List            =   "frmAddRegister.frx":0365
         TabIndex        =   4
         Top             =   1320
         Width           =   1935
      End
      Begin VB.PictureBox Picture2 
         Height          =   1935
         Left            =   4200
         ScaleHeight     =   1875
         ScaleWidth      =   1515
         TabIndex        =   3
         Top             =   2040
         Width           =   1575
      End
      Begin VB.ComboBox ccmbSex 
         BackColor       =   &H00C0FFFF&
         Height          =   300
         ItemData        =   "frmAddRegister.frx":038F
         Left            =   360
         List            =   "frmAddRegister.frx":0399
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   3720
         Width           =   855
      End
      Begin VSFlex8Ctl.VSFlexGrid cgrdHistory 
         Height          =   2535
         Left            =   360
         TabIndex        =   2
         Top             =   4560
         Visible         =   0   'False
         Width           =   5415
         _cx             =   2088772943
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
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
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
      Begin MSComCtl2.DTPicker cdtp出生 
         Height          =   300
         Left            =   2400
         TabIndex        =   6
         Top             =   3720
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         Format          =   59572224
         CurrentDate     =   40960
      End
      Begin MSComCtl2.DTPicker cdtpDate 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "yyyy-MM-dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   3
         EndProperty
         Height          =   315
         Left            =   8760
         TabIndex        =   63
         Top             =   960
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   59572224
         CurrentDate     =   36950
         MaxDate         =   73050
         MinDate         =   17899
      End
      Begin 录入控件.ctlInputDictGrid ctlInputDictGrid1 
         Height          =   2535
         Left            =   8280
         TabIndex        =   64
         Top             =   6480
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   4471
         Cols            =   10
         Count           =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin dyCatchPhoto.ctlCatchPhoto cctlCatchPhoto 
         Height          =   3570
         Left            =   6120
         TabIndex        =   84
         Top             =   2040
         Width           =   4485
         _ExtentX        =   8017
         _ExtentY        =   6297
         BackColor       =   0
         FontSize        =   9.75
         OriginalSize    =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "体检类别："
         Height          =   180
         Index           =   3
         Left            =   4800
         TabIndex        =   80
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label Label4 
         Caption         =   "体检人员类型："
         Height          =   255
         Left            =   2880
         TabIndex        =   79
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "体检表："
         Height          =   180
         Left            =   360
         TabIndex        =   78
         Top             =   1800
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "体检日期："
         Height          =   180
         Index           =   2
         Left            =   8760
         TabIndex        =   77
         Top             =   720
         Width           =   900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名："
         Height          =   180
         Left            =   2640
         TabIndex        =   76
         Top             =   2760
         Width           =   540
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别："
         Height          =   180
         Left            =   360
         TabIndex        =   75
         Top             =   3480
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄："
         Height          =   180
         Index           =   6
         Left            =   1440
         TabIndex        =   74
         Top             =   3480
         Width           =   540
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "身份证号："
         Height          =   180
         Left            =   360
         TabIndex        =   73
         Top             =   2760
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "系统编号："
         Height          =   180
         Index           =   0
         Left            =   360
         TabIndex        =   72
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "注：刷条码前请确保文本框中内容为空"
         Height          =   180
         Left            =   360
         TabIndex        =   71
         Top             =   720
         Visible         =   0   'False
         Width           =   3060
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C000&
         Caption         =   "非快速录入时黄色为必录项，快速录入时只需刷二代身份证"
         Height          =   180
         Left            =   360
         TabIndex        =   70
         Top             =   480
         Visible         =   0   'False
         Width           =   4680
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "出生日期："
         Height          =   180
         Left            =   2400
         TabIndex        =   69
         Top             =   3480
         Width           =   900
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "注：先刷条码，再刷身份证"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   7680
         TabIndex        =   68
         Top             =   6240
         Visible         =   0   'False
         Width           =   2160
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "请将二代身份证放在读卡器上！"
         Height          =   180
         Left            =   360
         TabIndex        =   67
         Top             =   2520
         Width           =   2520
      End
      Begin VB.Label clblHintCheck 
         Caption         =   "注意：校核之后只允许照相，其它内容即使修改，也不会保存。"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3480
         TabIndex        =   66
         Top             =   720
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.Label clblHistory 
         Caption         =   "双击行，导入体检基本信息和附加信息。"
         Height          =   255
         Left            =   360
         TabIndex        =   65
         Top             =   4320
         Visible         =   0   'False
         Width           =   3375
      End
   End
   Begin MSComctlLib.Toolbar ctbMain 
      Align           =   1  'Align Top
      Height          =   615
      Left            =   0
      TabIndex        =   83
      Top             =   0
      Width           =   12360
      _ExtentX        =   21802
      _ExtentY        =   1085
      ButtonWidth     =   609
      ButtonHeight    =   926
      Appearance      =   1
      ImageList       =   "cimg按钮图标"
      _Version        =   393216
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
   Begin MSComDlg.CommonDialog ccmdFile 
      Left            =   0
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmAddRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************************
'功能：职业病体检登记界面控制；
'      二代身份证数据读取
'      填充组合框选项
'作者: 刘云乐
'时间：2012-03
'**********************************************************************

Public pstr系统编号 As String
'2012-08-18 于登淼 ↓
'增加复查相关变量
Public pstr复查系统编号 As String
'2012-08-18 于登淼 ↑

'2012-06-15 于登淼 ↓
'添加公共变量，记录当前体检人员状态
Public mintState As Integer '0表示从未保存；1表示校核通过；2表示修改后已保存
'2012-06-15 于登淼 ↑

Private mobj旧体检 As Object                   '年检人员的最近次的体检。
Private mobj体检 As Object                     '新职业病对象，提供获取系统编号和试管编号，保存登记信息的方法。
Private mobj体检集 As Object                   '体检集，用来定位需要年检的体检人员信息。
Private mobj体检表模板 As Object               '体检表模板，获取所有的非复查体检表模板名称。
Private WithEvents mobjGUI As cls界面通用对象  '界面通用对象，用来初始化工具栏，控制录入板控件。
Attribute mobjGUI.VB_VarHelpID = -1
 Public pblnOk As Boolean
Public selectedDeptName As Collection
'业务设置。
Private mblnTakePhoto As Boolean               '业务设置‘是否照相’。
Private mbln快速录入 As Boolean

Private mcolTubeNo As New Collection           '当前体检表可选的试管字母。

Private mstr单位申请编号 As String             '单位定位出申请编号。
Private mblnInUse As Boolean

'新选择的体检项目、收费项目
Private mcol体检项目 As New Collection
Private mcol收费项目 As New Collection               'item:编号,key：编号。

Public pstr系统编号名称 As String

Private mobj记忆  As cls用户操作记忆
Private mstr默认年龄 As String

Private Sub Ccmb体检人类别_Click()
    Dim lobj As Object
    Dim lsql As String
    Dim a As String  '体检类型
    Dim b As String  '体检类别
    Dim c As String  '体检表
    Dim i As Integer
    a = ccmb体检人类型.Text
    b = Ccmb体检人类别.Text
    ccmbTemplate.Clear
    lsql = "select 体检表名称 from 职业病体检_体检表模板基本信息表 where 体检人员类型='" & a & "' and 体检类别='" & b & "'"
    Set lobj = dafuncGetData(lsql)
    If lobj.RecordCount > 0 Then
        lobj.MoveFirst
        For i = 0 To lobj.RecordCount - 1
            c = lobj("体检表名称")
            ccmbTemplate.AddItem c, i
        lobj.MoveNext
        Next
    End If
End Sub

Private Sub ccmb体检人类型_Click()
    Dim lobj As Object
    Dim lsql As String
    Dim a As String  '体检类型
    Dim b As String  '体检类别
    Dim c As String  '体检表
    Dim i As Integer
    a = ccmb体检人类型.Text
    b = Ccmb体检人类别.Text
    ccmbTemplate.Clear
    lsql = "select 体检表名称 from 职业病体检_体检表模板基本信息表 where 体检人员类型='" & a & "' and 体检类别='" & b & "'"
    Set lobj = dafuncGetData(lsql)
    If lobj.RecordCount > 0 Then
        lobj.MoveFirst
        For i = 0 To lobj.RecordCount - 1
            c = lobj("体检表名称")
            ccmbTemplate.AddItem c, i
        lobj.MoveNext
        Next
    End If
End Sub
Sub subInit各科体检状态(paraCol As Collection, paraSysNo As String)
    Dim i As Integer
    Dim paraDeptNo As Integer
    Dim paraState, strSQL As String
    
    
    For i = 1 To 19: paraState = paraState & "0": Next
    paraState = paraState & "1"
    
    For i = 1 To paraCol.Count
        paraDeptNo = CInt(Left(paraCol.Item(i).Item(1), 2))
        paraState = Left(paraState, paraDeptNo - 1) & "1" & Right(paraState, Len(paraState) - (paraDeptNo))
    Next
    paraState = Left(paraState, 17)
    strSQL = "update 职业病体检_体检基本信息表 set 各科体检状态='" & paraState & "' where 系统编号='" & paraSysNo & "'"
    dafuncGetData strSQL
End Sub
Private Sub save优化的体检项目(ByRef para体检项目 As Collection, ByVal para系统编号 As String)
    Dim lstrSql As String
    Dim MedicProjt As String
    Dim rs As Object
    Dim i As Integer
    Dim col体检项目 As Collection
    On Error GoTo errHandler
    
    Set rs = dafuncGetData("select 名称 from 系统管理_字典_字典内容表 where ID = (select ID from 系统管理_字典_字典表列表 where 名称='职业病体检科室字典') and 名称 like '%科'")
    
    For i = 1 To rs.RecordCount
        
        lstrSql = "delete 职业病体检_结果信息_" & rs("名称") & " where 系统编号='" & para系统编号 & "'"
        dafuncGetData lstrSql
        rs.MoveNext
    Next i
    
    Set col体检项目 = para体检项目
    
    For i = 1 To col体检项目.Count
        MedicProjt = Left(Trim(col体检项目(i)("编码")), 2)
        
        lstrSql = "select 名称 from 系统管理_字典_字典内容表 where ID = (select ID from 系统管理_字典_字典表列表 where 名称='职业病体检科室字典') and 编号= '" & MedicProjt & "'"
        Set rs = dafuncGetData(lstrSql)
        
        lstrSql = "insert into 职业病体检_结果信息_" & rs("名称") & "(系统编号,体检项目) values(" _
            & "'" & para系统编号 & "','" & col体检项目(i)("编码") & "')"
        dafuncGetData lstrSql
    Next i
    
    
    Exit Sub
errHandler:
   sfsub错误处理 "职业病对象", "frmImportExcel", "save优化的体检项目", Err.Number, Err.Description, False
End Sub

Private Sub ccmd单位定位_Click()
    FrmCompany.Command1.Visible = True
    FrmCompany.Show 1
'    ccmbUnit.Text = FrmCompany.cgrdMain.TextMatrix(FrmCompany.cgrdMain.Row, 1)
End Sub

Private Sub clblsysno_Click()
    If clblsysno.Text = "" Or (clblsysno.Text <> "" And Len(clblsysno.Text) = 15) Then
'        clblsysno.Text = mobj体检.Func分配职业病体检系统编号 & (ccmb体检人类型.ListIndex + 1)
        clblsysno.Text = mobj体检.Func分配职业病体检系统编号 & "1"    '系统编号尾号全部为"1"
    End If
End Sub
'Private Sub Form_Activate()
'    On Error Resume Next
'    If mblnTakePhoto Then
'        '重新初始化照相控件。
'        cctlCatchPhoto.funcInitVideo
'    End If
'    'ctxtName.SetFocus
'End Sub
Private Sub Form_Load()
    Set mobj体检 = CreateObject("职业病对象.clsMedicalExam")
    Set mobjGUI = New cls界面通用对象
    
    '设置工具栏上所需要的各种按钮。
    Dim lcol工具栏按钮 As New Collection           '工具栏上的按钮初始化集合。
    With lcol工具栏按钮
        .Add "清空(&Cl)110"
'        .Add "|"
'        .Add "体检项目(&T)102"
'        .Add "载入照片(&E)103"
        .Add "照相(&R)101"
        .Add "|"
        .Add "保存"
        .Add "|"
        .Add "退出"
    End With
    With mobjGUI
        Set .Form = Me
        Set .c工具栏 = ctbMain
        '调用界面通用对象提供的方法，对界面控件进行初始化。
        .subInitialize lcol工具栏按钮, ""
    End With
'    ctbMain.Buttons(3).Visible = False
'    ctbMain.Buttons(4).Visible = False
    cdtpDate.Value = Now
    subClear
    subAddList
    sub读卡器初始化
    Timer2.Enabled = True
End Sub
Private Sub subClear()
   clblsysno.Text = ""
   ccmb体检人类型.Text = ""
   Ccmb体检人类别.Text = ""
   ccmbTemplate.Text = ""
   ctxt身份证号.Text = ""
   ctxtName.Text = ""
   ccmbSex.Text = ""
   ctxtAge.Text = ""
   ctxt籍贯.Text = ""
   ctxt邮编.Text = ""
   ctxt住址.Text = ""
   ccmb文化程度.Text = ""
   Ccmb婚否.Text = ""
   ccmb民族.Text = ""
   ctxt电话.Text = ""
   ctxt工龄.Text = ""
   ctxt出生地.Text = ""
   ccmb照射源.Text = ""
   ccmb职业类别.Text = ""
   ccmb危害因素.Text = ""
   ccmb现工种.Text = ""
   ccmb职务.Text = ""
   ctxt危害工龄.Text = ""
   ctxt放射剂量.Text = ""
   ccmbUnit.Text = ""
   ctxt负责人.Text = ""
   ctxt联系电话.Text = ""
   ccmb经济性质.Text = ""
   Ccmb行业类别.Text = ""
   ctxt单位地址.Text = ""
End Sub
'录完身份证号后，获取年龄，性别，出生日期
Private Sub ctxt身份证号_lostfocus()
    Dim ldatBirth As String
    Dim lstrSex As String
    On Error GoTo errHandler
    If Trim(ctxt身份证号.Text) <> "" Then
            '正确时从身份证号中获取出生日期。
            sub根据公民身份号码获取生日和性别 ctxt身份证号.Text, ldatBirth, lstrSex
            If Not IsDate(ldatBirth) Then
                MsgBox ("身份证号不合法！")
                Exit Sub
            End If
            
            '查找是否需要录入出生日期，需要时自动根据身份证号填写出生日期
            On Error Resume Next
            If IsDate(ldatBirth) Then
                cdtp出生.Value = ldatBirth
'                ctxtAge.Text = DateDiff("yyyy", ldatBirth, Date)
                ctxtAge.Text = Year(Date) - Year(ldatBirth)
            End If
            ccmbSex.Text = lstrSex
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "职业病界面", "frmAddregister", "Sub ctxt身份证号_lostfocus", Err.Number, Err.Description, True
End Sub
Private Sub subAddList()
    ccmb体检人类型.ListIndex = 1
    Ccmb体检人类别.ListIndex = 1
    
End Sub

'二代身份证读卡器，初始化，PC与终端的连接
Private Sub sub读卡器初始化()
    'CVR_InitComm
    On Error GoTo errHandler
   Dim n, ret, nLen
    Comm = False
    
    For n = 1001 To 1016 Step 1     '依次检查USB端口1001-1016
 
      If (InitComm(n)) Then
            Comm = True
       
            'StateLabel.Caption = "成功打开端口！"
            'ret = MsgBox("成功打开端口！请将卡置于阅读器上。", vbOKOnly + vbInformation, "提示")
        
            Exit For
                    
        End If
       
    Next n
    If (Comm = False) Then
     For n = 1 To 4 Step 1     '依次检查串口1-16
    
        If (InitComm(n)) Then
            Comm = True
       
            'StateLabel.Caption = "成功打开端口！"
            'ret = MsgBox("成功打开端口！请将卡置于阅读器上。", vbOKOnly + vbInformation, "提示")
    
           Exit For
                    
        End If
       
       Next n
    End If
    
   
  
    If (Comm = False) Then
    
            ret = MsgBox("打开端口不成功！请检查设备连接。", vbOKOnly + vbCritical, "错误")
            
            Exit Sub
    
    End If
    Exit Sub
errHandler:
'    sfsub错误处理 "职业病界面部件", "frmregister", "func读卡器初始化", Err.Number, Err.Description, True
End Sub
Private Sub Timer2_Timer()
    On Error Resume Next
    Dim n, ret, nLen
    Dim iname As String * 31
    Dim isex  As String * 3
    Dim folk As String * 10
    Dim code As String * 19
    Dim addr As String * 71
    Dim birthday As String * 9
    Dim startdate As String * 9
    Dim enddate As String * 9
    Dim agency As String * 31
    Dim Msg As String * 300
    Dim Msg1 As String * 256
    Dim IINSNDN As String * 64
    Dim SAMID As String * 36
    Dim LenT As Integer
    ChDir (App.Path)                '改变当前默认路径为应用程序所在路径
    ret = Authenticate()
    If (ret) Then
       ret = ReadBaseInfos(iname, isex, folk, birthday, code, addr, agency, startdate, enddate)
       ctxtName.Text = Trim(Split(iname, "")(0))
        Timer2.Enabled = False

        '当检测到有身份证并读取数据成功后，关闭timer2
        ctxt身份证号.Enabled = True
        ctxt身份证号.SetFocus      '先setfocus后lostfocus目的是，读取完身份证号后调用  ctxt身份证号.lostfocus 事件函数，以计算出性别，年龄，出生年月
        'Call sub获取证号
        ctxt身份证号.Text = Trim(code)
'        ctxt身份证号_KeyDown 13, 1
        ctxt身份证号.Enabled = False
        ctxtName.Enabled = True
        ctxtName.SetFocus
        ctxtName.Text = Trim(Split(iname, "")(0))
        ctxtName.Enabled = False
        ctxt住址.Text = Trim(addr)
        ccmb民族.Text = Trim(folk)
        Picture2.Picture = LoadPicture(App.Path & "\photo.bmp")
    End If
    Exit Sub
errHandler:
'    sfsub错误处理 "职业病界面部件", "frmregister", "timer2_timer", Err.Number, Err.Description, True
End Sub
Private Sub sub保存()
    Dim lsql As String
    Dim SNO As String
    Dim name, sex, age As String
    Dim a, b, c, D, E, F As String
    Dim G, H, i, j, K, L, M, n As String
    Dim o, P, Q, R, S, T, U, V, W, X, Y, Z As String
            U = ccmb体检人类型.Text
            V = Ccmb体检人类别.Text
            W = cdtp出生.Value
            X = Left(ccmb文化程度.Text, 2)
            Y = Trim(Split(ccmbTemplate.Text, "-")(2))
            age = Trim(ctxtAge.Text)
            SNO = Trim(clblsysno.Text)
            name = ctxtName.Text
           sex = ccmbSex.Text
            a = ccmbUnit.Text
            b = ccmb危害因素.Text
            c = Right(ccmb照射源.Text, 2)
            D = ccmb职业类别.Text
            E = ccmb现工种.Text
            F = ccmb职务.Text
            G = ctxt危害工龄.Text
            H = Trim(ctxt放射剂量.Text)
            i = Trim(ctxt籍贯.Text)
            j = Trim(ctxt邮编.Text)
            K = Trim(ctxt住址.Text)
            L = Ccmb婚否.Text
            M = Trim(ctxt电话.Text)
            n = Trim(ctxt工龄.Text)
            o = Trim(ctxt出生地.Text)
            P = ctxt负责人.Text
            Q = ctxt联系电话.Text
            R = ccmb经济性质.Text
            S = Ccmb行业类别.Text
            T = ctxt单位地址.Text
    lsql = "update 职业病体检_体检人员基本信息表 set 姓名='" & name & "',性别='" & sex & "',年龄='" & age & "',出生日期='" & W & "',出生地='" & o & "',危害因素='" & Y & "',职业分类='" & D & "',照射源='" & c & "',现工种='" & E & "',职务或职称='" & F & "',放射剂量='" & H & "',工龄='" & n & "',职业危害工龄='" & G & "',电话号码='" & M & "',住址='" & K & "',邮编='" & j & "',文化程度='" & X & "',籍贯='" & i & "',民族='" & ccmb民族.Text & "',婚否='" & L & "',体检表类型='" & U & "',体检表类别='" & V & "'where 系统编号='" & SNO & "'"
    dafuncGetData (lsql)
End Sub

Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    Dim i As Integer
    Dim lstr流水号 As String
    Dim lstr系统编号 As String
    Dim lcol原体检项目 As Collection
    Dim lobjrec类型 As Object
    Dim lobj体检表编号 As Object
    Dim lobjRec As Object
    Dim lstrError As String
    
    '2012-06-13 于登淼 ↓
    '存储身份证照片变量和系统编号退回变量，身份证相关信息
    Dim lobjRec身份证照片 As Object
    Dim lobjRec系统编号退回 As Object
    Dim paraSysNo As String
    Dim lstrSex As String
    Dim lstrBirth As String
    Dim lstrSysNo As String

    Select Case Operate
    Case "清空"
        subClear
    
    Case "保存"
        '新增，保存之前先将基本信息表添加完成
        Dim lobj1 As Object
        Dim S As String
        '单位必填
        If ccmbUnit.Text = "" Then
            MsgBox "没有录入单位信息"
            Exit Sub
        Else
            Dim a As String
            Dim OB As Object
            Set OB = dafuncGetData("select * from  单位档案_单位基本信息表 where 单位名称='" & Trim(ccmbUnit.Text) & "'")
            If OB.RecordCount > 0 Then
                a = OB("申请编号")
            ElseIf OB.RecordCount = 0 Then
                MsgBox "您没有录入单位信息，请按定位，先录入单位信息"
                Exit Sub
            Else
                MsgBox "您选择的单位重复录入,请删除重复单位后再选择"
                Exit Sub
            End If
        End If
        '添加体检人员基本信息表
        S = "select * from 职业病体检_体检人员基本信息表 where 系统编号='" & clblsysno.Text & "'"
        Set lobj1 = dafuncGetData(S)
        If lobj1.RecordCount = 0 Then
            dafuncGetData ("insert into 职业病体检_体检人员基本信息表 values('" & clblsysno.Text & "','','','','','','','" & a & "','" & Trim(ccmbUnit.Text) & "','" & Now & "','','','','','','','','','','','','','','','',null,null,null,null,null,'','')")
            dafuncGetData ("update  职业病体检_体检人员基本信息表 set 公民身份号码='" & ctxt身份证号.Text & "',建档日期='" & Now & "'where 系统编号='" & clblsysno.Text & "'")
        End If
        '添加体检基本信息表
        S = "select * from 职业病体检_体检基本信息表 where 系统编号='" & clblsysno.Text & "'"
        Set lobj1 = dafuncGetData(S)
        If lobj1.RecordCount = 0 Then
            dafuncGetData ("insert into 职业病体检_体检基本信息表 values('" & clblsysno.Text & "','','" & ccmbTemplate.Text & "','" & ccmb体检人类型.Text & "','" & Ccmb体检人类别.Text & "','" & Now & "',null,null,null,null,null,'0','',null,null,'',null,null,null)")
        End If
        If Trim(ctxtName.Text) = "" Then
            MsgBox "姓名不能为空！", vbExclamation, "系统提示"
            Exit Sub
        End If
        If Trim(ccmbSex.Text) = "" Then
            MsgBox "性别不能为空！", vbExclamation, "系统提示"
            Exit Sub
        End If
        If Trim(ctxtAge.Text) = "" Then
            MsgBox "年龄不能为空！", vbExclamation, "系统提示"
            Exit Sub
        End If
        sub保存
            Set mcol体检项目 = New Collection
            mobj体检.体检表.体检表名 = ccmbTemplate.Text
        
        '如果系统编号长度小于15，初步判定是操作失误
        If Len(clblsysno.Text) < 15 Then
            MousePointer = 0
            MsgBox "系统编号错误，请检查！", vbInformation, "系统提示"
            Exit Sub
        End If
        If Len(ctxt身份证号.Text) = 0 Then
            MsgBox ("未填入身份证号，不允许保存当前内容！")
            Exit Sub
        End If
        
        If dafuncGetData("select 体检日期 from 职业病体检_体检基本数据库 where 公民身份号码 = '" & Trim(ctxt身份证号.Text) & "' and convert(varchar(10),体检日期,111) = convert(varchar(10),getdate(),111)  and 体检状态 <> 0 ").RecordCount > 0 Then
            MsgBox "同一个身份证号当天不能多次录入！！！"
            subClear
             Exit Sub
        End If
        '保存身份证照片
        If mblnTakePhoto Then
            Dim lobjPhoto As Object
            Set lobjPhoto = cctlCatchPhoto.Photo
'        ElseIf Not Picture1.Picture Is Nothing Then
'            Set lobjPhoto = Picture1.Picture
            pmsub保存图片 lobjPhoto, Trim(clblsysno.Text), "职业病体检"
         Else
            Set lobjRec身份证照片 = CreateObject("职业病对象.clsPersonExamed")
            lobjRec身份证照片.func保存身份证照片 Picture2.Image, clblsysno.Text, "职业病体检"
            lobjRec身份证照片.func保存身份证照片 Picture2.Image, clblsysno.Text & "IDcard", "职业病体检"
            Set lobjRec身份证照片 = Nothing
        End If
        
        Set lobj体检表编号 = CreateObject("职业病对象.clsmedicalexamsheet")
        lobj体检表编号.体检表编号 = ccmbTemplate.Text
        With mobj体检
            .系统编号 = Trim(clblsysno.Text)
            
            If .体检表.体检表名 <> ccmbTemplate.Text Then
                .体检表.体检表名 = ccmbTemplate.Text
            End If

              '添加性别和年龄 2015-12-25 by 牟俊
'            mobj体检.体检人员.性别 = Trim(ccmbSex.Text)
            mobj体检.体检人员.年龄 = Trim(ctxtAge.Text)

            .体检人员.系统编号 = Trim(clblsysno.Text)
            .体检人员.姓名 = ctxtName.Text
            .体检人员.性别 = ccmbSex.Text
            .体检人员.单位名称 = ccmbUnit.Text
            .体检人员.危害因素 = ccmb危害因素.Text
            .体检人员.照射源 = Right(ccmb照射源.Text, 2)
            .体检人员.职业分类 = ccmb职业类别.Text
            .体检人员.现工种 = ccmb现工种.Text
            .体检人员.职务或职称 = ccmb职务.Text
            .体检人员.职业危害工龄 = ctxt危害工龄.Text
            .体检人员.放射剂量 = Trim(ctxt放射剂量.Text)
            .体检人员.籍贯 = Trim(ctxt籍贯.Text)
            .体检人员.邮编 = Trim(ctxt邮编.Text)
            .体检人员.住址 = Trim(ctxt住址.Text)
            .体检人员.婚否 = Ccmb婚否.Text
            .体检人员.电话号码 = Trim(ctxt电话.Text)
            .体检人员.工龄 = Trim(ctxt工龄.Text)
            .体检人员.出生地 = Trim(ctxt出生地.Text)
            .体检人员.负责人 = ctxt负责人.Text
            .体检人员.联系电话 = ctxt联系电话.Text
            .体检人员.经济性质 = ccmb经济性质.Text
            .体检人员.行业类别 = Ccmb行业类别.Text
            .体检人员.单位地址 = ctxt单位地址.Text

            If Not Picture2.Picture Is Nothing Then
                .体检人员.像片 = Picture2.Picture
        
            End If
            If Val(ctxtAge.Text) > 0 Then
                .体检人员.出生日期 = DateAdd("yyyy", -Val(ctxtAge.Text), Date)
            Else
                '如果输入字符，则记忆该年龄。
                mobj记忆.sub覆盖记忆值 "体检年龄", ctxtAge.Text
                mstr默认年龄 = ctxtAge.Text
            End If
            .体检人员.年龄 = ctxtAge.Text
            
            On Error Resume Next
            .体检人员.公民身份号码 = ctxt身份证号.Text

            .体检人员.文化程度 = Left(ccmb文化程度.Text, 2)
            
            .体检人员.民族 = ccmb民族.Text
'            If ccmbUnit.Text = "" Then
'                .体检人员.单位申请编号 = ""
'            Else
'                If .体检人员.单位申请编号 <> mstr单位申请编号 Then
'                    .体检人员.单位申请编号 = mstr单位申请编号
'                End If
'            End If

            .体检日期 = cdtpDate.Value
            .体检人类型 = ccmb体检人类型.Text
            .体检人类别 = Ccmb体检人类别.Text
            Set lobj体检表编号 = CreateObject("职业病对象.clsmedicalexamsheet")
            lobj体检表编号.体检表编号 = Trim(ccmbTemplate.Text)
            If mcol体检项目.Count = 0 Then
                Set mcol体检项目 = lobj体检表编号.体检表项目集("")
            End If
            Set .col体检项目 = mcol体检项目
            save优化的体检项目 mcol体检项目, Trim(clblsysno.Text)
            subInit各科体检状态 mcol体检项目, Trim(clblsysno.Text)
        End With
        
        Set mcol体检项目 = New Collection
       
        mobj体检.体检表.体检表名 = ccmbTemplate.Text
        Cancel = True
        frmProcess.proPercent.Value = 8
        cgrdHistory.rows = 1
        cgrdHistory.Visible = False
        clblHistory.Visible = False
        
        '添加打印标签功能
        Dim strsql1 As String
        strsql1 = "select distinct left(体检项目,2) as 项目  from 职业病体检_体检表模板体检项目表 where 体检表名称='" & ccmbTemplate.Text & "'"
        Dim objds1 As Object
        Set objds1 = dafuncGetData(strsql1)
        Dim lobjFile As Object
        Set lobjFile = CreateObject("职业病文书.cls文书")
        Dim csysno As Collection
        Set csysno = New Collection
        
        csysno.Add (mobj体检.系统编号)
        
        lobjFile.func打印体检清单 csysno ''暂停打印标签功能 2015-9-1 by lanchao
        '  Set lobjFile = Nothing
        '  '更改当前体检状态。打印清单之后，就进入体检状态。
        ''  pobj业务对象.func写入单人当前体检状态 mobj体检.系统编号, 2
        Dim c As Integer
           c = objds1.RecordCount
        objds1.MoveFirst
        For i = 0 To c - 1
        If objds1("项目") = "01" Then  '01 五官科
'           lobjFile.func打印试管标签 mobj体检.系统编号, mobj体检.体检人员.姓名, "免疫.血清"
            lobjFile.func打印试管标签 mobj体检.系统编号, mobj体检.体检人员.姓名, mobj体检.体检人员.性别, mobj体检.体检人员.年龄, "五官科"
        End If
        If objds1("项目") = "02" Then  '02 内科
'           lobjFile.func打印试管标签 mobj体检.系统编号, mobj体检.体检人员.姓名, "免疫.血清"
            lobjFile.func打印试管标签 mobj体检.系统编号, mobj体检.体检人员.姓名, mobj体检.体检人员.性别, mobj体检.体检人员.年龄, "内科"
        End If
        If objds1("项目") = "03" Then  '03 外科
'           lobjFile.func打印试管标签 mobj体检.系统编号, mobj体检.体检人员.姓名, "免疫.血清"
            lobjFile.func打印试管标签 mobj体检.系统编号, mobj体检.体检人员.姓名, mobj体检.体检人员.性别, mobj体检.体检人员.年龄, "外科"
        End If
        If objds1("项目") = "08" Then  '08 电测听
'           lobjFile.func打印试管标签 mobj体检.系统编号, mobj体检.体检人员.姓名, "免疫.血清"
            lobjFile.func打印试管标签 mobj体检.系统编号, mobj体检.体检人员.姓名, mobj体检.体检人员.性别, mobj体检.体检人员.年龄, "电测听"
        End If
        If objds1("项目") = "09" Then  '09 X光
'           lobjFile.func打印试管标签 mobj体检.系统编号, mobj体检.体检人员.姓名, "免疫.血清"
            lobjFile.func打印试管标签 mobj体检.系统编号, mobj体检.体检人员.姓名, mobj体检.体检人员.性别, mobj体检.体检人员.年龄, "X射线"
        End If
        If objds1("项目") = "10" Then  '10 心电
'           lobjFile.func打印试管标签 mobj体检.系统编号, mobj体检.体检人员.姓名, "免疫.血清"
            lobjFile.func打印试管标签 mobj体检.系统编号, mobj体检.体检人员.姓名, mobj体检.体检人员.性别, mobj体检.体检人员.年龄, "心电"
        End If
        If objds1("项目") = "11" Then  '11 B超
'           lobjFile.func打印试管标签 mobj体检.系统编号, mobj体检.体检人员.姓名, "免疫.血清"
            lobjFile.func打印试管标签 mobj体检.系统编号, mobj体检.体检人员.姓名, mobj体检.体检人员.性别, mobj体检.体检人员.年龄, "B超影像科"
        End If
        If objds1("项目") = "12" Then  '12 肺功能
'           lobjFile.func打印试管标签 mobj体检.系统编号, mobj体检.体检人员.姓名, "免疫.血清"
            lobjFile.func打印试管标签 mobj体检.系统编号, mobj体检.体检人员.姓名, mobj体检.体检人员.性别, mobj体检.体检人员.年龄, "肺功能"
        End If
        If objds1("项目") = "05" Then  '05 免疫血清
'           lobjFile.func打印试管标签 mobj体检.系统编号, mobj体检.体检人员.姓名, "免疫.血清"
            lobjFile.func打印试管标签 mobj体检.系统编号, mobj体检.体检人员.姓名, mobj体检.体检人员.性别, mobj体检.体检人员.年龄, "免疫.血清"
        End If
        If objds1("项目") = "06" Then  '06 尿常规
'           lobjFile.func打印试管标签 mobj体检.系统编号, mobj体检.体检人员.姓名, "尿常规"
            lobjFile.func打印试管标签 mobj体检.系统编号, mobj体检.体检人员.姓名, mobj体检.体检人员.性别, mobj体检.体检人员.年龄, "尿常规"
        End If
        If objds1("项目") = "07" Then  '07 染色体
'           lobjFile.func打印试管标签 mobj体检.系统编号, mobj体检.体检人员.姓名, "染色体"
            lobjFile.func打印试管标签 mobj体检.系统编号, mobj体检.体检人员.姓名, mobj体检.体检人员.性别, mobj体检.体检人员.年龄, "染色体"
        End If
        If objds1("项目") = "04" Then  '04 血常规
'           lobjFile.func打印试管标签 mobj体检.系统编号, mobj体检.体检人员.姓名, "血常规.静脉血"
            lobjFile.func打印试管标签 mobj体检.系统编号, mobj体检.体检人员.姓名, mobj体检.体检人员.性别, mobj体检.体检人员.年龄, "血常规.静脉血"
        End If
        If objds1("项目") = "17" Then  '17 生化
            lobjFile.func打印试管标签 mobj体检.系统编号, mobj体检.体检人员.姓名, mobj体检.体检人员.性别, mobj体检.体检人员.年龄, "生化"
        End If
'        If objds1("项目") = "17" Then  '17 生化
'            '对生化不是直接打印生化，而是打印常规项目  2015-12-10 by 牟俊   ↓
'            Dim lobject As Object
''            Set lobject = dafuncGetData("select distinct right(体检项目,2) as 项目  from 职业病体检_体检表模板体检项目表 where 体检表名称='" & ccmbTemplate.Text & "'and 体检项目 like '1702%'and 体检项目<>'17020'")
'            Set lobject = dafuncGetData("select distinct right(体检项目,2) as 项目编号,名称 as 项目  from 职业病体检_体检表模板体检项目表 a,职业病体检_体检项目设置表 b where a.体检表名称='" & ccmbTemplate.Text & "'and a.体检项目 like '1702%'and a.体检项目=b.编码 and a.体检项目<>'17020' and b.属性='常规'")
'            If lobject.RecordCount > 0 Then
'                Dim xiangmu As String
'                Dim zhongjian As String
'                Dim X As Integer
'                lobject.MoveFirst
'                For X = 0 To lobject.RecordCount - 1
'                zhongjian = zhongjian + "," + lobject("项目")
'                lobject.MoveNext
'                Next X
'                xiangmu = Right(zhongjian, Len(zhongjian) - 1)
'                 lobjFile.func打印试管标签 mobj体检.系统编号, mobj体检.体检人员.姓名, mobj体检.体检人员.性别, mobj体检.体检人员.年龄, xiangmu
'            End If
'            '2015-12-10 by  牟俊  ↑
'        End If
        objds1.MoveNext
        Next i
        Unload frmProcess
        MousePointer = 0
        Dim strSQL As String
        strSQL = "update 职业病体检_体检基本信息表 set 体检状态= '2'  where 系统编号='" & mobj体检.系统编号 & "'"
        dafuncGetData (strSQL)

        '2015-3-15停止读卡 刘伟
        Dim ret
        ret = CloseComm()
        Unload Me
        
    Case "照相"
        mblnTakePhoto = True
        If mblnTakePhoto Then
        '重新初始化照相控件。
            cctlCatchPhoto.funcInitVideo
        End If

    Case "退出"
        sub返回编号
        Unload Me
    End Select
End Sub

Private Sub sub返回编号()
        If clblsysno.Text <> "" And Len(clblsysno.Text) = 15 Then
'            mobj体检.Func退回职业病体检系统编号 Trim(clblsysno.Text)
            Dim lobjRec1 As Object
            Dim lobjRec2 As Object
            dasubBeginTran
                
            '首先判断该系统编号的记录是否已存在。
            dasubSetQueryTimeout 6000
            Set lobjRec1 = dafuncGetData("select 系统编号 from 职业病体检_体检基本信息表 where 系统编号='" & clblsysno.Text & "'")
            If lobjRec1.RecordCount = 0 Then
                '该系统编号没有保存，可以退回。
                dasubSetQueryTimeout 6000
                Dim max As Integer
                Dim XXX As String
                Set lobjRec2 = dafuncGetData("select 设置值 from 职业病体检_业务设置信息表 where 设置项目 = '编号最大流水号' and datepart(yyyy,说明) = datepart(yyyy,getdate())")
                max = Val(lobjRec2("设置值"))
                max = max - 1
                XXX = Str(max)
                dafuncGetData ("update 职业病体检_业务设置信息表 set 设置值='" & XXX & "' where 设置项目 = '编号最大流水号' and datepart(yyyy,说明) = datepart(yyyy,getdate())")
            End If
            dasubCommitTran
        End If
End Sub
