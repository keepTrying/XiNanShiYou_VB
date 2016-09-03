VERSION 5.00
Object = "{8099FCC2-0A81-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6.ocx"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmInputTestResult 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "体检结果录入"
   ClientHeight    =   7980
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   11070
   ClipControls    =   0   'False
   Icon            =   "FrmInputTestResult.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   11070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame3 
      Height          =   580
      Left            =   5640
      TabIndex        =   47
      Top             =   600
      Width           =   4815
      Begin VB.TextBox ctxtDoctor 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   375
         Left            =   720
         TabIndex        =   48
         Top             =   120
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker cdtpInputDate 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "yyyy-MM-dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   3
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   49
         Top             =   120
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
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
         Format          =   25165824
         CurrentDate     =   36951
         MaxDate         =   73050
         MinDate         =   17899
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "日期："
         Height          =   180
         Index           =   2
         Left            =   2280
         TabIndex        =   51
         Top             =   240
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "医师："
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   50
         Top             =   240
         Width           =   540
      End
   End
   Begin TabDlg.SSTab ctabPerson 
      Height          =   6975
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   12303
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "   单个处理(&F7)   "
      TabPicture(0)   =   "FrmInputTestResult.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1(3)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(4)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "clblInfo(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(5)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "clblInfo(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(6)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "clblInfo(2)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(7)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "clblInfo(3)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "clblInfo(4)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label1(8)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cgrdSingleList"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "ctxtSingleNo"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cpicPhoto(0)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cdtpSingleQuery"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "ccmdSingleQuery"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Frame2"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "coptType(0)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "coptType(1)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "coptType(2)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "coptType(3)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "cchkUnEnd(0)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "cchkEnd(0)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).ControlCount=   23
      TabCaption(1)   =   "   批量处理(&F8)   "
      TabPicture(1)   =   "FrmInputTestResult.frx":045E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cdtpQueryDate"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cgrdPerson"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cpicPhoto(1)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "ccmdAdd"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "ccmdRemove"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "ccmdClear"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "ccmbSheet"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "ctxt单位名称"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "cchkUnEnd(1)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "cchkEnd(1)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "ctxtBatchNo"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "coptBatchType(0)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "coptBatchType(1)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "coptBatchType(2)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).ControlCount=   15
      Begin VB.OptionButton coptBatchType 
         Caption         =   "系统编号"
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   45
         Top             =   1800
         Width           =   1095
      End
      Begin VB.OptionButton coptBatchType 
         Caption         =   "单位名称"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   44
         Top             =   1320
         Width           =   1095
      End
      Begin VB.OptionButton coptBatchType 
         Caption         =   "体检日期"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   43
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox ctxtBatchNo 
         Height          =   375
         Left            =   1320
         TabIndex        =   42
         Top             =   1740
         Width           =   2415
      End
      Begin VB.CheckBox cchkEnd 
         BackColor       =   &H00F0F3E9&
         Caption         =   "已填结果"
         Height          =   255
         Index           =   1
         Left            =   4320
         TabIndex        =   39
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CheckBox cchkUnEnd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "未填结果"
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   38
         Top             =   2280
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox cchkEnd 
         BackColor       =   &H00F0F3E9&
         Caption         =   "已填结果"
         Height          =   255
         Index           =   0
         Left            =   -72000
         TabIndex        =   37
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CheckBox cchkUnEnd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "未填结果"
         Height          =   255
         Index           =   0
         Left            =   -73080
         TabIndex        =   36
         Top             =   2760
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.OptionButton coptType 
         Caption         =   "姓名"
         Height          =   255
         Index           =   3
         Left            =   -71040
         TabIndex        =   35
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox ctxt单位名称 
         Height          =   375
         Left            =   1320
         TabIndex        =   34
         Top             =   1320
         Width           =   2415
      End
      Begin VB.OptionButton coptType 
         Caption         =   "体检单号"
         Height          =   255
         Index           =   2
         Left            =   -72240
         TabIndex        =   33
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton coptType 
         Caption         =   "试管编号"
         Height          =   255
         Index           =   1
         Left            =   -73560
         TabIndex        =   32
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton coptType 
         Caption         =   "系统编号"
         Height          =   255
         Index           =   0
         Left            =   -74880
         TabIndex        =   31
         Top             =   480
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.Frame Frame2 
         Height          =   135
         Left            =   -74880
         TabIndex        =   30
         Top             =   2460
         Width           =   5220
      End
      Begin VB.CommandButton ccmdSingleQuery 
         Caption         =   "查询(&Q)"
         Height          =   375
         Left            =   -70680
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   2640
         Width           =   1065
      End
      Begin MSComCtl2.DTPicker cdtpSingleQuery 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "yyyy-MM-dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   3
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   28
         Top             =   2640
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
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
         Format          =   25165824
         CurrentDate     =   36957
         MaxDate         =   73050
         MinDate         =   17899
      End
      Begin VB.ComboBox ccmbSheet 
         Height          =   300
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   400
         Width           =   2535
      End
      Begin VB.CommandButton ccmdClear 
         Caption         =   "清空(&C)"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2160
         Width           =   975
      End
      Begin VB.CommandButton ccmdRemove 
         Caption         =   "去掉(&R)"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1140
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2160
         Width           =   975
      End
      Begin VB.CommandButton ccmdAdd 
         Caption         =   "添加(&A)"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2160
         Width           =   975
      End
      Begin VB.PictureBox cpicPhoto 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1650
         Index           =   0
         Left            =   -71160
         ScaleHeight     =   1620
         ScaleWidth      =   1305
         TabIndex        =   13
         Top             =   780
         Width           =   1335
      End
      Begin VB.TextBox ctxtSingleNo 
         Height          =   375
         Left            =   -73920
         MaxLength       =   20
         TabIndex        =   0
         Top             =   840
         Width           =   2655
      End
      Begin VB.PictureBox cpicPhoto 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1815
         Index           =   1
         Left            =   3840
         ScaleHeight     =   1785
         ScaleWidth      =   1425
         TabIndex        =   11
         Top             =   360
         Width           =   1455
      End
      Begin VSFlex6Ctl.vsFlexGrid cgrdPerson 
         Height          =   4215
         Left            =   120
         TabIndex        =   10
         Top             =   2640
         Width           =   5295
         _cx             =   25371772
         _cy             =   25369867
         _ConvInfo       =   1
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   16777215
         BackColorAlternate=   16777215
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
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
         FormatString    =   "^系统编号         |试管编号|^姓名    |^性别|^单位名称               |^年龄|^体检单号"
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   1
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0   'False
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
      End
      Begin MSComCtl2.DTPicker cdtpQueryDate 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "yyyy-MM-dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   3
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         Top             =   840
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
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
         Format          =   25165824
         CurrentDate     =   36951
         MaxDate         =   73050
         MinDate         =   17899
      End
      Begin VSFlex6Ctl.vsFlexGrid cgrdSingleList 
         Height          =   3765
         Left            =   -74940
         TabIndex        =   41
         Top             =   3120
         Width           =   5415
         _cx             =   25371983
         _cy             =   25369073
         _ConvInfo       =   1
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   16777215
         BackColorAlternate=   15791081
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
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
         FormatString    =   "^系统编号         |试管编号|^姓名    |^性别|^单位名称               |^年龄|^体检单号"
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   1
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0   'False
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "选择体检表"
         Height          =   180
         Left            =   120
         TabIndex        =   25
         Top             =   480
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "试管编号："
         Height          =   180
         Index           =   8
         Left            =   -74880
         TabIndex        =   24
         Top             =   1440
         Width           =   900
      End
      Begin VB.Label clblInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   4
         Left            =   -73920
         TabIndex        =   23
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label clblInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   3
         Left            =   -73920
         TabIndex        =   21
         Top             =   2040
         Width           =   2655
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单位名称："
         Height          =   180
         Index           =   7
         Left            =   -74880
         TabIndex        =   20
         Top             =   2160
         Width           =   900
      End
      Begin VB.Label clblInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   2
         Left            =   -72240
         TabIndex        =   19
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄："
         Height          =   180
         Index           =   6
         Left            =   -72840
         TabIndex        =   18
         Top             =   1800
         Width           =   540
      End
      Begin VB.Label clblInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   1
         Left            =   -73920
         TabIndex        =   17
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别："
         Height          =   180
         Index           =   5
         Left            =   -74880
         TabIndex        =   16
         Top             =   1800
         Width           =   540
      End
      Begin VB.Label clblInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   0
         Left            =   -72240
         TabIndex        =   15
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名："
         Height          =   180
         Index           =   4
         Left            =   -72840
         TabIndex        =   14
         Top             =   1440
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "系统编号："
         Height          =   180
         Index           =   3
         Left            =   -74880
         TabIndex        =   12
         Top             =   960
         Width           =   900
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   5640
      TabIndex        =   9
      Top             =   1200
      Width           =   5335
      Begin VSFlex6DAOCtl.vsFlexGrid cgrdInput 
         Height          =   5895
         Left            =   120
         TabIndex        =   40
         Top             =   480
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   10398
         _ConvInfo       =   1
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
         BackColorBkg    =   16777215
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483640
         GridColorFixed  =   -2147483640
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   3
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
         Editable        =   0   'False
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "红色表示不合格,可以双击名称修改"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   1560
         TabIndex        =   27
         Top             =   240
         Width           =   2790
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "体检项目结果："
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   1260
      End
   End
   Begin MSComctlLib.StatusBar cstbMain 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   7605
      Width           =   11070
      _ExtentX        =   19526
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   19473
         EndProperty
      EndProperty
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
   Begin MSComctlLib.Toolbar ctbMain 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   11070
      _ExtentX        =   19526
      _ExtentY        =   1005
      ButtonWidth     =   820
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      ImageList       =   "cimg按钮图标"
      _Version        =   393216
      Begin VB.CheckBox cchkGrid 
         Caption         =   "网格录入"
         Height          =   255
         Left            =   8520
         TabIndex        =   46
         Top             =   200
         Width           =   1455
      End
      Begin VB.CheckBox cchk刷条码 
         Caption         =   "刷条码"
         Height          =   375
         Left            =   7320
         TabIndex        =   6
         Top             =   120
         Width           =   1095
      End
   End
   Begin MSComctlLib.ImageList cimg按钮图标 
      Left            =   5280
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmInputTestResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'作者：杨春

Private WithEvents mobj界面通用对象 As cls界面通用对象    '提供工具栏初始化、热键处理。
Attribute mobj界面通用对象.VB_VarHelpID = -1

Private mobj体检医师  As Object   'clsMedicalExamer    获取当前体检医师可以作的指定属性（常规/化验）的体检项目

Private mstr体检表名称 As String  '批量处理时，当前一批体检记录所使用的体检表模板名称。
Private mstr操作名称  As String   '对应属性"操作名称"。
Private mstr体检项目属性 As String
Private mstr系统编号固定部分 As String

Private mblnSys As Boolean

Private mblnInUse As Boolean      '对应属性"pblnInUse"

Private mobj记忆 As cls用户操作记忆 '修改：2001-12-29（增加该对象）。

Private mintFixed As Integer
Private mcol体检项目 As Collection

'功能：表明当前窗体是否一加载，以便主导航界面判断当前窗体是否已执行过Form_Load。
Public Property Get pblnInUse() As Boolean
Attribute pblnInUse.VB_Description = "'功能：表明当前窗体是否一加载，以便主导航界面判断当前窗体是否已执行过Form_Load。"
    pblnInUse = mblnInUse
End Property

Private Sub cchkEnd_Click(Index As Integer)
    If Index = 0 Then
        ccmdSingleQuery_Click
    Else
        ccmdAdd_Click
    End If
End Sub



'2006-6-19(网格录入）
Private Sub cchkGrid_Click()
    Dim j As Long
    
    On Error Resume Next
    subResizeTab
    
    For j = mintFixed + 1 To cgrdPerson.Cols - 1
        cgrdPerson.ColHidden(j) = IIf(cchkGrid.Value = 0, True, False)
    Next
    
    
End Sub

Private Sub cchkUnEnd_Click(Index As Integer)
    If Index = 0 Then
        ccmdSingleQuery_Click
    Else
        ccmdAdd_Click
    End If
End Sub

Private Sub ccmbSheet_Click()
    On Error Resume Next
    cgrdPerson.Rows = 1
    cgrdInput.Rows = 1
End Sub

Private Sub ccmdSingleQuery_Click()
    On Error GoTo errHandler
    
    '显示指定体检日期的未下结论的体检人员名单。
    subShowSingleList
    
    Exit Sub
    
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检界面部件", "FrmInputTestResult", "ccmdSingleQuery_Click", 6666, lstrError, False
End Sub




Private Sub cgrdInput_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lstr单项结论 As String
    On Error GoTo errHandler
    If Row > 0 Then
        lstr单项结论 = pobj业务对象.func获取单项结论(cgrdInput.TextMatrix(Row, 0), cgrdInput.TextMatrix(Row, 2))
        If lstr单项结论 = "不合格" Then
            '设置颜色。
            cgrdInput.Cell(flexcpBackColor, Row, 2, Row, 2) = &H8A5AFA
        Else
            '设置颜色。
            cgrdInput.Cell(flexcpBackColor, Row, 2, Row, 2) = vbWhite
        End If
    End If
    Exit Sub
errHandler:
End Sub

Private Sub cgrdInput_DblClick()
    '修改颜色。
    On Error Resume Next
    If cgrdInput.Row > 0 Then
        If cgrdInput.Cell(flexcpBackColor, cgrdInput.Row, 2, cgrdInput.Row, 2) = &H8A5AFA Then
            cgrdInput.Cell(flexcpBackColor, cgrdInput.Row, 2, cgrdInput.Row, 2) = vbWhite
        Else
            cgrdInput.Cell(flexcpBackColor, cgrdInput.Row, 2, cgrdInput.Row, 2) = &H8A5AFA
        End If
    End If
End Sub

Private Sub cgrdInput_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    On Error GoTo errHandler
    If Col = 2 And KeyCode = 13 Then
        '换行。
        If Row = cgrdInput.Rows - 1 Then
            cgrdInput.Row = 1
        Else
            cgrdInput.Row = cgrdInput.Row + 1
        End If
        cgrdInput.Col = 2
    End If
    Exit Sub
errHandler:

End Sub

Private Sub cgrdInput_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If cgrdInput.Col = 2 And Button = 1 Then
        cgrdInput.EditCell
    End If
End Sub
'2006-6-19(网格录入)
Private Sub cgrdPerson_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lstr单项结论 As String
    On Error GoTo errHandler
    If Row > 0 Then
        lstr单项结论 = pobj业务对象.func获取单项结论(cgrdPerson.TextMatrix(Row, 0), cgrdPerson.TextMatrix(Row, Col))
        If lstr单项结论 = "不合格" Then
            '设置颜色。
            cgrdPerson.Cell(flexcpBackColor, Row, Col, Row, Col) = &H8A5AFA
        Else
            '设置颜色。
            cgrdPerson.Cell(flexcpBackColor, Row, Col, Row, Col) = vbWhite
        End If
    End If
    Exit Sub
errHandler:
End Sub

Private Sub cgrdPerson_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Row < 1 Or Col <= mintFixed Then
        Cancel = True
    End If
    
End Sub

Private Sub cgrdPerson_SelChange()
    Dim lobj体检 As Object
    Dim lstrNo As String '系统编号。
    
    On Error GoTo errHandler
    If cgrdPerson.Row < 1 Or mblnSys Then Exit Sub
    MousePointer = 11
    
    cstbMain.Panels(1).Text = "正在获取体检人员照片，请稍候..."
    
    '界面按时不可操作。
    ctbMain.Enabled = False
    ctabPerson.Enabled = False
    Frame1.Enabled = False
    
    '获取系统编号。
    lstrNo = cgrdPerson.TextMatrix(cgrdPerson.Row, 0)
    
    '创建体检对象。
    Set lobj体检 = CreateObject("体检对象.clsMedicalExam")
    
    lobj体检.系统编号 = lstrNo
        
    '显示相片。
    If Not lobj体检.体检人员.像片 Is Nothing Then
        cpicPhoto(1).Picture = lobj体检.体检人员.像片
    Else
        cpicPhoto(1).Picture = Nothing
    End If
    
errHandler:
    If Err <> 0 Then
        Dim lstrError As String
        lstrError = func错误处理(Err.Number, Err.Description)
        sfsub错误处理 "体检界面部件", "FrmInputTestResult", "cgrdPerson_Click", 6666, lstrError, False
    End If
    Set lobj体检 = Nothing
    '界面恢复可以操作。
    ctbMain.Enabled = True
    ctabPerson.Enabled = True
    ccmdRemove.Enabled = True
    Frame1.Enabled = True
    MousePointer = 0
    cstbMain.Panels(1).Text = ""
    Exit Sub
    Resume

End Sub

Private Sub cgrdSingleList_SelChange()
    Dim lobj体检 As Object
    Dim lstrSysNo As String '系统编号。
    
    On Error GoTo errHandler
    If cgrdSingleList.Row < 1 Or mblnSys Then Exit Sub
    MousePointer = 11
    
    cstbMain.Panels(1).Text = "正在获取体检人员照片，请稍候..."
    
    '界面按时不可操作。
    ctbMain.Enabled = False
    ctabPerson.Enabled = False
    Frame1.Enabled = False
    
    '获取系统编号。
    lstrSysNo = cgrdSingleList.TextMatrix(cgrdSingleList.Row, 0)
    
    '创建体检对象。
    Set lobj体检 = CreateObject("体检对象.clsMedicalExam")
    
    lobj体检.系统编号 = lstrSysNo
        
    '显示基本信息。
    With lobj体检.体检人员
        clblInfo(0) = .姓名
        clblInfo(1) = .性别
        clblInfo(2) = .年龄
        clblInfo(3) = .单位名称
        ctxtSingleNo.Text = lstrSysNo
        clblInfo(4) = lobj体检.试管编号
        
        '显示相片。
        If Not .像片 Is Nothing Then
            cpicPhoto(0).Picture = .像片
        Else
            cpicPhoto(0).Picture = Nothing
        End If
    End With
    
    '设置体检结果录入网格。
    subShowInputGrid lstrSysNo
    
    ctbMain.Buttons(1).Enabled = True
    
errHandler:
    If Err <> 0 Then
        Dim lstrError As String
        lstrError = func错误处理(Err.Number, Err.Description)
        sfsub错误处理 "体检界面部件", "FrmInputTestResult", "cgrdSingleList_SelChange", 6666, lstrError, False
    End If
    '界面恢复可以操作。
    ctbMain.Enabled = True
    ctabPerson.Enabled = True
    Frame1.Enabled = True
    MousePointer = 0
    cstbMain.Panels(1).Text = ""
    
    Exit Sub
    Resume
End Sub

Private Sub coptBatchType_Click(Index As Integer)
    On Error Resume Next
    If coptBatchType(Index).Value Then
        If Index = 0 Then
            cdtpQueryDate.SetFocus
        ElseIf Index = 1 Then
            ctxt单位名称.SetFocus
        Else
            ctxtBatchNo.SetFocus
        End If
    End If
    If coptBatchType(2).Value Then
        cchk刷条码.Visible = True
    Else
        cchk刷条码.Visible = False
        cchk刷条码.Value = 0
    End If
    
End Sub

Private Sub coptType_Click(Index As Integer)
    On Error Resume Next
    Label1(3).Caption = coptType(Index).Caption
    ctxtSingleNo.Text = ""
    ctxtSingleNo.SetFocus
    If coptType(0).Value Then
        cchk刷条码.Visible = True
    Else
        cchk刷条码.Visible = False
        cchk刷条码.Value = 0
    End If
End Sub

Private Sub ctxtBatchNo_GotFocus()
    On Error Resume Next
    '若输入系统编号，设置系统编号的固定部分，方便录入。
    If ctxtBatchNo.Text = "" And cchk刷条码.Value = 0 And cchk刷条码.Visible Then
        ctxtBatchNo.Text = mstr系统编号固定部分
        ctxtBatchNo.SelLength = 0
        ctxtBatchNo.SelStart = Len(mstr系统编号固定部分)
    End If
End Sub

Private Sub ctxtBatchNo_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 And Trim(ctxtBatchNo.Text) <> "" Then
        ccmdAdd_Click
        If cchk刷条码.Value = 1 Then
            ctxtBatchNo.Text = ""
        End If
    End If
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    ctxtSingleNo.SetFocus
    ctxtSingleNo.SelStart = Len(ctxtSingleNo)
    ctxtSingleNo.SelLength = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    '处理热健。
    Select Case KeyCode
    Case vbKeyF7
        ctabPerson.Tab = 0
    Case vbKeyF8
        ctabPerson.Tab = 1
    End Select
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 39 Then
        '不允许输入“'”。
        KeyAscii = 0
    End If

End Sub

Private Sub Form_Load()
    Dim i As Long
    
    On Error GoTo errHandler
    
    If mblnInUse Then Exit Sub
    
    '设置窗体已加载标志。
    mblnInUse = True
    
    '创建mobj界面通用对象，初始化工具栏。
    Dim lcol工具栏 As Collection
    Set lcol工具栏 = New Collection
    With lcol工具栏
        .Add "保存"
        .Add "|"
        .Add "退出"
    End With
    Set mobj界面通用对象 = New cls界面通用对象
    With mobj界面通用对象
        Set .Form = Me
        Set .c工具栏 = ctbMain
        Set .c状态栏 = cstbMain
    
        .subInitialize lcol工具栏, ""
    End With
    
    '初始化体检结果录入网格。
    With cgrdInput
        .Rows = 1
        .Cols = 6
        .TextMatrix(0, 0) = "编码"
        .ColWidth(0) = 700
        .TextMatrix(0, 1) = "名称"
        .ColWidth(1) = 1600
        .TextMatrix(0, 2) = "体检结果"
        .ColWidth(2) = 1000
        .ColHidden(3) = True '存放各体检项目的枚举来源。
        .ColWidth(4) = 1200
        .TextMatrix(0, 4) = "标准值"
        .ColWidth(5) = 600
        .TextMatrix(0, 5) = "单位"
        .ShowComboButton = True
        
        .Editable = True
    End With
    '创建本窗体的全局对象mobj体检医师。
    Set mobj体检医师 = CreateObject("体检对象.clsMedicalExaminer")
    mobj体检医师.编号 = um用户编号
    
    '获取系统编号固定部分。
    Dim lobj体检 As Object '体检对象，获取系统编号的固定部分。
    Set lobj体检 = CreateObject("体检对象.clsMedicalExam")
    mstr系统编号固定部分 = lobj体检.系统编号固定部分
    
    ctxtSingleNo.Text = mstr系统编号固定部分
    
    '体检医师"显示框显示当前用户名。
    ctxtDoctor.Text = um用户名
    cdtpInputDate.Value = Date
    cdtpQueryDate.Value = Date
    cdtpSingleQuery.Value = Date
    
    cgrdInput.Rows = 1
    ctxtSingleNo.TabIndex = 0
    
        
    '修改：2001-11-2（杨春）。显示所有体检表名称。
    Dim lobj体检表模板集 As Object
    Dim lcolInfo As Collection
    Set lobj体检表模板集 = CreateObject("体检对象.ClsMedicalExamTemplateSet")
    Set lcolInfo = lobj体检表模板集.元素集
    
    ccmbSheet.Clear
    '修改：2002-8-14（杨春）增加“<所有>”选择项。
    If lcolInfo.Count > 0 Then
        ccmbSheet.AddItem "<所有>"
    End If
    For i = 1 To lcolInfo.Count
        ccmbSheet.AddItem lcolInfo(i)
    Next i
    If ccmbSheet.ListCount > 1 Then
        ccmbSheet.ListIndex = 1
    Else
        cstbMain.Panels(1).Text = "请先进入“体检表设置”操作界面设置体检表。"
    End If
    Set lcolInfo = Nothing
    Set lobj体检表模板集 = Nothing
    
    '修改：2001-12-29（获取操作记忆值）。
    On Error Resume Next
    Set mobj记忆 = New cls用户操作记忆
    mobj记忆.用户编号 = um用户编号
    mobj记忆.业务名 = "体检管理"
    
    '修改：2002-9-26（杨春）获取刷条码操作记忆值。
    If mobj记忆.记忆项值("录入结果时刷条码") = "是" Then
        cchk刷条码.Value = 1
    Else
        cchk刷条码.Value = 0
    End If
    
    cchkGrid.Visible = False
    ctabPerson.Tab = 0
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检界面部件", "FrmInputTestResult", "Form_Load", 6666, lstrError, False
    Exit Sub
    Resume
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    
    subResizeTab
End Sub

Private Sub subResizeTab()
    On Error Resume Next
    
    If ctabPerson.Tab = 0 Or cchkGrid.Value = 0 Then
        Frame1.Visible = True
        ctabPerson.Width = (Me.ScaleWidth - ctabPerson.Left - 120) / 2
        Frame1.Left = ctabPerson.Width + ctabPerson.Left + 60
        Frame1.Width = Me.ScaleWidth - Frame1.Left - 60
        Frame3.Left = Frame1.Left
        ctabPerson.Height = Me.ScaleHeight - ctabPerson.Top - 60
        Frame1.Height = ctabPerson.Height - Frame3.Height
        
        cgrdInput.Width = Frame1.Width - cgrdInput.Left - 60
        cgrdInput.Height = Frame1.Height - cgrdInput.Top - 60
    Else
        '网格录入。
        Frame1.Visible = False
        ctabPerson.Width = Me.ScaleWidth - ctabPerson.Left - 60
         
    End If
    Select Case ctabPerson.Tab
    Case 0
        cgrdSingleList.Left = 60
        cgrdSingleList.Width = ctabPerson.Width - cgrdSingleList.Left - 60
        cgrdSingleList.Height = ctabPerson.Height - cgrdSingleList.Top - 60
        Frame2.Width = cgrdSingleList.Width
        cpicPhoto(0).Left = ctabPerson.Width - cpicPhoto(0).Width - 120
        Frame3.Top = 600
    Case 1
        cgrdPerson.Left = 60
        cgrdPerson.Width = ctabPerson.Width - cgrdPerson.Left - 60
        cgrdPerson.Height = ctabPerson.Height - cgrdPerson.Top - 60
        'cpicPhoto(1).Left = ctabPerson.Width - cpicPhoto(1).Width - 120
        If cchkGrid.Value = 0 Then
            Frame3.Top = 600
'            cgrdPerson.SelectionMode = flexSelectionByRow
            cgrdPerson.Editable = False
        Else
            '网格录入
            Frame3.Top = 1100
'            cgrdPerson.SelectionMode = flexSelectionFree
            cgrdPerson.Editable = True
        End If
        
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    '修改：2002-9-26（杨春）保存操作记忆值。
    mobj记忆.sub覆盖记忆值 "录入结果时刷条码", IIf(cchk刷条码.Value = 1, "是", "否")
    
    '释放本界面的全局对象。
    Set mobj界面通用对象 = Nothing
    Set mobj体检医师 = Nothing
    Set mobj记忆 = Nothing
    
    '设置标志pblnInUse，表明窗体已不在使用。
    mblnInUse = False
    
End Sub

'功能：添加人员到网格。
Private Sub ccmdAdd_Click()
    On Error GoTo errHandler
    '修改：2001-11-2（杨春）判断是否有体检表，若没有，进行提示。
    If ccmbSheet.ListCount = 0 Then
        sffuncMsg "请先进入“体检表设置”操作界面设置好体检表。", sf警告
        Exit Sub
    End If
    If coptBatchType(1).Value And ctxt单位名称 = "" Then
        MsgBox "请输入单位名称！", vbOKOnly + vbExclamation, "系统提示"
        ctxt单位名称.SetFocus
        Exit Sub
    End If
    If coptBatchType(2).Value And ctxtBatchNo = "" Then
        MsgBox "请输入系统编号（或刷入体检表上的条码）！", vbOKOnly + vbExclamation, "系统提示"
        ctxtBatchNo.SetFocus
        Exit Sub
    End If
    
    MousePointer = 11
    '批量录入方式。
    
    '显示一批体检人员到网格中。
    subShowBatchPerson
        
    
    '判断“保存”按钮是否可用。
    If cgrdPerson.Rows > 1 Then
        ctbMain.Buttons(1).Enabled = True
        ccmdClear.Enabled = True
    Else
        ctbMain.Buttons(1).Enabled = False
        ccmdClear.Enabled = False
    End If
    cstbMain.Panels(1) = ""
    MousePointer = 0
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检界面部件", "FrmInputTestResult", "ccmdAdd_Click", 6666, lstrError, False
    cdtpQueryDate.SetFocus
    cstbMain.Panels(1) = ""
    MousePointer = 0
    Exit Sub
    Resume
End Sub

Private Sub cchk刷条码_Click()
    On Error GoTo errHandler
    If Not cchk刷条码.Visible Then Exit Sub
    
    If cchk刷条码.Value = 1 Then
        '清空系统编号输入框。
        ctxtSingleNo = ""
        ctxtBatchNo = ""
    Else
        '可以输入试管编号。
        If ctabPerson.Tab = 0 Then
            If ctxtSingleNo = "" Then
                ctxtSingleNo = mstr系统编号固定部分
            End If
        End If
    End If
    If ctabPerson.Tab = 0 Then
        ctxtSingleNo.SelStart = Len(ctxtSingleNo)
        ctxtSingleNo.SelLength = 0
        ctxtSingleNo.SetFocus
    Else
        If coptBatchType(2) Then
            ctxtBatchNo.SetFocus
        End If
        
    End If
    Exit Sub
errHandler:
    
End Sub

Private Sub ccmdClear_Click()
    On Error Resume Next
    cgrdPerson.Rows = 1
    cgrdInput.Rows = 1
    ctbMain.Buttons(1).Enabled = False
    ccmdClear.Enabled = False
    ccmdRemove.Enabled = False
    
    Set cpicPhoto(1).Picture = Nothing
End Sub

Private Sub ccmdRemove_Click()
    On Error Resume Next
    If cgrdPerson.Row > 0 Then
        cgrdPerson.RemoveItem cgrdPerson.Row
        If cgrdPerson.Rows = 1 Then
            cgrdInput.Rows = 1
            ctbMain.Buttons(1).Enabled = False
            ccmdRemove.Enabled = False
            ccmdClear.Enabled = False
        End If
    End If
    
End Sub

Private Sub cdtpQueryDate_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errHandler
    If KeyCode = 13 Then
        ccmdAdd_Click
    End If
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检界面部件", "FrmInputTestResult", "cdtpQueryDate_KeyDown", 6666, lstrError, False
End Sub

Private Sub cgrdInput_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim lstrEnum As String   '当前体检结果的枚举来源（以英文逗号或中文逗号隔开）。
    Dim i As Long
    
    On Error GoTo errHandler
    '只有体检结果列可以录入。
    If Col <> 2 Then
        Cancel = True
    Else
        '根据最后隐藏列存放的枚举来源设置当前单元的下拉列表。
        lstrEnum = cgrdInput.TextMatrix(cgrdInput.Row, 3)
        If lstrEnum = "" Then
            '没有枚举来源。
            cgrdInput.ColComboList(2) = ""
            
        Else
            '设置体检结果列的录入方式为下拉选择。
            cgrdInput.ColComboList(2) = "|" & lstrEnum
        End If
        
    End If
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检界面部件", "FrmInputTestResult", "cgrdInput_BeforeEdit", 6666, lstrError, False
    Exit Sub
    Resume
End Sub


Private Sub ctxtSingleNo_GotFocus()
    On Error Resume Next
    '若输入系统编号，设置系统编号的固定部分，方便录入。
    If ctxtSingleNo.Text = "" And cchk刷条码.Value = 0 And cchk刷条码.Visible Then
        ctxtSingleNo.Text = mstr系统编号固定部分
        ctxtSingleNo.SelLength = 0
        ctxtSingleNo.SelStart = Len(mstr系统编号固定部分)
        ctbMain.Buttons(1).Enabled = False
        
    End If
End Sub

Private Sub ctxtSingleNo_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errHandler
    
    If KeyCode = 13 And Trim(ctxtSingleNo.Text) <> "" Then
        '显示人员信息。
        subShowSinglePerson
            
        If cgrdInput.Rows > 1 Then
            ctbMain.Buttons(1).Enabled = True
        End If
    End If
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "体检界面部件", "FrmInputTestResult", "ctxtSingleNo_KeyDown", 6666, lstrError, False
    Exit Sub
    Resume
End Sub






Private Sub ctabPerson_Click(PreviousTab As Integer)
    On Error Resume Next
    If PreviousTab = 0 Then
        coptBatchType(0).Value = True
        cchkGrid.Visible = True
        ctbMain.Buttons(1).Enabled = True
        cdtpQueryDate.SetFocus
    Else
        coptType(0).Value = True
        cchkGrid.Visible = False
        ctxtSingleNo.SetFocus
    End If
    subResizeTab
End Sub

'功能：根据系统编号，设置体检结果录入网格。
'作者：杨春。
Private Sub subShowInputGrid(ByVal paraSysNo As String)
    Dim lcolInfo As Collection '存放当前系统编号中当前医师可以做得指定属性的体检项目及其结果。
    Dim lobjItem As Variant    'clsFactTestItem,lcolInfo中的元素。
    Dim lstrEnum As String
    Dim i As Long
    Dim j As Long
    On Error GoTo errHandler
    
    '使用优化算法获取体检项目。
    '修改：2001-4-22(杨春)。
    Dim lobjRec As Object
    Dim lobjTemp As Object
    
    '获取指定属性（常规/化验）的体检项目：clsFactTestItem(体检项目编码，体检项目名称，缺省值，枚举来源，体检结果)。
    '修改：2002-10-14（杨春）若选择所有体检表，则获取所有体检表上可作项目。
    If ctabPerson.Tab = 1 Then
        If ccmbSheet.Text = "<所有>" Then
            Set lobjRec = mobj体检医师.Func获取本人所有体检表上可作的体检项目(mstr体检项目属性)
        Else
            Set lobjRec = mobj体检医师.Func获取指定体检表上可作的体检项目(ccmbSheet.Text, mstr体检项目属性)
        End If
    Else
        Set lobjRec = mobj体检医师.Func优化的获取本人可作的体检项目(paraSysNo, mstr体检项目属性)
    End If
    
    '显示体检项目在cgrdInput中。
    cgrdInput.Rows = 1
    
    Set mcol体检项目 = New Collection
    
    If lobjRec.recordcount > 0 Then
        cgrdInput.Rows = lobjRec.recordcount + 1
        i = 1
        Do While Not lobjRec.EOF
            cgrdInput.TextMatrix(i, 0) = lobjRec!体检项目编号
            cgrdInput.TextMatrix(i, 1) = lobjRec!体检项目名称
            cgrdInput.TextMatrix(i, 2) = IIf(IsNull(lobjRec!体检结果), "", lobjRec!体检结果)
            
            '根据单项结论设置颜色。
            Dim lstr单项结论 As String
            If IIf(IsNull(lobjRec!单项结论), "", lobjRec!单项结论) = "" And cgrdInput.TextMatrix(i, 2) <> "" Then
                '重新下单项结论。
                lstr单项结论 = pobj业务对象.func获取单项结论(lobjRec!体检项目编号, IIf(IsNull(lobjRec!体检结果), "", lobjRec!体检结果))
            Else
                lstr单项结论 = IIf(IsNull(lobjRec!单项结论), "", lobjRec!单项结论)
            End If
            If lstr单项结论 = "不合格" Then
                '设置颜色。
                cgrdInput.Cell(flexcpBackColor, i, 2, i, 2) = &H8A5AFA
            Else
                '设置颜色。
                cgrdInput.Cell(flexcpBackColor, i, 2, i, 2) = vbWhite
            End If
            '将枚举来源串转换为Grid可以识别的ColComboList串（以“|”隔开）。
            lstrEnum = IIf(IsNull(lobjRec!枚举来源), "", lobjRec!枚举来源)
            lstrEnum = gffuncStrReplace(lstrEnum, ",", "|")
            lstrEnum = gffuncStrReplace(lstrEnum, "，", "|")
            cgrdInput.TextMatrix(i, 3) = lstrEnum

            cgrdInput.TextMatrix(i, 4) = IIf(IsNull(lobjRec!标准值), "", lobjRec!标准值)
            cgrdInput.TextMatrix(i, 5) = IIf(IsNull(lobjRec!单位), "", lobjRec!单位)

            '2006-6-19(为了进行网格录入，在cgrdperson后面增加列显示体检项目名称)。
            If ctabPerson.Tab = 1 Then
                For j = mintFixed + 1 To cgrdPerson.Cols - 1
                    If cgrdPerson.TextMatrix(0, j) = lobjRec!体检项目名称 Then Exit For
                Next
                If j = cgrdPerson.Cols Then
                    cgrdPerson.Cols = cgrdPerson.Cols + 1
                    cgrdPerson.TextMatrix(0, j) = lobjRec!体检项目名称
                    
                    cgrdPerson.ColHidden(j) = IIf(cchkGrid.Value = 0, True, False)
                    If lstrEnum = "" Then
                        cgrdPerson.ColComboList(j) = ""
                    Else
                        cgrdPerson.ColComboList(j) = "|" & lstrEnum
                    End If
                End If
                
                mcol体检项目.Add lobjRec("体检项目编号").Value, lobjRec!体检项目名称
            End If
            i = i + 1
            lobjRec.movenext
        Loop
        cgrdInput.Select 1, 2, 1, 2
    Else
        cgrdInput.Rows = 1
        
        Err.Raise 6666, , "对不起，该体检人员体检表上的所有" & mstr体检项目属性 & "体检项目，你都不可以操作。或许该体检人员所使用的体检表上有没有配置" & mstr体检项目属性 & "项目。请进入业务设置的“体检医师设置”检查你可操作的项目，并进入“体检表设置”检查体检表上配置的项目。"
    End If
    
    Exit Sub
errHandler:
    sfsub错误处理 "体检界面部件", "FrmInputTestResult", "subShowInputGrid", Err.Number, Err.Description, True
    Exit Sub
    Resume
End Sub

'功能：根据单个处理方式输入的编号显示体检人员信息和体检项目结果；或根据批量处理方式输入的单个编号获取体检人员信息并加入网格中，显示体检项目及其结果。
Private Sub subShowSinglePerson()
    On Error GoTo errHandler
    Dim lobj体检 As Object     '体检对象。
    Dim lobj体检集 As Object   '体检集对象，用于根据试管编号+日期获取系统编号。
    Dim lobjRec As Object
    
    Dim lstrNo As String       '系统编号或试管编号。
    Dim llngNoType As Long     '编号类型：0 系统编号/1 试管编号。
    Dim lstrSysNo As String    '系统编号。
    Dim i As Long
    
    
    '获取输入的系统编号（或试管编号）。
    lstrNo = Trim(ctxtSingleNo.Text)
    
    If lstrNo <> "" Then
        '创建体检对象。
        Set lobj体检 = CreateObject("体检对象.clsMedicalExam")
        
        '获取编号类型。
        If coptType(0).Value Then
            llngNoType = 0
        ElseIf coptType(1).Value Then
            '根据试管编号获取系统编号。
            llngNoType = 1
        ElseIf coptType(2).Value Then
            '根据体检单号获取系统编号。
            llngNoType = 2
        Else
            '根据姓名获取系统编号
            llngNoType = 3
        End If
        
        If llngNoType <> 0 Then
            lstrNo = lobj体检.func根据其他编号获取系统编号(lstrNo, llngNoType)
        End If
        
        '输入的是系统编号。
        lobj体检.系统编号 = lstrNo
        
        lstrSysNo = lobj体检.系统编号
        ctxtSingleNo.Text = lstrSysNo
        
        '清空界面。
        If ctabPerson.Tab = 0 Then
            clblInfo(0) = ""
            clblInfo(1) = ""
            clblInfo(2) = ""
            clblInfo(3) = ""
            clblInfo(4) = ""
            cpicPhoto(0).Picture = Nothing
        End If
        
        '判断是否存在。
        If Not lobj体检.是否已存在 Then
            Err.Raise 6666, , "不存在你输入编号的体检人员。请重新输入。"
        End If
        
        '判断是否已下体检结论。
        If lobj体检.体检状态 = P_ENDED_STATUS Then
            Err.Raise 6666, , "你输入编号的体检已被医师确定了体检结论，不允许修改已下了体检结论的体检结果。" & Chr(13) & Chr(10) & "若确实要修改，请下结论的医师进入“体检结论录入”操作界面先取消下结论，再回到此操作界面修改。"
        End If
        
        '显示人员信息（包括相片）。
        If ctabPerson.Tab = 0 Then
            '单个处理方式。
            With lobj体检.体检人员
                clblInfo(0) = .姓名
                clblInfo(1) = .性别
                clblInfo(2) = .年龄
                clblInfo(3) = .单位名称
                If llngNoType = 1 Then '系统编号输入方式，需要显示试管编号。
                    clblInfo(4) = lobj体检.体检单号
                    Label1(8).Caption = "体检单号："
                Else
                    clblInfo(4) = lobj体检.试管编号
                    Label1(8).Caption = "试管编号："
                End If
                
                '显示相片。
                If Not .像片 Is Nothing Then
                    cpicPhoto(0).Picture = .像片
                End If
            End With
            
            '设置体检结果录入网格。
            subShowInputGrid lstrSysNo
            
            cgrdSingleList.Row = 0
            
        Else
            '修改：2001-11-2（杨春）因为只查询指定体检表的体检记录，所以不需判断体检表是否相同。
            
            '批量处理方式，把人员信息加入到cgrdPerson中（注意检查体检表是相同的）。
            If cgrdPerson.Rows = 1 Then
'                '若cgrdPerson中原没有记录，设置mstr体检表名称。
'                mstr体检表名称 = lobj体检.体检表.体检表名
'
'                '设置体检结果录入网格
                subShowInputGrid lstrSysNo
            Else
'                '判断体检人员的体检表名是否一致。
                '修改：2002-8-14（杨春）体检表可以选择所有。
                If ccmbSheet.Text <> "<所有>" Then
                    If ccmbSheet.Text <> lobj体检.体检表.体检表名 Then
                        Err.Raise 6666, , "你输入编号体检的体检表“" & lobj体检.体检表.体检表名 & "”与指定体检表不一致，不能批量录入体检表不相同的体检结果。"
                    End If
                End If
            End If

            '判断该人员是否已在网格中，若不在则可以加入网格。
            For i = 1 To cgrdPerson.Rows - 1
                If cgrdPerson.TextMatrix(i, 0) = lstrSysNo Then
                    '已在网格中存在，不再加入。
                    Exit Sub
                End If
            Next
            
            '把人员添加到体检人员网格中。
            cgrdPerson.Rows = cgrdPerson.Rows + 1
            
            i = cgrdPerson.Rows - 1
            cgrdPerson.TextMatrix(i, 0) = lstrSysNo
            
            '修改：2002-10-11（杨春）增加显示试管编号。
            cgrdPerson.TextMatrix(i, 1) = lobj体检.试管编号
            With lobj体检.体检人员
                cgrdPerson.TextMatrix(i, 2) = .姓名
                cgrdPerson.TextMatrix(i, 3) = .性别
                cgrdPerson.TextMatrix(i, 4) = .单位名称
                cgrdPerson.TextMatrix(i, 5) = .年龄
            End With
            
        End If
        
    End If

    ctbMain.Buttons(1).Enabled = True
    cstbMain.Panels(1) = ""
    cgrdInput.Row = 1
    cgrdInput.Col = 2
    cgrdInput.SetFocus
    SendKeys ""
    Exit Sub
errHandler:
    If ctabPerson.Tab = 0 Then
        ctbMain.Buttons(1).Enabled = False
    End If
    sfsub错误处理 "体检界面部件", "FrmInputTestResult", "subShowSinglePerson", Err.Number, Err.Description, True
    Exit Sub
    Resume
End Sub

'功能：把一批（指定编号范围，或体检日期）的体检人员加入网格，并显示体检项目及其结果。
Private Sub subShowBatchPerson()
    Dim lobjRec As Object        '通过业务对象获取的指定范围内可以录入体检结果的体检记录。
    Dim llngNoType  As Integer   '编号方式：0系统编号/1试管编号。
    Dim llngStartRow As Long     '当前体检人员网格的最大行+1。
    Dim llngRow As Long          '当前添加的行。
    Dim i As Long
    Dim lobjResult As Object
    
    On Error GoTo errHandler
    cstbMain.Panels(1) = "正在获取体检记录，请稍候..."
    
    '获取编号类型。
    llngNoType = 0 '系统编号。
    
    
    '输入的是体检日期。
    '修改：2001-11-2（增加查询参数：体检表名称）。
    '修改：2002-8-14（杨春）体检表可以选择所有。
    Set lobjRec = pobj业务对象.Func获取可修改的体检记录(IIf(coptBatchType(2).Value, ctxtBatchNo, ""), IIf(coptBatchType(2).Value, ctxtBatchNo, ""), IIf(coptBatchType(0).Value, Str(cdtpQueryDate.Value), ""), llngNoType, IIf(ccmbSheet.Text = "<所有>", "", ccmbSheet.Text), IIf(coptBatchType(1).Value, ctxt单位名称.Text, ""))

    If lobjRec.recordcount > 0 Then
   
        lobjRec.Filter = ""
        If cchkUnEnd(1).Value = 1 And cchkEnd(1).Value = 0 Then
            lobjRec.Filter = "体检状态<>2"
        ElseIf cchkUnEnd(1).Value = 0 And cchkEnd(1).Value = 1 Then
            lobjRec.Filter = "体检状态=2"
        ElseIf cchkUnEnd(1).Value = 0 And cchkEnd(1).Value = 0 Then
            lobjRec.Filter = "体检状态=-1"
        End If
        
        cgrdPerson.Redraw = False
        mblnSys = True
        mintFixed = 6
        If cgrdPerson.Rows = 1 And lobjRec.recordcount > 0 Then
            '修改：2001-11-2（杨春）不需要判断体检表名称。
            '设置体检结果录入网格中。
            subShowInputGrid lobjRec!系统编号
        End If
        
        '显示人员信息到cgrdPerson中（注意检查体检表是相同的）。
        llngStartRow = cgrdPerson.Rows - 1
        Do While Not lobjRec.EOF
            '修改：2001-11-2（杨春）不需要判断体检表名称。
    '        If lobjRec!体检表名称 = mstr体检表名称 Then
                '判断该人员是否已在网格中，若不在则可以加入网格。
                For i = 1 To llngStartRow
                    If cgrdPerson.TextMatrix(i, 0) = lobjRec!系统编号 Then
                        '已在网格中存在，不再加入。
                        GoTo LabelContinue
                    End If
                Next
                cgrdPerson.AddItem ""
                llngRow = cgrdPerson.Rows - 1
                With cgrdPerson
                    .TextMatrix(llngRow, 0) = lobjRec!系统编号
                    
                    '修改：2002-10-11（杨春）增加显示试管编号。
                    .TextMatrix(llngRow, 1) = IIf(IsNull(lobjRec!试管编号), "", lobjRec!试管编号)
                    .TextMatrix(llngRow, 2) = IIf(IsNull(lobjRec!姓名), "", lobjRec!姓名)
                    .TextMatrix(llngRow, 3) = IIf(IsNull(lobjRec!性别), "", lobjRec!性别)
                    .TextMatrix(llngRow, 4) = IIf(IsNull(lobjRec!单位名称), "", lobjRec!单位名称)
                    .TextMatrix(llngRow, 5) = IIf(IsNull(lobjRec!年龄), "", lobjRec!年龄)
                    .TextMatrix(llngRow, 6) = IIf(IsNull(lobjRec!体检单号), "", lobjRec!体检单号)
                    
                    If lobjRec!体检状态 = 2 Then
                        .Cell(flexcpBackColor, llngRow, 0, llngRow, mintFixed) = cchkEnd(1).BackColor
                    Else
                        .Cell(flexcpBackColor, llngRow, 0, llngRow, mintFixed) = cchkUnEnd(1).BackColor
                    End If
                    
                    '2006-6-19(网格录入）
                    'If cchkGrid.Value = 1 Then
                        '获取该人的所有体检结果。
                        subShowPersonResult llngRow, lobjRec!系统编号
                    'End If
                End With
             
LabelContinue:  lobjRec.movenext
        Loop
    End If
    
    If cgrdPerson.Rows > 1 Then
        ccmdRemove.Enabled = True
        ccmdClear.Enabled = True
    Else
        ccmdRemove.Enabled = False
        ccmdClear.Enabled = False
    End If
    cgrdPerson.Redraw = True
    
    On Error Resume Next
    cgrdPerson.AutoSize 0, cgrdPerson.Cols - 1
    mblnSys = False
    cstbMain.Panels(1) = ""
    Exit Sub
errHandler:
    sfsub错误处理 "体检界面部件", "FrmInputTestResult", "subShowBatchPerson", Err.Number, Err.Description, True
    mblnSys = False
    Exit Sub
    Resume
End Sub

Private Sub subShowPersonResult(ByVal paraRow As Long, ByVal para系统编号 As String)
    Dim i As Long
    Dim lobjResult As Object
    
    
    Set lobjResult = pobj业务对象.func获取体检结果(para系统编号)
    Do While Not lobjResult.EOF
        For i = mintFixed + 1 To cgrdPerson.Cols - 1
            If cgrdPerson.TextMatrix(0, i) = lobjResult!体检项目名称 Then
                cgrdPerson.TextMatrix(paraRow, i) = IIf(IIf(IsNull(lobjResult!体检结果), "", lobjResult!体检结果) = "", lobjResult!缺省值, lobjResult!体检结果)
                '设置颜色。
                If IIf(IsNull(lobjResult!单项结论), "", lobjResult!单项结论) = "不合格" Then
                    cgrdPerson.Cell(flexcpBackColor, paraRow, i, paraRow, i) = &H8A5AFA
                Else
                    cgrdPerson.Cell(flexcpBackColor, paraRow, i, paraRow, i) = vbWhite
                End If
                Exit For
            End If
        Next
        lobjResult.movenext
    Loop

End Sub


Private Sub mobj界面通用对象_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    On Error GoTo errHandler
    Dim lcolNo As Collection     '系统编号集合。
    Dim lcolResult As Collection '体检结果集合，item:[体检项目，体检结果]。
    Dim lcolItem As Collection   '单个体检项目的体检结果：[体检项目，体检结果]。
    Dim lcolDetail As Collection
    Dim lblnNotOver As Boolean
    Dim i As Long
    Dim j As Long
    
    Select Case Operate
    Case "保存"
        MousePointer = 11
        cstbMain.Panels(1) = "正在保存，请稍候..."
        
        '暂时界面不能操作。
        ctbMain.Enabled = False
        ctabPerson.Enabled = False
        Frame1.Enabled = False
        cgrdInput.Select 1, 0, 1, 0
        
        Set lcolNo = New Collection
        Set lcolResult = New Collection
        
        '获取体检结果集合：[体检项目，体检结果]。
        lblnNotOver = False
                
        If ctabPerson.Tab = 0 Then
            '若是单个录入方式，把ctxtSingleNo.text对应的系统编号加入集合lcolNo中；
            '输入的是系统编号。
            lcolNo.Add ctxtSingleNo.Text
        Else
            '若是批量录入方式，把cgrdPerson网格中所有行的系统编号加入lcolNo中。
            For i = 1 To cgrdPerson.Rows - 1
                lcolNo.Add cgrdPerson.TextMatrix(i, 0)
                
                '2006-6-19(网格录入)
                Set lcolDetail = New Collection
                If cchkGrid.Value = 1 Then
                   
                    For j = mintFixed + 1 To cgrdPerson.Cols - 1
                        Set lcolItem = New Collection
                        lcolItem.Add mcol体检项目(cgrdPerson.TextMatrix(0, j)), "体检项目"
                        lcolItem.Add cgrdPerson.TextMatrix(i, j), "体检结果"
                        
                        If cgrdPerson.TextMatrix(i, j) = "" Then
                            lblnNotOver = True
                            lcolItem.Add "", "单项结论"
                        ElseIf cgrdPerson.Cell(flexcpBackColor, i, j, i, j) = &H8A5AFA Then
                            lcolItem.Add "不合格", "单项结论"
                        Else
                            lcolItem.Add "合格", "单项结论"
                        End If
                        
                        lcolDetail.Add lcolItem, lcolItem("体检项目")
                    Next
                    lcolResult.Add lcolDetail, cgrdPerson.TextMatrix(i, 0)
                End If
            Next
        End If
        
        If lcolNo.Count = 0 Then
            Err.Raise 6666, , "请选择体检人员，并录入体检结果后，再按“保存”。"
        End If
        
        If ctabPerson.Tab = 0 Or cchkGrid.Value = 0 Then
            
            For i = 1 To cgrdInput.Rows - 1
                Set lcolItem = New Collection
                lcolItem.Add cgrdInput.TextMatrix(i, 0), "体检项目"
                lcolItem.Add cgrdInput.TextMatrix(i, 2), "体检结果"
                
                '记录没有录完。
                If cgrdInput.TextMatrix(i, 2) = "" Then
                    lblnNotOver = True
                    lcolItem.Add "", "单项结论"
                ElseIf cgrdInput.Cell(flexcpBackColor, i, 2, i, 2) = &H8A5AFA Then
                    lcolItem.Add "不合格", "单项结论"
                Else
                    lcolItem.Add "合格", "单项结论"
                End If
                lcolResult.Add lcolItem, lcolItem("体检项目")
            Next
        End If
        
        '若没有录完，进行提示。
        If lblnNotOver Then
            If Not sffuncMsg("你没有录完所有体检项目的体检结果，是否坚持要保存？", sf询问) Then
                '用户选择不保存。
                GoTo errHandler
            End If
        End If
        
        '使用优化的算法保存体检结果。
        If ctabPerson.Tab = 0 Or cchkGrid.Value = 0 Then
            pobj业务对象.Sub优化的填写体检结果 lcolNo, lcolResult, um用户编号, cdtpInputDate.Value
        Else
            pobj业务对象.Sub批量填写体检结果 lcolNo, lcolResult, um用户编号, cdtpInputDate.Value
        End If
        '恢复界面。
        ctbMain.Buttons(1).Enabled = False
        ctbMain.Enabled = True
        ctabPerson.Enabled = True
        Frame1.Enabled = True
        
        '清空界面。
        cgrdInput.Rows = 1
        If ctabPerson.Tab = 0 Then
            ctxtSingleNo.Text = ""
            clblInfo(0).Caption = ""
            clblInfo(1).Caption = ""
            clblInfo(2).Caption = ""
            clblInfo(3).Caption = ""
            clblInfo(4) = ""
            cpicPhoto(0).Picture = Nothing
            
            ctxtSingleNo.SetFocus
            If cgrdSingleList.Row > 0 And cchkEnd(0).Value = 0 Then
                cgrdSingleList.RemoveItem cgrdSingleList.Row
            ElseIf cgrdSingleList.Row > 0 Then
                cgrdSingleList.Cell(flexcpBackColor, cgrdSingleList.Row, 0, cgrdSingleList.Row, cgrdSingleList.Cols - 1) = cchkEnd(0).BackColor
            End If
            
            cgrdSingleList.Row = 0
        Else
            ccmdClear_Click
            cdtpQueryDate.SetFocus
        End If
        
        MousePointer = 0
        cstbMain.Panels(1) = "保存成功！"
        Cancel = True
    End Select
    Exit Sub
    
errHandler:
    If Err.Number <> 0 Then
        Dim lstrError As String
        lstrError = func错误处理(Err.Number, Err.Description)
        sfsub错误处理 "体检界面部件", "FrmInputTestResult", "mobj界面通用对象_BeforeOperate", 6666, lstrError, False
    End If
    If Operate = "保存" Then
        '恢复界面可以操作。
        ctbMain.Enabled = True
        ctabPerson.Enabled = True
        Frame1.Enabled = True
    End If
    MousePointer = 0
    cstbMain.Panels(1) = ""
    Cancel = True
    Exit Sub
    Resume
End Sub

'功能：把一批（指定编号范围，或体检日期）的体检人员加入网格，并显示体检项目及其结果。
Private Sub subShowSingleList()
    Dim lobjRec As Object        '通过业务对象获取的指定范围内可以录入体检结果的体检记录。
    Dim llngStartRow As Long     '当前体检人员网格的最大行+1。
    Dim llngRow As Long          '当前添加的行。
    Dim i As Long
    
    On Error GoTo errHandler
    cstbMain.Panels(1) = "正在获取体检记录，请稍候..."
    
    cgrdSingleList.Rows = 1
    
    '输入的是体检日期。
    Set lobjRec = pobj业务对象.Func获取可修改的体检记录("", "", Str(cdtpSingleQuery.Value), 0, "")
    If lobjRec.recordcount = 0 Then
        cstbMain.Panels(1) = ""
        Exit Sub
    End If
    lobjRec.Filter = ""
    If cchkUnEnd(0).Value = 1 And cchkEnd(0).Value = 0 Then
        lobjRec.Filter = "体检状态<>2"
    ElseIf cchkUnEnd(0).Value = 0 And cchkEnd(0).Value = 1 Then
        lobjRec.Filter = "体检状态=2"
    ElseIf cchkUnEnd(0).Value = 0 And cchkEnd(0).Value = 0 Then
        lobjRec.Filter = "体检状态=-1"
    End If
    cgrdSingleList.Redraw = False
    mblnSys = True
    
    '显示人员信息到cgrdSingleList中（注意检查体检表是相同的）。
    cgrdSingleList.Rows = 1
    
    llngStartRow = cgrdSingleList.Rows - 1
    Do While Not lobjRec.EOF
        cgrdSingleList.AddItem ""
        llngRow = cgrdSingleList.Rows - 1
        With cgrdSingleList
            .TextMatrix(llngRow, 0) = lobjRec!系统编号
            '修改：2002-10-11（杨春）增加显示试管编号。
            .TextMatrix(llngRow, 1) = IIf(IsNull(lobjRec!试管编号), "", lobjRec!试管编号)
            .TextMatrix(llngRow, 2) = IIf(IsNull(lobjRec!姓名), "", lobjRec!姓名)
            .TextMatrix(llngRow, 3) = IIf(IsNull(lobjRec!性别), "", lobjRec!性别)
            .TextMatrix(llngRow, 4) = IIf(IsNull(lobjRec!单位名称), "", lobjRec!单位名称)
            .TextMatrix(llngRow, 5) = IIf(IsNull(lobjRec!年龄), "", lobjRec!年龄)
            .TextMatrix(llngRow, 6) = IIf(IsNull(lobjRec!体检单号), "", lobjRec!体检单号)
            If lobjRec!体检状态 = 2 Then
                .Cell(flexcpBackColor, llngRow, 0, llngRow, .Cols - 1) = cchkEnd(0).BackColor
            Else
                .Cell(flexcpBackColor, llngRow, 0, llngRow, .Cols - 1) = cchkUnEnd(0).BackColor
            End If
        End With
        
        lobjRec.movenext
    Loop
    cgrdSingleList.Redraw = True
    cgrdSingleList.Row = 0
    On Error Resume Next
    cgrdSingleList.AutoSize 0, cgrdSingleList.Cols - 1
    mblnSys = False
    cstbMain.Panels(1) = ""
    Exit Sub
errHandler:
    mblnSys = False
    sfsub错误处理 "体检界面部件", "FrmInputTestResult", "subShowSingleList", Err.Number, Err.Description, True
    Exit Sub
    Resume
End Sub


