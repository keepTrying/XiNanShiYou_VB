VERSION 5.00
Object = "{8099FCC2-0A81-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frm内部收费 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "内部收费"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11055
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   11055
   StartUpPosition =   3  '窗口缺省
   Begin VSFlex6Ctl.vsFlexGrid cind字典 
      Height          =   3615
      Index           =   1
      Left            =   3240
      TabIndex        =   41
      Tag             =   "收费标准"
      Top             =   2400
      Visible         =   0   'False
      Width           =   6210
      _cx             =   4205258
      _cy             =   4200680
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
      BackColor       =   14809599
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   14809599
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
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
      Ellipsis        =   1
      ExplorerBar     =   1
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   1
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
   Begin VSFlex6Ctl.vsFlexGrid cind字典 
      Height          =   3615
      Index           =   2
      Left            =   3000
      TabIndex        =   73
      Tag             =   "收费项目"
      Top             =   2520
      Visible         =   0   'False
      Width           =   6210
      _cx             =   24521418
      _cy             =   24516840
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
      BackColor       =   14809599
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   14809599
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
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
      Ellipsis        =   1
      ExplorerBar     =   1
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   1
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
   Begin VSFlex6Ctl.vsFlexGrid cind字典 
      Height          =   3615
      Index           =   0
      Left            =   2880
      TabIndex        =   40
      Tag             =   "收费项目"
      Top             =   2400
      Visible         =   0   'False
      Width           =   6210
      _cx             =   24521418
      _cy             =   24516840
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
      BackColor       =   14809599
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   14809599
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
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
      Ellipsis        =   1
      ExplorerBar     =   1
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   1
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
   Begin VB.Timer Ctim 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3960
      Top             =   720
   End
   Begin MSComctlLib.ImageList cimg按钮图标 
      Left            =   9930
      Top             =   7260
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   21
      ImageHeight     =   21
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Frame Frame7 
      Appearance      =   0  'Flat
      Caption         =   "费用计算"
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   2040
      TabIndex        =   25
      Top             =   6840
      Width           =   8970
      Begin VB.ComboBox cmb交费方式 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4020
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   570
         Width           =   1845
      End
      Begin VB.TextBox cinb收费输入 
         BackColor       =   &H00F0F0F0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "yyyy-M-d"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   3
         EndProperty
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Index           =   24
         Left            =   1005
         Locked          =   -1  'True
         TabIndex        =   56
         Top             =   570
         Width           =   1455
      End
      Begin VB.TextBox cinb收费输入 
         BackColor       =   &H00F0F0F0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """￥""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   23
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   57
         Top             =   570
         Width           =   1455
      End
      Begin VB.TextBox cinb收费输入 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """￥""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   22
         Left            =   7200
         MaxLength       =   12
         TabIndex        =   54
         Top             =   180
         Width           =   1455
      End
      Begin VB.TextBox cinb收费输入 
         BackColor       =   &H00F0F0F0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Index           =   21
         Left            =   4020
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   180
         Width           =   1845
      End
      Begin VB.TextBox cinb收费输入 
         BackColor       =   &H00F0F0F0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """￥""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Index           =   20
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   180
         Width           =   1455
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "交费日期"
         Height          =   180
         Index           =   25
         Left            =   150
         TabIndex        =   59
         Top             =   630
         Width           =   720
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "交费方式"
         Height          =   180
         Index           =   24
         Left            =   2805
         TabIndex        =   58
         Top             =   600
         Width           =   720
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "找补金额"
         Height          =   180
         Index           =   23
         Left            =   6345
         TabIndex        =   53
         Top             =   630
         Width           =   720
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "实收金额"
         Height          =   180
         Index           =   22
         Left            =   6345
         TabIndex        =   52
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "应收金额大写"
         Height          =   180
         Index           =   21
         Left            =   2805
         TabIndex        =   50
         Top             =   360
         Width           =   1080
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "应收金额"
         Height          =   180
         Index           =   20
         Left            =   150
         TabIndex        =   48
         Top             =   360
         Width           =   720
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "打折情况"
      Height          =   930
      Left            =   0
      TabIndex        =   24
      Top             =   6840
      Width           =   1920
      Begin VB.TextBox cinb收费输入 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         Height          =   300
         Index           =   19
         Left            =   1005
         MaxLength       =   4
         TabIndex        =   7
         Text            =   "1.00"
         Top             =   225
         Width           =   480
      End
      Begin VB.CheckBox cchk打印打折比率 
         Caption         =   "打印打折比率"
         Height          =   195
         Left            =   165
         TabIndex        =   36
         Top             =   600
         Width           =   1455
      End
      Begin MSComCtl2.UpDown cupd修改打折比率 
         Height          =   360
         Left            =   1500
         TabIndex        =   37
         Top             =   195
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "打折比率"
         Height          =   180
         Index           =   19
         Left            =   165
         TabIndex        =   69
         Top             =   285
         Width           =   720
      End
   End
   Begin MSComctlLib.StatusBar cstuShoufei 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   22
      Top             =   7380
      Visible         =   0   'False
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   18627
            MinWidth        =   18627
            Key             =   "操作提示"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar ctlb工具栏 
      Align           =   1  'Align Top
      Height          =   555
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   979
      ButtonWidth     =   1455
      ButtonHeight    =   926
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "cimg按钮图标"
      _Version        =   393216
      Begin VB.CheckBox Cchk定位 
         Caption         =   "单位定位"
         Height          =   180
         Left            =   9240
         TabIndex        =   87
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox cchk预览 
         Caption         =   "打印前预览"
         Height          =   255
         Left            =   7440
         TabIndex        =   86
         Top             =   240
         Width           =   1335
      End
   End
   Begin TabDlg.SSTab ctabShoufei 
      Height          =   6075
      Left            =   0
      TabIndex        =   26
      Top             =   720
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   10716
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "收费"
      TabPicture(0)   =   "frm内部收费.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Clab业务分类"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cchk同批收费"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cing收费基本信息表"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "ccmd选择"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Copt今天"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Copt本月"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Copt所有"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Cchk按基本信息查询"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cchk内部收费"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Ccbo业务分类"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "门诊收费"
      TabPicture(1)   =   "frm内部收费.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(1)=   "Frame12"
      Tab(1).ControlCount=   2
      Begin VB.ComboBox Ccbo业务分类 
         Height          =   300
         Left            =   6600
         TabIndex        =   77
         Top             =   5625
         Width           =   2295
      End
      Begin VB.CheckBox cchk内部收费 
         Caption         =   "内部收费(&I)"
         Height          =   300
         Left            =   165
         TabIndex        =   81
         Top             =   5625
         Width           =   1290
      End
      Begin VB.CheckBox Cchk按基本信息查询 
         Caption         =   "按基本信息查询"
         Height          =   300
         Left            =   9000
         TabIndex        =   80
         Top             =   5640
         Width           =   1575
      End
      Begin VB.OptionButton Copt所有 
         Caption         =   "所有"
         Height          =   255
         Left            =   4920
         TabIndex        =   76
         Top             =   5625
         Width           =   735
      End
      Begin VB.OptionButton Copt本月 
         Caption         =   "本月"
         Height          =   255
         Left            =   3840
         TabIndex        =   75
         Top             =   5625
         Width           =   735
      End
      Begin VB.OptionButton Copt今天 
         Caption         =   "今天"
         Height          =   255
         Left            =   2760
         TabIndex        =   74
         Top             =   5625
         Width           =   735
      End
      Begin VB.CommandButton ccmd选择 
         Caption         =   "全选"
         Enabled         =   0   'False
         Height          =   270
         Left            =   1680
         TabIndex        =   72
         Top             =   5625
         Width           =   780
      End
      Begin VB.Frame Frame5 
         Appearance      =   0  'Flat
         Caption         =   "费用清修改单 "
         ForeColor       =   &H80000008&
         Height          =   4440
         Left            =   1560
         TabIndex        =   34
         Top             =   1080
         Width           =   9300
         Begin VB.ComboBox Ccbo收费项目大类 
            Height          =   300
            Left            =   1320
            TabIndex        =   4
            Top             =   240
            Width           =   2295
         End
         Begin VSFlex6Ctl.vsFlexGrid cing费用清单 
            Height          =   3660
            Index           =   0
            Left            =   120
            TabIndex        =   38
            Top             =   600
            Width           =   9045
            _cx             =   4210258
            _cy             =   4200760
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
            BackColorBkg    =   -2147483636
            BackColorAlternate=   14737632
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
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
            Rows            =   50
            Cols            =   10
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
            TabBehavior     =   1
            OwnerDraw       =   0
            Editable        =   -1  'True
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
         Begin VB.TextBox cinb收费输入 
            Height          =   300
            Index           =   6
            Left            =   7440
            TabIndex        =   6
            Top             =   240
            Width           =   1740
         End
         Begin VB.TextBox cinb收费输入 
            Height          =   300
            Index           =   5
            Left            =   4680
            TabIndex        =   5
            Top             =   240
            Width           =   1620
         End
         Begin VB.Label Clab收费项目大类 
            AutoSize        =   -1  'True
            Caption         =   "收费项目大类"
            Height          =   180
            Left            =   120
            TabIndex        =   82
            Top             =   360
            Width           =   960
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "收费标准"
            Height          =   180
            Index           =   6
            Left            =   6600
            TabIndex        =   47
            Top             =   330
            Width           =   720
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "收费项目"
            Height          =   180
            Index           =   5
            Left            =   3840
            TabIndex        =   46
            Top             =   360
            Width           =   720
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "按“Del”键可以删除当前选中的项目"
            Height          =   180
            Left            =   120
            TabIndex        =   35
            Top             =   3960
            Width           =   2970
         End
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         Caption         =   "基本信息"
         ForeColor       =   &H80000008&
         Height          =   550
         Left            =   120
         TabIndex        =   32
         Top             =   450
         Width           =   10725
         Begin VB.TextBox cinb收费输入 
            Height          =   300
            Index           =   4
            Left            =   9120
            TabIndex        =   3
            Top             =   200
            Width           =   1485
         End
         Begin VB.TextBox cinb收费输入 
            Height          =   300
            Index           =   3
            Left            =   4680
            TabIndex        =   2
            Top             =   200
            Width           =   1995
         End
         Begin VB.TextBox cinb收费输入 
            Height          =   300
            Index           =   2
            Left            =   3000
            TabIndex        =   1
            Top             =   200
            Width           =   840
         End
         Begin VB.TextBox cinb收费输入 
            Height          =   300
            Index           =   1
            Left            =   810
            TabIndex        =   0
            Top             =   200
            Width           =   1425
         End
         Begin VB.Label Clab片区 
            BackStyle       =   0  'Transparent
            Caption         =   "片区：(不详)"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   6720
            TabIndex        =   85
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "主管科室"
            Height          =   180
            Index           =   4
            Left            =   8280
            TabIndex        =   45
            Top             =   225
            Width           =   720
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "交费单位"
            Height          =   180
            Index           =   3
            Left            =   3960
            TabIndex        =   44
            Top             =   225
            Width           =   720
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "交费人"
            Height          =   180
            Index           =   2
            Left            =   2400
            TabIndex        =   43
            Top             =   225
            Width           =   540
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "收费编号"
            Height          =   180
            Index           =   1
            Left            =   105
            TabIndex        =   42
            Top             =   225
            Width           =   720
         End
      End
      Begin VB.Frame Frame12 
         Appearance      =   0  'Flat
         Caption         =   "基本信息"
         ForeColor       =   &H80000008&
         Height          =   5595
         Left            =   -74880
         TabIndex        =   29
         Top             =   360
         Width           =   3885
         Begin VB.TextBox cinb收费输入 
            Height          =   300
            Index           =   8
            Left            =   1095
            TabIndex        =   9
            Top             =   255
            Width           =   2640
         End
         Begin VB.TextBox cinb收费输入 
            Height          =   300
            Index           =   9
            Left            =   1095
            TabIndex        =   10
            Top             =   795
            Width           =   2640
         End
         Begin VB.TextBox cinb收费输入 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
            Height          =   300
            Index           =   10
            Left            =   1095
            MaxLength       =   3
            TabIndex        =   11
            Top             =   1350
            Width           =   420
         End
         Begin VB.TextBox cinb收费输入 
            Height          =   300
            Index           =   12
            Left            =   1095
            TabIndex        =   13
            Top             =   1890
            Width           =   2640
         End
         Begin VB.TextBox cinb收费输入 
            Height          =   300
            Index           =   13
            Left            =   1095
            TabIndex        =   14
            Top             =   2430
            Width           =   2640
         End
         Begin VB.TextBox cinb收费输入 
            Height          =   300
            Index           =   14
            Left            =   1095
            TabIndex        =   16
            Top             =   4185
            Width           =   720
         End
         Begin VB.TextBox cinb收费输入 
            Height          =   300
            Index           =   15
            Left            =   2760
            TabIndex        =   17
            Top             =   4200
            Width           =   960
         End
         Begin VB.TextBox cinb收费输入 
            Height          =   300
            Index           =   16
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   4680
            Width           =   2640
         End
         Begin VB.TextBox cinb收费输入 
            Height          =   300
            Index           =   11
            Left            =   2070
            TabIndex        =   12
            Top             =   1350
            Width           =   420
         End
         Begin MSComCtl2.DTPicker cdtp日期 
            Height          =   360
            Index           =   1
            Left            =   1110
            TabIndex        =   55
            Top             =   3585
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   635
            _Version        =   393216
            Format          =   20185089
            CurrentDate     =   36968
         End
         Begin MSComCtl2.DTPicker cdtp日期 
            Height          =   360
            Index           =   0
            Left            =   1095
            TabIndex        =   15
            Top             =   2985
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   635
            _Version        =   393216
            Format          =   20185089
            CurrentDate     =   36968
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "入院操作员"
            Height          =   180
            Index           =   14
            Left            =   120
            TabIndex        =   79
            Top             =   4245
            Width           =   900
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "经治医生"
            Height          =   225
            Index           =   15
            Left            =   1920
            TabIndex        =   78
            Top             =   4245
            Width           =   720
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0:男,其它:女"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   2595
            TabIndex        =   71
            Top             =   1410
            Width           =   1080
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "主管科室"
            Height          =   180
            Index           =   16
            Left            =   120
            TabIndex        =   66
            Top             =   4680
            Width           =   720
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "病种"
            Height          =   180
            Index           =   13
            Left            =   120
            TabIndex        =   65
            Top             =   2490
            Width           =   360
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "住院号"
            Height          =   180
            Index           =   12
            Left            =   120
            TabIndex        =   64
            Top             =   1950
            Width           =   540
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "性别"
            Height          =   180
            Index           =   11
            Left            =   1620
            TabIndex        =   63
            Top             =   1410
            Width           =   390
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "年龄"
            Height          =   180
            Index           =   10
            Left            =   120
            TabIndex        =   62
            Top             =   1410
            Width           =   360
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "交费人"
            Height          =   180
            Index           =   8
            Left            =   120
            TabIndex        =   61
            Top             =   330
            Width           =   540
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "交费单位"
            Height          =   180
            Index           =   9
            Left            =   120
            TabIndex        =   60
            Top             =   855
            Width           =   765
         End
         Begin VB.Label Label26 
            BackStyle       =   0  'Transparent
            Caption         =   "入院日期"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   3030
            Width           =   735
         End
         Begin VB.Label Label33 
            BackStyle       =   0  'Transparent
            Caption         =   "出院日期"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   3645
            Width           =   735
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         Caption         =   "费用清单 "
         ForeColor       =   &H80000008&
         Height          =   5595
         Left            =   -70845
         TabIndex        =   27
         Top             =   360
         Width           =   6800
         Begin VB.ComboBox Ccbo门收费大类 
            Height          =   300
            Left            =   1320
            TabIndex        =   18
            Top             =   240
            Width           =   2295
         End
         Begin VB.TextBox cinb收费输入 
            Height          =   300
            Index           =   18
            Left            =   1320
            TabIndex        =   21
            Top             =   720
            Width           =   2895
         End
         Begin VB.TextBox cinb收费输入 
            Height          =   300
            Index           =   17
            Left            =   4560
            TabIndex        =   20
            Top             =   240
            Width           =   2085
         End
         Begin VSFlex6Ctl.vsFlexGrid cing费用清单 
            Height          =   4395
            Index           =   1
            Left            =   135
            TabIndex        =   39
            Top             =   1140
            Width           =   6495
            _cx             =   4205760
            _cy             =   4202056
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
            BackColorBkg    =   -2147483636
            BackColorAlternate=   14737632
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
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
            Rows            =   50
            Cols            =   10
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
            Editable        =   -1  'True
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
         Begin VB.Label Clab门收费大类 
            AutoSize        =   -1  'True
            Caption         =   "收费项目大类"
            Height          =   180
            Left            =   120
            TabIndex        =   84
            Top             =   360
            Width           =   1080
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            Caption         =   "收费标准"
            Height          =   180
            Index           =   18
            Left            =   120
            TabIndex        =   68
            Top             =   810
            Width           =   720
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            Caption         =   "收费项目"
            Height          =   180
            Index           =   17
            Left            =   3750
            TabIndex        =   67
            Top             =   330
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "按“Del”键可以删除当前选中的项目"
            Height          =   225
            Left            =   120
            TabIndex        =   28
            Top             =   3960
            Width           =   2970
         End
      End
      Begin VSFlex6Ctl.vsFlexGrid cing收费基本信息表 
         Height          =   4440
         Left            =   120
         TabIndex        =   70
         Tag             =   "cing收费基本信息表"
         Top             =   1080
         Width           =   1365
         _cx             =   4196712
         _cy             =   4202136
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
         BackColorBkg    =   -2147483636
         BackColorAlternate=   14737632
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
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
         Rows            =   50
         Cols            =   10
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
         TabBehavior     =   1
         OwnerDraw       =   0
         Editable        =   -1  'True
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
      Begin VB.CheckBox cchk同批收费 
         Caption         =   "同批收费"
         Height          =   195
         Left            =   8880
         TabIndex        =   33
         Top             =   5280
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label Clab业务分类 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "业务分类"
         Height          =   180
         Left            =   5760
         TabIndex        =   83
         Top             =   5625
         Width           =   720
      End
   End
End
Attribute VB_Name = "frm内部收费"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Const WS_THICKFRAME = &H40000
Private Const GWL_STYLE = (-16)

Private Const 工程名 = "收费界面组件"
Private Const 模块名 = "frm收费管理"

Private Const 收费_收费编号 = 1
Private Const 收费_交费人 = 2
Private Const 收费_交费单位 = 3
Private Const 收费_主管科室 = 4
Private Const 收费_收费项目 = 5
Private Const 收费_收费标准 = 6
Private Const 门诊收费_收费编号 = 7
Private Const 门诊收费_交费单位 = 9
Private Const 门诊收费_交费人 = 8
Private Const 门诊收费_年龄 = 10
Private Const 门诊收费_性别 = 11
Private Const 门诊收费_住院号 = 12
Private Const 门诊收费_病种 = 13
Private Const 门诊收费_入院操作员 = 14
Private Const 门诊收费_经治医生 = 15
Private Const 门诊收费_主管科室 = 16
Private Const 门诊收费_收费项目 = 17
Private Const 门诊收费_收费标准 = 18
Private Const 打折比率 = 19
Private Const 应收金额 = 20
Private Const 应收金额大写 = 21
Private Const 实收金额 = 22
Private Const 找补金额 = 23
Private Const 交费日期 = 24

Private Const 收费 = 0
Private Const 门诊收费 = 1
Private Const 入院 = 0
Private Const 出院 = 1
''''''''''''''''''''''''''''''''''''''''''''
'修改人： 徐冀川修改
'功能：调整交费基本信息表格中字段内容
'时间：2001-12-20
'''''''''''''''''''''''''''''''''''''''''''
Private Const 基本信息_选择 = 0
Private Const 基本信息_收费批号 = 1
Private Const 基本信息_收费编号 = 2
Private Const 基本信息_交费人 = 3
Private Const 基本信息_交费单位名称 = 4
Private Const 基本信息_金额 = 5

Private Const 费用清单_收费项目编号 = 0
Private Const 费用清单_收费项目名称 = 1
Private Const 费用清单_单价 = 2
Private Const 费用清单_数量 = 3
'Private Const 费用清单_计量单位 = 4
Private Const 费用清单_金额 = 4

Private Const 字典_收费项目 = 0
Private Const 字典_收费标准 = 1
Private Const 字典_主管科室 = 2

Private Const 字典_收费项目编号 = 0
Private Const 字典_收费标准编号 = 0
Private Const 字典_收费标准名称 = 1
Private Const 字典_收费项目名称 = 1
Private Const 字典_名称 = 1
Private Const 字典_助记符 = 2
Private Const 字典_单价 = 3
Private Const 字典_最小单价 = 4
Private Const 字典_最大单价 = 5
Private Const 字典_计量单位 = 6
Private Const 字典_票据类型编号 = 7

Dim mrds收费项目 As Recordset   '保存获取的收费项目，用于初始化字典
Dim mrds收费标准 As Recordset   '保存获取的收费标准，用于初始化字典
Dim mrds主管科室 As Recordset   '保存获取的主管科室，用于初始化字典

Dim mrds交费方式 As Recordset   '保存获取的交费方式,用于初始下拉列表

Dim WithEvents mobj界面通用对象 As cls界面通用对象
Attribute mobj界面通用对象.VB_VarHelpID = -1

Public pblnInUse As Boolean


Dim mstrUndoCount As String          '用于保存表格中原来的字符串,以便在输入不合法时能够还原
Dim mstrUndoMoney As String          '用于保存表格中原来的字符串,以便在输入不合法时能够还原
Dim mcur最小单价 As Currency
Dim mcur最大单价 As Currency
Dim mintCurInput As Integer     '当前输入框的索引
Dim mlngX As Long               '鼠标在"cind字典"中按下的X位置
Dim mlngY As Long               '鼠标在"cind字典"中按下的X位置
Dim mstr交费单位编号 As String  '从单位定位接口得到的交费单位的编号
Dim mstr主管科室编号 As String  '收费信息中的主管科室编号
Dim mint交费方式编号 As Integer '交费方式的编号
Dim mstr收费批号 As String
Dim mcur总金额 As Currency
Dim mcur门诊收费总金额 As Currency
Dim mcur基本信息总金额 As Currency


Dim mobj收费管理 As Object
Dim mobj业务设置 As Object
Dim mobj单位档案 As Object
Dim mint打折控制 As Integer
Dim mint科目级数 As Integer
Dim msng打折比率 As Single
'何嘉新增
Dim mbln使用 As Boolean         '是否能使用系统
Dim mint是否右键 As Integer     '在交费单位文本框上是否使用了右键
Dim mstr收费编号 As String      '定义变量记录收费编号
Dim mblntemp As Boolean         '判断增加项目函数是否执行过

'修改：2002-10-17（杨春）记忆打印前预览。
Private mobj记忆  As cls用户操作记忆

Private Sub Ccbo门收费大类_Click()
On Error Resume Next
   Me.MousePointer = 11
   sub收费项目填充 (Ccbo门收费大类.Text)
   Me.MousePointer = 0
   cinb收费输入(17).SetFocus
End Sub

Private Sub Ccbo门收费大类_GotFocus()
On Error Resume Next
    cind字典(字典_收费标准).Visible = False
    cind字典(字典_收费项目).Visible = False
    cinb收费输入(17).Text = ""
End Sub

Private Sub Ccbo门收费大类_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii <> 0 Then KeyAscii = 0
End Sub

Private Sub Ccbo收费项目大类_Click()
On Error Resume Next
   Me.MousePointer = 11
   sub收费项目填充 (Ccbo收费项目大类.Text)
   Me.MousePointer = 0
   cinb收费输入(5).SetFocus
End Sub

Private Sub Ccbo收费项目大类_GotFocus()
On Error Resume Next
    cind字典(字典_收费标准).Visible = False
    cind字典(字典_收费项目).Visible = False
    cinb收费输入(收费_收费项目).Text = ""
End Sub

Private Sub Ccbo收费项目大类_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii <> 0 Then KeyAscii = 0
End Sub

'功能:更新收费项目大类
'作者:徐冀川
'时间:2002/07/01
Private Sub Sub更新收费项目大类()
On Error GoTo errhandler
    Dim lstrSql As String           '定义变量记录SQL语句
    Dim lobjRec As Object           '定义变量记录数据集
    
    If Ccbo业务分类.Text = "所有业务" Then
    
        lstrSql = "select 业务分类 from 收费管理_费用信息表 where 业务分类 is not null" & _
           " group by 业务分类  order by 业务分类 "
    
        Set lobjRec = dafuncGetData(lstrSql)
        
        Ccbo业务分类.Clear
        
        Ccbo业务分类.AddItem "所有业务"
        
        Do While Not lobjRec.EOF
            If lobjRec("业务分类") = "" Then
                Ccbo业务分类.AddItem "其它业务"
            Else
                Ccbo业务分类.AddItem lobjRec("业务分类")
            End If
            lobjRec.MoveNext
        Loop
        
        If Ccbo业务分类.ListCount > 0 Then
            Ccbo业务分类.ListIndex = 0
            Ccbo业务分类.Refresh
        End If
    End If
Exit Sub
errhandler:
    sfsub错误处理 "收费管理界面对象", "frm收费", " Sub更新收费项目大类", Err.Number, Err.Description
End Sub

Private Sub Ccbo业务分类_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii <> 0 Then KeyAscii = 0
End Sub

''''''''''''''''''''''''''''''''''''''
'增加人：徐冀川
'功能：控制界面查询条件的属性
'时间：2001/12/20
'
'''''''''''''''''''''''''''''''''''
Private Sub Cchk按基本信息查询_Click()
On Error GoTo errhandle
    If Cchk按基本信息查询.Value = vbChecked Then
        Frame4.Enabled = True
        cinb收费输入(1).SetFocus
        Copt今天.Enabled = False
        Copt本月.Enabled = False
        Copt所有.Enabled = False
        
        cinb收费输入(1).BackColor = &H80000005
        cinb收费输入(收费_交费单位).BackColor = &H80000005
        cinb收费输入(收费_交费人).BackColor = &H80000005
        cinb收费输入(收费_主管科室).BackColor = &H80000005
        
    Else
        cinb收费输入(收费_收费编号).Text = ""
        cinb收费输入(收费_交费人).Text = ""
        cinb收费输入(收费_交费单位).Text = ""
        cinb收费输入(收费_主管科室).Text = ""
        Frame4.Enabled = False
        Copt今天.Enabled = True
        Copt本月.Enabled = True
        Copt所有.Enabled = True
        Copt今天.Value = True
        Copt本月.Value = False
        Copt所有.Value = False
        
        cinb收费输入(1).BackColor = &H80000000
        cinb收费输入(收费_交费单位).BackColor = &H80000000
        cinb收费输入(收费_交费人).BackColor = &H80000000
        cinb收费输入(收费_主管科室).BackColor = &H80000000
    End If
Exit Sub
errhandle:
    sfsub错误处理 "收费管理界面对象", "frm收费", " Cchk按基本信息查询_Click", Err.Number, Err.Description
End Sub

Private Sub cchk内部收费_Click()
On Error GoTo errhandle
    Dim lstrSql As String           '定义变量记录SQL语句
    Dim lobjRec As Object           '定义变量记录数据集
    
    Frame4.Enabled = True
    sub清除界面
    cing费用清单(ctabShoufei.Tab).Rows = 1
    
    If ctabShoufei.Tab = 收费 Then
        cing收费基本信息表.Rows = 1
        cinb收费输入(收费_收费项目).Text = ""
        cinb收费输入(收费_收费标准).Text = ""
    End If
    ''''''''''''''''''''''''''''''''''''''''
    '修改人：徐冀川
    '功能：控制内部收费查询条件控件的界面属性
    '时间：2001-12-20
    '
    ''''''''''''''''''''''''''''''''''''''''''
    If cchk内部收费.Value = vbChecked Then
        '修改人:徐冀川
        '功能:处理界面元素.
        '时间:2002/06/21
        
        '根据权限判断是否报废费用信息的权限.
        '修改人:徐冀川
        '时间:2002/06/22
        '你没有权限修改和删除内部收费信息
        
        If umfunc校验用户权限("收费管理_内部收费信息修改") Then
            Frame5.ForeColor = &H80000012
            Frame5.Caption = "你可以修改和删除内部收费信息"
            Label2.Enabled = True
            
            '功能:需加增加收费项目,对界面元素的控制
            '时间:2002/07/19
            Clab收费项目大类.Visible = True
            Ccbo收费项目大类.Visible = True
            lblCaption(5).Visible = True
            cinb收费输入(5).Visible = True
            
            Clab收费项目大类.Left = 120
            Ccbo收费项目大类.Left = 1250
            Ccbo收费项目大类.Width = 1500
            lblCaption(5).Left = 2900
            cinb收费输入(5).Left = 3600
            cinb收费输入(5).Width = 750
            cinb收费输入(收费_收费项目).Enabled = True
            Frame5.Left = 5880
            Frame5.Width = 4500
            cing费用清单(0).Top = 600
'            cing费用清单(0).Height = 2200
            
            Clab收费项目大类.Enabled = False
            Ccbo收费项目大类.Enabled = False
            lblCaption(5).Enabled = False
            cinb收费输入(5).Enabled = False
            
            
        Else
            Frame5.ForeColor = &HFF&
            Frame5.Caption = "你没有权限增加、修改和删除内部收费信息"
            Label2.Enabled = False
            cing费用清单(0).Top = 300
'            cing费用清单(0).Height = 2500
            Clab收费项目大类.Visible = False
            Ccbo收费项目大类.Visible = False
            lblCaption(5).Visible = False
            cinb收费输入(5).Visible = False
            cinb收费输入(收费_收费项目).Enabled = False
        End If
        
        If umfunc校验用户权限("收费管理_内部收费信息报废") Then
            ctlb工具栏.Buttons(7).Enabled = True
        End If
        cing收费基本信息表.Visible = True
        cing收费基本信息表.Width = 5685
        cing费用清单(0).Width = 4845
        Frame5.Left = 5880
        Frame5.Width = 5100
        lblCaption(6).Visible = False
        cinb收费输入(6).Visible = False
        
        ccmd选择.Enabled = True
        Copt今天.Enabled = True
        Copt今天.Value = True
        Copt本月.Enabled = True
        Copt所有.Enabled = True
        Cchk按基本信息查询.Enabled = True
        
        'cing费用清单(收费).Enabled = False
        cinb收费输入(收费_收费编号).Enabled = True
        cinb收费输入(收费_收费编号).SetFocus
        cchk同批收费.Enabled = True
        
        If cchk同批收费.Value = vbChecked Then
            cing收费基本信息表.Enabled = True
        Else
            cing收费基本信息表.Enabled = False
        End If
        ctlb工具栏.Buttons(1).Enabled = True
        cinb收费输入(收费_收费标准).Enabled = False
        cing收费基本信息表.Enabled = True
        'cing费用清单(收费).Editable = False
        
        '功能:获取系统中对收费的业务分类
        '时间:2002/06/25
        '作者:徐冀川
        Clab业务分类.Enabled = True
        Ccbo业务分类.Enabled = True
        Ccbo业务分类.BackColor = &H80000005
        lstrSql = "select 业务分类 from 收费管理_费用信息表 where 业务分类 is not null" & _
           " group by 业务分类  order by 业务分类 "
    
        Set lobjRec = dafuncGetData(lstrSql)
        
        Ccbo业务分类.Clear
        
        Ccbo业务分类.AddItem "所有业务"
        
        'lobjRec.MoveFirst
        Do While Not lobjRec.EOF
            If lobjRec("业务分类") = "" Then
                Ccbo业务分类.AddItem "其它业务"
            Else
                Ccbo业务分类.AddItem lobjRec("业务分类")
            End If
            lobjRec.MoveNext
        Loop
        
        Ccbo业务分类.ListIndex = 0
        Ccbo业务分类.Refresh
        
        Frame4.Enabled = False
        
        '修改：2001-11-22（允许按交费单位、主管科室）查询。
        cinb收费输入(1).BackColor = &H80000000
        cinb收费输入(收费_交费单位).BackColor = &H80000000
        cinb收费输入(收费_交费人).BackColor = &H80000000
        cinb收费输入(收费_主管科室).BackColor = &H80000000
    Else

        '修改人:徐冀川
        '功能:处理界面元素.
        '时间:2002/06/21
        
        
        Clab收费项目大类.Left = 120
        Ccbo收费项目大类.Left = 1320
        Ccbo收费项目大类.Width = 2295
        lblCaption(5).Left = 3960
        cinb收费输入(5).Left = 4800
        cinb收费输入(5).Width = 1380
        Frame5.ForeColor = &H80000012
        Frame5.Caption = "费用清单修改"
        Frame5.Left = 120
        Frame5.Width = 10845
        
        ctlb工具栏.Buttons(7).Enabled = False
        cing收费基本信息表.Visible = False

        cing费用清单(0).Width = 10600
        Clab收费项目大类.Visible = True
        Ccbo收费项目大类.Visible = True
        lblCaption(5).Visible = True
        lblCaption(6).Visible = True
        cinb收费输入(5).Visible = True
        cinb收费输入(6).Visible = True
        cing费用清单(0).Top = 600
'        cing费用清单(0).Height = 2100
         
        Clab片区.Caption = "片区：(不详)"
    
        '徐冀川（修改）在取消内部收费后，Cchk按基本信息查询设置为没有选中状态
        Cchk按基本信息查询.Value = Unchecked
        
        Frame4.Enabled = True
        ccmd选择.Enabled = False
        Copt今天.Enabled = False
        Copt本月.Enabled = False
        Copt所有.Enabled = False
        Cchk按基本信息查询.Enabled = False
         
        cing费用清单(收费).Enabled = True
        cinb收费输入(收费_收费编号).Enabled = False
        cchk同批收费.Enabled = False
        cing收费基本信息表.Enabled = False
        cing费用清单(收费).Editable = True
        ctlb工具栏.Buttons(1).Enabled = False
        cinb收费输入(收费_收费项目).Enabled = True
        cinb收费输入(收费_收费标准).Enabled = True
        cinb收费输入(收费_交费人).Enabled = True
        If cinb收费输入(收费_交费人).Enabled Then
            cinb收费输入(收费_交费人).SetFocus
        ElseIf cinb收费输入(收费_收费编号).Enabled Then
            cinb收费输入(收费_收费编号).SetFocus
        End If
        cinb收费输入(交费日期).Text = Date
        cinb收费输入(收费_主管科室).Text = um用户所属科室
        cinb收费输入(收费_交费单位).Enabled = True
        cinb收费输入(收费_交费人).Enabled = True
        cinb收费输入(收费_主管科室).Enabled = True
        
        '业务分类界面控制
        '时间:2002/05/06
        '作者: 徐冀川
    
        Clab业务分类.Enabled = False
        Ccbo业务分类.Enabled = False
        Ccbo业务分类.BackColor = &H80000000
        
        cinb收费输入(1).BackColor = &H80000005
        cinb收费输入(收费_交费单位).BackColor = &H80000005
        cinb收费输入(收费_交费人).BackColor = &H80000005
        cinb收费输入(收费_主管科室).BackColor = &H80000005
            
        Clab收费项目大类.Enabled = True
        Ccbo收费项目大类.Enabled = True
        lblCaption(5).Enabled = True
            
    End If
Exit Sub
errhandle:
    sfsub错误处理 "收费管理界面对象", "frm收费", " cchk内部收费_Click", Err.Number, Err.Description
End Sub
Private Sub cchk内部收费_GotFocus()
On Error GoTo errhandle
    cind字典(字典_收费标准).Visible = False
    cind字典(字典_收费项目).Visible = False
    cind字典(字典_主管科室).Visible = False
Exit Sub
errhandle:
    sfsub错误处理 "收费管理界面对象", "frm收费", " cchk内部收费_GotFocus", Err.Number, Err.Description
End Sub

Private Sub ccmd选择_Click()
On Error GoTo errhandle
    Dim i As Long
    Dim lcur总金额 As Currency
    If ccmd选择.Caption = "全选" Then
        
        '*******以下为何嘉新增-0823********
        mcur基本信息总金额 = 0
        '*******以上为何嘉新增-0823********
        
        For i = 1 To cing收费基本信息表.Rows - 1
            cing收费基本信息表.Cell(flexcpChecked, i, 0) = 1
            mcur基本信息总金额 = (mcur基本信息总金额 + cing收费基本信息表.TextMatrix(i, 基本信息_金额))
        Next
        
        '*******以下为何嘉新增-0823********
        lcur总金额 = mcur基本信息总金额 * CDbl(cinb收费输入(19).Text)
        '*******以上为何嘉新增-0823********
        
        ccmd选择.Caption = "清除"
    Else
        For i = 1 To cing收费基本信息表.Rows - 1
            cing收费基本信息表.Cell(flexcpChecked, i, 0) = 2
        Next
        mcur基本信息总金额 = 0
        
        '********以下为何嘉新增-0823*********
        lcur总金额 = 0
        '********以上为何嘉新增-0823*********
        
        ccmd选择.Caption = "全选"
    End If
    
    '********以下为何嘉新增-0823*********
    'cinb收费输入(应收金额).Text = mcur基本信息总金额
    cinb收费输入(应收金额).Text = lcur总金额
    '********以上为何嘉新增-0823*********
Exit Sub
errhandle:
    sfsub错误处理 "收费管理界面对象", "frm收费", " ccmd选择_Click", Err.Number, Err.Description
End Sub

Private Sub ccmd选择_GotFocus()
On Error GoTo errhandle
    cind字典(字典_收费标准).Visible = False
    cind字典(字典_收费项目).Visible = False
    cind字典(字典_主管科室).Visible = False
Exit Sub
errhandle:
    sfsub错误处理 "收费管理界面对象", "frm收费", "ccmd选择_GotFocus", Err.Number, Err.Description
End Sub

Private Sub cdtp日期_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo errhandle
    If KeyCode = vbKeyReturn Then
        Select Case Index
            Case 入院
                If cdtp日期(出院).Enabled Then cdtp日期(出院).SetFocus
            Case 出院
                If cinb收费输入(门诊收费_入院操作员).Enabled Then cinb收费输入(门诊收费_入院操作员).SetFocus
        End Select
    End If
Exit Sub
errhandle:
    sfsub错误处理 "收费管理界面对象", "frm收费", "cdtp日期_KeyDown", Err.Number, Err.Description
End Sub

Private Sub cinb收费输入_Change(Index As Integer)
    On Error GoTo errhandle
    Static lcurMoney As Currency
    Static lintAge As Integer
    Static lsngR As Single

    Select Case Index
        Case 应收金额
            cinb收费输入(应收金额大写).Text = FuncConvertToCapsStr(Val(cinb收费输入(应收金额).Text))
            cinb收费输入(找补金额).Text = Format(Val(cinb收费输入(实收金额).Text) - Val(cinb收费输入(应收金额).Text), "0.00")
            
        Case 打折比率
            If cinb收费输入(打折比率).Text = vbNullString Then cinb收费输入(打折比率).Text = "1.00"
            If Val(cinb收费输入(打折比率).Text) > 1 Then cinb收费输入(打折比率).Text = "1.00"
            If Val(cinb收费输入(打折比率).Text) < 0 Then cinb收费输入(打折比率).Text = "0.00"
            If Not IsNumeric(cinb收费输入(打折比率).Text) Then cinb收费输入(打折比率).Text = "1.00"
            
            
            If ctabShoufei.Tab = 收费 And cchk内部收费.Value = 1 Then
                cinb收费输入(应收金额).Text = mcur基本信息总金额 * Val(cinb收费输入(打折比率).Text)
            Else
                cinb收费输入(应收金额).Text = mcur总金额 * Val(cinb收费输入(打折比率).Text)
            End If
            cinb收费输入(应收金额大写).Text = FuncConvertToCapsStr(Val(cinb收费输入(应收金额).Text))
            cinb收费输入(找补金额).Text = Format(Val(cinb收费输入(实收金额).Text) - Val(cinb收费输入(应收金额).Text), "0.00")
            
        Case 实收金额
            If cinb收费输入(实收金额).Text = vbNullString Then cinb收费输入(实收金额).Text = 0
            If Not IsNumeric(cinb收费输入(实收金额).Text) Then
                cinb收费输入(实收金额).Text = CStr(lcurMoney)
            Else
                lcurMoney = Val(cinb收费输入(实收金额).Text)
            End If
            cinb收费输入(找补金额).Text = Format(Val(cinb收费输入(实收金额).Text) - Val(cinb收费输入(应收金额).Text), "0.00")
            
        Case 找补金额
            If Val(cinb收费输入(找补金额).Text) < 0 Then
                cinb收费输入(找补金额).ForeColor = &HFF
            Else
                cinb收费输入(找补金额).ForeColor = &HFF0000
            End If
            
        Case 收费_主管科室, 门诊收费_主管科室
        Case 门诊收费_年龄
            If cinb收费输入(门诊收费_年龄).Text = vbNullString Then cinb收费输入(门诊收费_年龄).Text = "0"
            If IsNumeric(cinb收费输入(门诊收费_年龄).Text) Then lintAge = Fix(cinb收费输入(门诊收费_年龄).Text)
            If lintAge < 0 Then lintAge = 0
            cinb收费输入(门诊收费_年龄).Text = CStr(lintAge)
            
        Case 门诊收费_性别
            If cinb收费输入(Index).Text = "" Then Exit Sub
            If cinb收费输入(Index).Text = "女" Then Exit Sub
            If Asc(cinb收费输入(Index).Text) = Asc("0") Or cinb收费输入(Index).Text = "男" Then
                cinb收费输入(Index).Text = "男"
            Else
                cinb收费输入(Index).Text = "女"
            End If
        Case 收费_收费项目, 门诊收费_收费项目
            Call func匹配收费项目(cinb收费输入(Index))
        Case 收费_收费标准, 门诊收费_收费标准
            Call func匹配收费标准(cinb收费输入(Index))
    End Select
    
errhandle:
    If Err.Number = 0 Then Exit Sub
    sfsub错误处理 "收费界面", "frm收费", "cinb收费输入_Change", Err.Number, Err.Description
End Sub

Private Sub cinb收费输入_GotFocus(Index As Integer)
On Error GoTo errhandle
    Dim i As Long
    Dim lrdsTemp As Recordset
    Dim j As Long
        
    cinb收费输入(Index).SelStart = 0
    cinb收费输入(Index).SelLength = Len(cinb收费输入(Index).Text)
    Set lrdsTemp = Nothing
    '保存当前输入框的索引
    mintCurInput = Index
    Select Case Index
'&  ===========================| 收费项目获得焦点 |==============================
        Case 收费_收费项目, 门诊收费_收费项目
            '不让窗体预先处理键盘事件
            Me.KeyPreview = False
            '决定要显示的字典
            If Not cind字典(字典_收费项目).Visible Then
                cind字典(字典_收费标准).Visible = False
                cind字典(字典_收费项目).Visible = True
                cind字典(字典_主管科室).Visible = False
                'cinb收费输入(Index).SetFocus
            Else
                cinb收费输入(Index).SetFocus
            End If
'&  ===========================| 收费标准获得焦点 |==============================
        Case 收费_收费标准, 门诊收费_收费标准
            '不让窗体预先处理键盘事件
            Me.KeyPreview = False
            '决定要显示的字典
            If Not cind字典(字典_收费标准).Visible Then
                cind字典(字典_收费标准).Visible = True
                cind字典(字典_收费项目).Visible = False
                cind字典(字典_主管科室).Visible = False
                cinb收费输入(Index).SetFocus
            Else
                cinb收费输入(Index).SetFocus
            End If
'&  ===========================| 交费单位获得焦点 |=============================
        Case 收费_交费单位, 门诊收费_交费单位
            '***************以下何嘉修改09-12******************
            If mint是否右键 = 1 Then Exit Sub
            '***************以上何嘉修改09-12******************
            If (Index = 收费_交费单位 And ctabShoufei.Tab = 收费) Or (Index = 门诊收费_交费单位 And ctabShoufei.Tab = 门诊收费) Then
                Dim lrds打折信息 As Recordset               '单位的打折信息
                'Dim lobj打折信息 As Object
                '关闭字典
                cind字典(字典_收费标准).Visible = False
                cind字典(字典_收费项目).Visible = False
                cind字典(字典_主管科室).Visible = False
                
                '调用单位档案的定位接口获取单位信息
                '功能：单位定位的功能，根据Checkbox设置值，判断是否启动 徐冀川 2002/09/30
                If Cchk定位.Value = 1 Then
                    Set lrdsTemp = mobj单位档案.func单位简单定位(100, 100)
                    If Not (lrdsTemp Is Nothing) Then
                        If lrdsTemp.RecordCount > 0 Then
                            '显示单位名称`
                            cinb收费输入(Index).Text = lrdsTemp("单位名称")
                            If lrdsTemp("片区") = "" Then
                                Clab片区.Caption = "片区：(不详)"
                            Else
                                Clab片区.Caption = "(" + lrdsTemp("片区") + ")"
                            End If
                            '保存单位的申请编号
                            mstr交费单位编号 = lrdsTemp("申请编号")
                            '设置焦点
                            If cinb收费输入(Index).Enabled Then
                                cinb收费输入(Index).SetFocus
                            End If
                        End If
                    End If
                End If
                
                If Not (mobj业务设置 Is Nothing) Then
                    '查询打折信息
                    Set lrds打折信息 = mobj业务设置.func查询打折信息("单位编号='" & mstr交费单位编号 & "'")
                End If
                If Not (lrds打折信息 Is Nothing) Then
                    If Not (lrds打折信息.BOF And lrds打折信息.EOF) Then
                        '显示打折信息
                        'cinb收费输入(打折比率).Text = lrds打折信息("打折比率")
                        'msng打折比率 = lrds打折信息

                        If mint打折控制 = 0 Then
                            cinb收费输入(打折比率).Text = "1.00"
                        Else
                            cinb收费输入(打折比率).Text = lrds打折信息("打折比率")
                        End If
                    Else
                        cinb收费输入(打折比率).Text = "1.00"
                    End If
                Else
                    cinb收费输入(打折比率).Text = "1.00"
                End If
                Set lrds打折信息 = Nothing
                Set lrdsTemp = Nothing
            End If
        Case 收费_主管科室, 门诊收费_主管科室
            Me.KeyPreview = False
            If Not cind字典(字典_收费标准).Visible Then
                cind字典(字典_收费标准).Visible = False
                cind字典(字典_收费项目).Visible = False
                cind字典(字典_主管科室).Visible = True
                cinb收费输入(Index).SetFocus
            Else
                cinb收费输入(Index).SetFocus
            End If
'&  ===========================| 其它获得焦点 |==============================
        Case Else
            cind字典(字典_收费标准).Visible = False
            cind字典(字典_收费项目).Visible = False
            cind字典(字典_主管科室).Visible = False
    End Select
Exit Sub
errhandle:
    sfsub错误处理 "收费界面", "frm收费", "cinb收费输入_GotFocus", Err.Number, Err.Description
End Sub

'在此处理按键 "UP","DOWN"
Private Sub cinb收费输入_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo errhandle

     
     
    '功能：在mstr交费单位编号已保存有单位编号时，如果用户又采用手工输入方式，需要清空变量中保存的编号
    '时间：2002/09/30 徐冀川
    If (Index = 收费_交费单位 Or Index = 门诊收费_交费单位) And KeyCode <> 13 Then
       mstr交费单位编号 = ""
    End If

    '判断按键
    Select Case KeyCode
        '处理 "UP"
        Case vbKeyUp
            Select Case Index
                '影响字典 cind字典(字典_收费项目)
                Case 收费_收费项目, 门诊收费_收费项目
                    With cind字典(字典_收费项目)
                        If .RowSel > 1 Then
                            .RowSel = .RowSel - 1
                            .Select .RowSel, 0, .RowSel
                            .TopRow = .RowSel
                        End If
                    End With
                '影响字典 cind字典(字典_收费标准)
                Case 收费_收费标准, 门诊收费_收费标准
                    With cind字典(字典_收费标准)
                        If .RowSel > 1 Then
                            .RowSel = .RowSel - 1
                            .Select .RowSel, 0, .RowSel
                            .TopRow = .RowSel
                        End If
                    End With
                Case 收费_主管科室, 门诊收费_主管科室
                    With cind字典(字典_主管科室)
                        If .RowSel > 1 Then
                            .RowSel = .RowSel - 1
                            .Select .RowSel, 0, .RowSel
                            .TopRow = .RowSel
                        End If
                    End With
            End Select
        '处理 "DOWN"
        Case vbKeyDown
            Select Case Index
                Case 收费_收费项目, 门诊收费_收费项目
                    With cind字典(字典_收费项目)
                        If .RowSel < .Rows - 1 Then
                            .RowSel = .RowSel + 1
                            .Select .RowSel, 0, .RowSel
                            .TopRow = .RowSel
                        End If
                    End With
                Case 收费_收费标准, 门诊收费_收费标准
                    With cind字典(字典_收费标准)
                        If .RowSel < .Rows - 1 Then
                            .RowSel = .RowSel + 1
                            .Select .RowSel, 0, .RowSel
                            .TopRow = .RowSel
                        End If
                    End With
                Case 收费_主管科室, 门诊收费_主管科室
                    With cind字典(字典_主管科室)
                        If .RowSel < .Rows - 1 Then
                            .RowSel = .RowSel + 1
                            .Select .RowSel, 0, .RowSel
                            .TopRow = .RowSel
                        End If
                    End With
            End Select
        Case vbKeyEscape
            cind字典(字典_收费标准).Visible = False
            cind字典(字典_收费项目).Visible = False
            cind字典(字典_主管科室).Visible = False
            Me.KeyPreview = True
        Case Else
    End Select
Exit Sub
errhandle:
    sfsub错误处理 "收费界面", "frm收费", "cinb收费输入_KeyDown", Err.Number, Err.Description
End Sub

Private Sub cinb收费输入_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo errhandle
    Dim lrds收费标准 As Recordset
    Dim i As Long
    Dim j As Long
    Dim lcurMoney As Currency
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        '处理回车键
        Case vbKeyReturn
            KeyAscii = 0
            '*******以下为何嘉新增 09-12 **************
            mint是否右键 = 0
            '*******以上为何嘉新增 09-12 **************
            Select Case Index
'&  ===========================| 收费_收费项目, 门诊收费_收费项目 |==============================
                Case 收费_收费编号
                    mobj界面通用对象_BeforeOperate "查询", False
                    
                Case 收费_收费项目, 门诊收费_收费项目
                    
                    cinb收费输入(收费_收费项目).Text = cind字典(字典_收费项目).TextMatrix(cind字典(字典_收费项目).RowSel, 字典_收费项目名称)
                    If Not func检查项目是否已选(cind字典(字典_收费项目).TextMatrix(cind字典(字典_收费项目).RowSel, 字典_收费项目编号)) Then
                        
                        '只有在内部收费，并且有权限的情况才执行，徐冀川，2002/07/22
                        If cchk内部收费.Value = vbChecked And ctabShoufei.Caption = "收费" Then
                            '向数据库中增加收费项目
                            Dim lblntemp As Boolean         '记录函数返回值
                            lblntemp = False
                            lblntemp = func增加收费项目(mstr收费编号, cind字典(字典_收费项目).TextMatrix(cind字典(字典_收费项目).RowSel, 字典_收费项目编号), cind字典(字典_收费项目).TextMatrix(cind字典(字典_收费项目).RowSel, 字典_单价))
                            
                            
                            '功能：将修改处理发消息形式发送 时间：2002/08/05 作者：徐冀川
                            '修改:增加了详细的收费项目信息
                            
                            Dim lstrItemName As String      '定义变量记录收费项目名称
                            Dim lstrQuerySql As String      '定义变量记录记录SQL语句
                            Dim lstrPrice As String         '定义变量记录单价
                            Dim lstrIntro As String         '定义详细说明
                            
                            lstrIntro = ""
                            lstrItemName = cind字典(字典_收费项目).TextMatrix(cind字典(字典_收费项目).RowSel, 字典_收费项目名称)
                            lstrPrice = cind字典(字典_收费项目).TextMatrix(cind字典(字典_收费项目).RowSel, 字典_单价)
                            lstrIntro = "收费项目名称：" & lstrItemName & "，单价： " & lstrPrice & "元。"
                            
                            sub消息发送 mstr收费编号, "收费编号为：" & mstr收费编号 & "的费用信息的中被增加了收费项目。" & lstrIntro
                        End If
                        
                        '判断网格的刷新,徐冀川，2002/07/22
                        If (mblntemp = False) Or (mblntemp = True And lblntemp = True) Then
                            cing费用清单(ctabShoufei.Tab).AddItem cind字典(字典_收费项目).TextMatrix(cind字典(字典_收费项目).RowSel, 字典_收费项目编号) & vbTab & _
                                                       cind字典(字典_收费项目).TextMatrix(cind字典(字典_收费项目).RowSel, 字典_收费项目名称) & vbTab & _
                                                       cind字典(字典_收费项目).TextMatrix(cind字典(字典_收费项目).RowSel, 字典_单价) & vbTab & _
                                                       "1" & vbTab & _
                                                       cind字典(字典_收费项目).TextMatrix(cind字典(字典_收费项目).RowSel, 字典_单价)
                            For i = 1 To cing费用清单(ctabShoufei.Tab).Rows - 1
                                lcurMoney = lcurMoney + Val(cing费用清单(ctabShoufei.Tab).TextMatrix(i, 费用清单_金额))
                            Next
                            mcur总金额 = lcurMoney
                            If cchk内部收费.Value = 0 Then
                                cinb收费输入(应收金额) = lcurMoney * Val(cinb收费输入(打折比率).Text)
                                cinb收费输入(应收金额大写) = FuncConvertToCapsStr(Val(cinb收费输入(应收金额)))
                            End If
                            cinb收费输入(Index).SelStart = 0
                            cinb收费输入(Index).SelLength = Len(cinb收费输入(Index).Text)
                        End If
                        
                        '在成功向数据库中增加收费项目后，要重新初始化变量，徐冀川，2002/07/22
                        lblntemp = False
                        mblntemp = False
                        
                        
                        '刷新界面
                        For i = 1 To cing收费基本信息表.Rows - 1
                            If cing收费基本信息表.Cell(flexcpText, i, 2) = mstr收费编号 Then
                            cing收费基本信息表.TextMatrix(i, 5) = mcur总金额 * Val(cinb收费输入(打折比率).Text)
                            End If
                        Next
                        
                        
                        '判断是否有选中的费用信息
                        Dim lbln是否有选中数 As Boolean
                        lbln是否有选中数 = False
    
                        For i = 1 To cing收费基本信息表.Rows - 1
                            If cing收费基本信息表.Cell(flexcpChecked, i, 0) = 1 Then
                                lbln是否有选中数 = True
                                Exit For
                            End If
                        Next
    
                        If lbln是否有选中数 = True Then
                            sub界面数据刷新
                        End If
                        
                    Else
                        sffuncMsg "该收费项目已选！" & vbCrLf & "如需修改数量,请在网格中直接修改.", sf警告
                        Exit Sub
                    End If
                
                Case 收费_收费标准, 门诊收费_收费标准
                    If mobj业务设置 Is Nothing Then
                        sffuncMsg "业务对象 ""mobj业务设置"" 尚未创建！", sf警告
                        Exit Sub
                    End If
                    
                    Set lrds收费标准 = mobj收费管理.funcExecute("select a.收费项目编号,b.收费项目名称,a.单价,a.数量,b.计量单位,金额=a.单价*a.数量 from 收费管理_收费标准信息表 a,收费管理_收费项目字典表 b where b.收费项目编号=a.收费项目编号 and 收费标准名称='" & cind字典(字典_收费标准).TextMatrix(cind字典(字典_收费标准).RowSel, 字典_收费标准名称) & "'", "cls费用信息")
                    If lrds收费标准 Is Nothing Then
                        sffuncMsg "未找到指定的收费标准！", sf警告
                        Exit Sub
                    End If
                    If lrds收费标准.BOF And lrds收费标准.EOF Then
                        sffuncMsg "收费标准中无收费项目！", sf警告
                        Exit Sub
                    Else
                        lrds收费标准.MoveFirst
                        Dim llngItemCount As Long
                        For i = 0 To lrds收费标准.RecordCount - 1
                            If Not func检查项目是否已选(lrds收费标准("收费项目编号")) Then
                            
                            cing费用清单(ctabShoufei.Tab).AddItem lrds收费标准("收费项目编号") & vbTab & _
                                                                  lrds收费标准("收费项目名称") & vbTab & _
                                                                  lrds收费标准("单价") & vbTab & _
                                                                  lrds收费标准("数量") & vbTab & _
                                                                  lrds收费标准("金额")
                            llngItemCount = llngItemCount + 1
                            Else
                            End If
                            If Not lrds收费标准.EOF Then lrds收费标准.MoveNext
                        Next
                        For i = 1 To cing费用清单(ctabShoufei.Tab).Rows - 1
                            lcurMoney = lcurMoney + Val(cing费用清单(ctabShoufei.Tab).TextMatrix(i, 费用清单_金额))
                        Next
                        mcur总金额 = lcurMoney
                        cinb收费输入(应收金额) = lcurMoney * Val(cinb收费输入(打折比率).Text)
                        cinb收费输入(应收金额大写) = FuncConvertToCapsStr(Val(cinb收费输入(应收金额)))
                        cinb收费输入(Index).SelStart = 0
                        cinb收费输入(Index).SelLength = Len(cinb收费输入(Index).Text)
                        If llngItemCount = lrds收费标准.RecordCount Then
                            MsgBox "收费标准中的所有收费项目(" & llngItemCount & "条)已添加到费用清单中！" & vbCrLf & vbCrLf & "(本次共添加所有 " & lrds收费标准.RecordCount & " 条中的 " & llngItemCount & " 条收费项目。)", vbInformation, "系统提示"
                        ElseIf llngItemCount = 0 Then
                            MsgBox "收费标准中的所有收费项目在费用清单中已添加！" & vbCrLf & vbCrLf & "(本次共添加所有 " & lrds收费标准.RecordCount & " 条中的 " & llngItemCount & " 条收费项目。)", vbInformation, "系统提示"
                        Else
                            MsgBox "收费标准中部分收费项目在费用清单中已添加,其余的 " & llngItemCount & " 条已添加到费用清单！" & vbCrLf & vbCrLf & "(本次共添加所有 " & lrds收费标准.RecordCount & " 条中的 " & llngItemCount & " 条收费项目。)", vbInformation, "系统提示"
                        End If
                    End If
                
                Case 收费_主管科室, 门诊收费_主管科室
                    If cinb收费输入(Index + 1).Enabled Then
                        cinb收费输入(Index + 1).SetFocus
                    Else
                        cinb收费输入(收费_收费编号).SetFocus
                    End If
                    If cind字典(字典_主管科室).Visible Then cinb收费输入(Index).Text = cind字典(字典_主管科室).TextMatrix(cind字典(字典_主管科室).RowSel, 1)
                    mstr主管科室编号 = cind字典(字典_主管科室).TextMatrix(cind字典(字典_主管科室).RowSel, 0)
                Case 实收金额
                    Call mobj界面通用对象_BeforeOperate("收费", False)
                Case Else
                    '在收费界面中移动焦点
                    If Index < 收费_主管科室 Then
                        If cinb收费输入(Index + 1).Enabled Then cinb收费输入(Index + 1).SetFocus
                    End If
                    If Index > 收费_收费标准 And Index < 门诊收费_收费项目 Then
                        If Index = 门诊收费_病种 Then
                            If cdtp日期(0).Enabled Then cdtp日期(0).SetFocus
                        Else
                            If cinb收费输入(Index + 1).Enabled Then cinb收费输入(Index + 1).SetFocus
                        End If
                    End If
            End Select
        End Select
Exit Sub
errhandle:
    sfsub错误处理 "收费界面", "frm收费", "cinb收费输入_KeyPress", Err.Number, Err.Description
End Sub

'***************************以下为何嘉新增09-12***************************
Private Sub cinb收费输入_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errhandle
    If Index = 收费_交费单位 Or Index = 门诊收费_交费单位 Then
        If Button = 2 Then
            mint是否右键 = 1
        Else
            mint是否右键 = 0
        End If
    End If
Exit Sub
errhandle:
    sfsub错误处理 "收费界面", "frm收费", "cinb收费输入_MouseDown", Err.Number, Err.Description
End Sub
'***************************以上为何嘉新增09-12***************************

'***************************以下为何嘉新增09-12***************************
Private Sub cinb收费输入_Validate(Index As Integer, Cancel As Boolean)
On Error Resume Next
    mint是否右键 = 0
End Sub
'***************************以上为何嘉新增09-12***************************

Private Sub cind字典_DblClick(Index As Integer)
On Error Resume Next
    cinb收费输入_KeyPress mintCurInput, vbKeyReturn
End Sub

Private Sub cind字典_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Select Case KeyCode
        Case vbKeyEscape
            cind字典(Index).Visible = False
            Me.KeyPreview = True
        Case vbKeyReturn
            cind字典_DblClick (Index)
        Case Else
    End Select
End Sub



Private Sub cind字典_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    mlngX = X
    mlngY = Y
End Sub

Private Sub cind字典_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errhandle
    If (Button = vbLeftButton) And Y < (cind字典(Index).RowHeight(0) * cind字典(Index).Rows - 1) And X < cind字典(Index).ColPos(cind字典(Index).Cols - 1) + cind字典(Index).ColWidth(cind字典(Index).Cols - 1) Then
        If cind字典(Index).Top > 0 And (cind字典(Index).Top + cind字典(Index).Height) < Me.Height And cind字典(Index).Left > 0 And (cind字典(Index).Left + cind字典(Index).Width) < Me.Width Then
            cind字典(Index).Move cind字典(Index).Left + X - mlngX, cind字典(Index).Top + Y - mlngY
        End If
    End If
    If cind字典(Index).Top <= 0 Then cind字典(Index).Top = 1
    If cind字典(Index).Left <= 0 Then cind字典(Index).Left = 1
    If cind字典(Index).Top + cind字典(Index).Height >= Me.Height Then cind字典(Index).Top = Me.Height - cind字典(Index).Height - 1
    If cind字典(Index).Left + cind字典(Index).Width >= Me.Width Then cind字典(Index).Left = Me.Width - cind字典(Index).Width - 1
Exit Sub
errhandle:
    sfsub错误处理 "收费界面", "frm收费", "cind字典_MouseMove", Err.Number, Err.Description
End Sub

Private Sub cing费用清单_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    Dim lcurMoney As Currency
    
    On Error GoTo errhandle
    ctlb工具栏.Buttons("收费(&G)").Enabled = True
    Select Case Col
        Case 费用清单_数量
            '判断输入的是否数值
            If Len(cing费用清单(ctabShoufei.Tab).TextMatrix(Row, Col)) > 4 Then
                cing费用清单(ctabShoufei.Tab).TextMatrix(Row, Col) = mstrUndoCount
            Else
                If IsNumeric(cing费用清单(ctabShoufei.Tab).TextMatrix(Row, Col)) And Val(cing费用清单(ctabShoufei.Tab).TextMatrix(Row, Col)) > 0 Then   '是数值
                 '计算金额
                    cing费用清单(ctabShoufei.Tab).TextMatrix(Row, 费用清单_金额) = cing费用清单(ctabShoufei.Tab).TextMatrix(Row, 费用清单_单价) * cing费用清单(ctabShoufei.Tab).TextMatrix(Row, 费用清单_数量)
                Else                                                            '不是数值
                    'Undo
                    cing费用清单(ctabShoufei.Tab).TextMatrix(Row, Col) = mstrUndoCount
                End If
            End If
        Case 费用清单_单价
            Dim lcur单价 As Currency
            If mcur最小单价 = mcur最大单价 Then
                sffuncMsg "该收费项目单价已定,不可修改！", sf警告
                cing费用清单(ctabShoufei.Tab).TextMatrix(Row, Col) = mstrUndoMoney
                If cinb收费输入(mintCurInput).Enabled Then cinb收费输入(mintCurInput).SetFocus
                ctlb工具栏.Buttons("收费(&G)").Enabled = True
                Exit Sub
            End If
            
            If IsNumeric(cing费用清单(ctabShoufei.Tab).TextMatrix(Row, Col)) Then
                If Val(cing费用清单(ctabShoufei.Tab).TextMatrix(Row, Col)) > 0 Then
                        If Val(cing费用清单(ctabShoufei.Tab).TextMatrix(Row, Col)) <= mcur最大单价 And Val(cing费用清单(ctabShoufei.Tab).TextMatrix(Row, Col)) >= mcur最小单价 Then
                            cing费用清单(ctabShoufei.Tab).TextMatrix(Row, 费用清单_金额) = cing费用清单(ctabShoufei.Tab).TextMatrix(Row, 费用清单_单价) * cing费用清单(ctabShoufei.Tab).TextMatrix(Row, 费用清单_数量)
                        Else
                            sffuncMsg "输入的单价超出范围！", sf警告
                            cing费用清单(ctabShoufei.Tab).TextMatrix(Row, Col) = mstrUndoMoney
                        End If
                Else
                    cing费用清单(ctabShoufei.Tab).TextMatrix(Row, Col) = mstrUndoMoney
                End If
            Else
                cing费用清单(ctabShoufei.Tab).TextMatrix(Row, Col) = mstrUndoMoney
            End If
        Case Else
    End Select
    
    For i = 1 To cing费用清单(ctabShoufei.Tab).Rows - 1
        lcurMoney = lcurMoney + Val(cing费用清单(ctabShoufei.Tab).TextMatrix(i, 费用清单_金额))
    Next
    mcur总金额 = lcurMoney
    
    '判断是否有选中的费用信息有才变化

    If cchk内部收费.Value = 0 Then
        cinb收费输入(应收金额) = lcurMoney * Val(cinb收费输入(打折比率).Text)
        cinb收费输入(应收金额大写) = FuncConvertToCapsStr(Val(cinb收费输入(应收金额)))
    End If
    
    '处理数据库中的值
     If umfunc校验用户权限("收费管理_内部收费信息修改") Then
        If mstr收费编号 = "" And cing费用清单(ctabShoufei.Tab).TextMatrix(cing费用清单(ctabShoufei.Tab).RowSel, 0) = "" Then
           sffuncMsg "费用信息错误无法删除！"
           Exit Sub
        Else
           Dim lstr收费项目编号 As String       '记录收费编号
           Dim lsing As Currency                '记录单价
           Dim lCurrency As Currency            '记录金额
           Dim lcount As Long                   '记录数量
           
           lstr收费项目编号 = cing费用清单(ctabShoufei.Tab).TextMatrix(cing费用清单(ctabShoufei.Tab).RowSel, 0)
           lsing = cing费用清单(ctabShoufei.Tab).TextMatrix(cing费用清单(ctabShoufei.Tab).RowSel, 2)
           lcount = cing费用清单(ctabShoufei.Tab).TextMatrix(cing费用清单(ctabShoufei.Tab).RowSel, 3)
           lCurrency = lsing * lcount
           sub修改收费项目数据 mstr收费编号, lstr收费项目编号, lsing, lcount, lCurrency
           
          
           '增加修改收费项目的详细信息,修改的收费项目名称，修改的金额 时间：2002/09/17　作者：徐冀川
           Dim lstr收费项目名称 As String
           Dim lstrIntro As String
           lstrIntro = ""
           lstr收费项目名称 = cing费用清单(ctabShoufei.Tab).TextMatrix(cing费用清单(ctabShoufei.Tab).RowSel, 1)
           lstrIntro = "收费项目名称：" & lstr收费项目名称 & "，单价：" & lsing & "元" & "，数量：" & lcount & "，金额：" & lCurrency & " 元"
           
           '功能：将修改处理发消息形式发送　时间：2002/08/05　作者：徐冀川
           sub消息发送 mstr收费编号, "收费编号为：" & mstr收费编号 & "的费用信息的单价或数量已被修改。" & lstrIntro
        End If
     End If
     
    '刷新界面
    For i = 1 To cing收费基本信息表.Rows - 1
        If cing收费基本信息表.Cell(flexcpText, i, 2) = mstr收费编号 Then
        cing收费基本信息表.TextMatrix(i, 5) = mcur总金额 * Val(cinb收费输入(打折比率).Text)
        End If
    Next
    
    Dim lbln是否有选中数 As Boolean
    lbln是否有选中数 = False
    
    For i = 1 To cing收费基本信息表.Rows - 1
        If cing收费基本信息表.Cell(flexcpChecked, i, 0) = 1 Then
            lbln是否有选中数 = True
            Exit For
        End If
    Next
    
    If lbln是否有选中数 = True Then
        sub界面数据刷新
    End If

errhandle:
    If cinb收费输入(mintCurInput).Enabled Then cinb收费输入(mintCurInput).SetFocus
    ctlb工具栏.Buttons("收费(&G)").Enabled = True
    If Err.Number = 0 Then Exit Sub
    sfsub错误处理 工程名, 模块名, "cing费用清单_AfterEdit", Err.Number, Err.Description
End Sub

'功能：在内部收费中增加费用信息的收费项目
'注意事项: 因为费用信息中，收费项目中的数量回定为1,单价与金额相同
'          所以在处理时，进行相应简化
'时间：2002/07/19
'作者：徐冀川
Private Function func增加收费项目(ByVal para收费编号 As String, ByVal Para收费项目编号 As String, ByVal Para单价 As Currency) As Boolean
On Error GoTo errHanler
    Dim lstrSql As String           '定义变量记录SQL语句
    Dim lstr收费批号 As String      '定义变量记录收费批号
    Dim lstr收费编号 As String      '定义变量记录收据编号
    Dim lstr交费人 As String        '定义变量记录交费人姓名
    Dim lstr交费单位名称 As String  '定义变量记录交费单位名称
    Dim lstr交费单位编号 As String  '定义变量记录交费单位编号
    Dim lstr交费日期 As String      '定义变量记录交费日期
    Dim lstr主管科经手人 As String  '定义变量记录主管科经手人
    Dim lstr主管科编号 As String    '定义变量记录主管科编号
    Dim lstr业务分类 As String      '定义变量记录业务分类
    Dim lobjTemp As Object          '定义临时变量记录集

    '初始化函数
    mblntemp = True
    func增加收费项目 = False

    lstrSql = "select * from 收费管理_费用信息表 where 收费编号='" & para收费编号 & "'"
    Set lobjTemp = dafuncGetData(lstrSql)
    
    '获取收费信息
    If lobjTemp.RecordCount > 0 Then
        lstr收费批号 = lobjTemp("收费批号")
        lstr收费编号 = lobjTemp("收费编号")
        lstr交费人 = IIf(IsNull(lobjTemp("交费人")), "不详姓名", lobjTemp("交费人"))
        lstr交费单位名称 = IIf(IsNull(lobjTemp("交费单位名称")), "不详单位", lobjTemp("交费单位名称"))
        lstr交费单位编号 = IIf(IsNull(lobjTemp("交费单位编号")), "不详单位编号", lobjTemp("交费单位编号"))
        lstr交费日期 = IIf(IsNull(lobjTemp("交费日期")), "交费日期不详", lobjTemp("交费日期"))
        lstr主管科经手人 = IIf(IsNull(lobjTemp("主管科室经手人")), "不详姓名", lobjTemp("主管科室经手人"))
        lstr主管科编号 = IIf(IsNull(lobjTemp("主管科室编号")), "不详编号", lobjTemp("主管科室编号"))
        lstr业务分类 = IIf(IsNull(lobjTemp("业务分类")), "不详编号", lobjTemp("业务分类"))
        
        '将已有的费用信息结合新增的收费项目信息，向数据库中插入一条收费项目数据
        lstrSql = "insert into 收费管理_费用信息表 (收费批号,收费编号,收费项目编号,数量,单价," & _
                  "金额,交费人,交费单位名称,交费单位编号,交费日期,主管科室经手人,主管科室编号,业务分类) " & _
                  " values ( '" & lstr收费批号 & "','" & lstr收费编号 & "','" & Para收费项目编号 & "'," & _
                   1 & " ," & Para单价 & "," & Para单价 & ",'" & lstr交费人 & "','" & _
                   lstr交费单位名称 & "','" & lstr交费单位编号 & "','" & lstr交费日期 & "','" & _
                   lstr主管科经手人 & "','" & lstr主管科编号 & "','" & lstr业务分类 & "')"
        dafuncGetData (lstrSql)
         
    Else
        Exit Function
    End If
    func增加收费项目 = True
Exit Function
errHanler:
    func增加收费项目 = False
    sfsub错误处理 工程名, 模块名, "sub增加收费项目", Err.Number, Err.Description
End Function


Private Sub sub修改收费项目数据(ByVal para收费编号 As String, ByVal Para项目编号 As String, ByVal Para单价 As Double, ByVal Para数量 As Long, ByVal Para金额 As Double)
On Error GoTo errhandler
    Dim lstrSql As String           '定义变量记录SQL语句
    
    lstrSql = "update  收费管理_费用信息表 set 数量= '" & Para数量 & "'," & _
              " 单价=convert(money,'" & Para单价 & "'), 金额=convert(money,'" & Para金额 & "')" & _
              " where 收费编号='" & para收费编号 & "' and 收费项目编号='" & Para项目编号 & "'"
    dafuncGetData (lstrSql)
Exit Sub
errhandler:
    sfsub错误处理 工程名, 模块名, "sub修改收费项目数据", Err.Number, Err.Description
End Sub

Private Sub cing费用清单_BeforeEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error Resume Next
    '功能：根据权限来操作对内部收费信息的修改.
    '作者：徐冀川
    '时间：2002/07/01
    If cchk内部收费.Value = vbChecked And ctabShoufei.Tab = 0 Then
        If umfunc校验用户权限("收费管理_内部收费信息修改") Then
            Cancel = False
        Else
            Cancel = True
            'sffuncMsg "没有修改内部收费的权限！"
            Exit Sub
        End If
    End If
    
    Select Case Col
        Case 费用清单_数量
            ctlb工具栏.Buttons("收费(&G)").Enabled = False
            mstrUndoCount = cing费用清单(ctabShoufei.Tab).TextMatrix(Row, Col)
            
        Case 费用清单_单价
            ctlb工具栏.Buttons("收费(&G)").Enabled = False
            mrds收费项目.MoveFirst
            'mrds收费项目.find "收费项目编号=" & cind字典(字典_收费项目).TextMatrix(cind字典(字典_收费项目).RowSel, 字典_收费项目编号)
            '修改:获取收费编号，从界面上的表格中获得. 时间：2002/02/28 徐冀川
            mrds收费项目.find "收费项目编号=" & cing费用清单(ctabShoufei.Tab).TextMatrix(Row, 0)
            If mrds收费项目.RecordCount > 0 Then
                mcur最小单价 = mrds收费项目("最小单价").Value
                mcur最大单价 = mrds收费项目("最大单价").Value
            Else
                sffuncMsg "未找到该收费项目的设置信息，该设置信息可能已被修改或删除，请退出收费界面，重新进入！"
            End If
            mstrUndoMoney = cing费用清单(ctabShoufei.Tab).TextMatrix(Row, Col)
        Case Else
            ctlb工具栏.Buttons("收费(&G)").Enabled = True
            Cancel = True
    End Select
End Sub



Private Sub cing费用清单_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next
' 功能：在内部收费不允许删除费用信息，时间：2002/02/6 徐冀川
If cchk内部收费.Value = vbChecked And ctabShoufei.Tab = 0 And KeyCode = vbKeyDelete Then
    If umfunc校验用户权限("收费管理_内部收费信息修改") Then
        Select Case KeyCode
        Case vbKeyDelete
            mobj界面通用对象_BeforeOperate "删除", False
        Case vbKeyEscape
            cind字典(字典_收费项目).Visible = False
            cind字典(字典_收费标准).Visible = False
            Me.KeyPreview = True
        End Select
    Else
        sffuncMsg "没有修改内部收费的权限！"
    End If
Else
    Select Case KeyCode
        Case vbKeyDelete
            mobj界面通用对象_BeforeOperate "删除", False
        Case vbKeyEscape
            cind字典(字典_收费项目).Visible = False
            cind字典(字典_收费标准).Visible = False
            Me.KeyPreview = True
    End Select
End If
End Sub

Private Sub cing费用清单_LostFocus(Index As Integer)
On Error Resume Next
    ctlb工具栏.Buttons("收费(&G)").Enabled = True
End Sub

Private Sub cing收费基本信息表_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error Resume Next
    If cing收费基本信息表.Cell(flexcpChecked, cing收费基本信息表.RowSel, 0) = 2 Then
        mcur基本信息总金额 = mcur基本信息总金额 - Val(cing收费基本信息表.TextMatrix(cing收费基本信息表.RowSel, 基本信息_金额))
    Else
        mcur基本信息总金额 = mcur基本信息总金额 + Val(cing收费基本信息表.TextMatrix(cing收费基本信息表.RowSel, 基本信息_金额))
    End If
    cinb收费输入(应收金额).Text = mcur基本信息总金额 * (cinb收费输入(打折比率).Text)
End Sub

Private Sub cing收费基本信息表_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error Resume Next
    If Col > 0 Then Cancel = True
End Sub

'功能：刷新界面的应收金额
'作者：徐冀川
'时间：2002/07/22
Private Sub sub界面数据刷新()
On Error GoTo errhander
    Dim i As Integer
    Dim lcurMoney As Currency
    lcurMoney = 0
    For i = 1 To cing收费基本信息表.Rows - 1
        If cing收费基本信息表.Cell(flexcpChecked, i, 0) = 1 Then
            lcurMoney = lcurMoney + Val(cing收费基本信息表.TextMatrix(i, 5))
        End If
    Next
    
    mcur基本信息总金额 = lcurMoney
    
    cinb收费输入(应收金额) = lcurMoney * Val(cinb收费输入(打折比率).Text)
    cinb收费输入(应收金额大写) = FuncConvertToCapsStr(Val(cinb收费输入(应收金额)))
Exit Sub
errhander:
    sfsub错误处理 "收费界面", "frm收费", "sub界面数据刷新", Err.Number, Err.Description
End Sub



Private Sub cing收费基本信息表_Click()
    Dim lrds明细费用 As Recordset
    Dim lrds打折比率 As Recordset
    Dim i As Long
    On Error GoTo errhandle
    If cing收费基本信息表.RowSel < 1 Then Exit Sub
    cing费用清单(收费).Rows = 1
    Set lrds明细费用 = mobj收费管理.funcExecute("select a.收费批号,a.收费编号,a.收费项目编号,b.收费项目名称,a.单价,a.数量,b.计量单位,a.金额,a.交费单位编号,a.交费单位名称,a.交费人,d.片区,主管科室名称 = c.名称" & _
                                                " from  收费管理_费用信息表 a left join 收费管理_收费项目字典表 b on a.收费项目编号 = b.收费项目编号" & _
                                                " left join 系统管理_科室字典表 c on a.主管科室编号=c.编号" & _
                                                " left join 单位档案_单位基本信息表 d on a.交费单位编号=d.申请编号" & _
                                                " where a.收费状态= 0 and a.收费编号='" & cing收费基本信息表.TextMatrix(cing收费基本信息表.RowSel, 基本信息_收费编号) & "'", "cls费用信息")
                                                
                                                
    '记录当前的收费项编号
    mstr收费编号 = cing收费基本信息表.TextMatrix(cing收费基本信息表.RowSel, 基本信息_收费编号)
    
    If lrds明细费用 Is Nothing Then Exit Sub
    If lrds明细费用.BOF And lrds明细费用.EOF Then Exit Sub
    lrds明细费用.MoveFirst
    For i = 0 To lrds明细费用.RecordCount - 1
        cing费用清单(收费).AddItem lrds明细费用("收费项目编号") & vbTab & _
                                   lrds明细费用("收费项目名称") & vbTab & _
                                   lrds明细费用("单价") & vbTab & _
                                   lrds明细费用("数量") & vbTab & _
                                   lrds明细费用("金额")
        If Not lrds明细费用.EOF Then lrds明细费用.MoveNext
    Next
    lrds明细费用.MoveFirst
    mstr交费单位编号 = lrds明细费用("交费单位编号").Value
    '徐冀川（修改2001/12/20）：增加在界面上对收费编号的显示
    cinb收费输入(收费_收费编号).Text = lrds明细费用("收费编号")
    cinb收费输入(收费_主管科室).Text = lrds明细费用("主管科室名称")
    
    
    '现示片区信息
    If IIf(IsNull(lrds明细费用("片区").Value), "", lrds明细费用("片区").Value) = "" Then
        Clab片区.Caption = "片区：(不详)"
    Else
        Clab片区.Caption = "(" + lrds明细费用("片区").Value + ")"
    End If
    
    cinb收费输入(收费_交费人).Text = lrds明细费用("交费人")
    cinb收费输入(收费_交费单位).Text = IIf(IsNull(lrds明细费用("交费单位名称").Value), "", lrds明细费用("交费单位名称").Value)
    Set lrds打折比率 = mobj收费管理.funcExecute("select 打折比率 from 收费管理_打折信息表 where 单位编号='" & mstr交费单位编号 & "'", "cls费用信息")
    If lrds打折比率.BOF And lrds打折比率.EOF Then
        cinb收费输入(打折比率).Text = "1.00"
    Else
        cinb收费输入(打折比率).Text = Format(lrds打折比率("打折比率").Value, "0.00")
    End If
errhandle:
    If Err.Number = 0 Then Exit Sub
    sfsub错误处理 "收费界面", "frm收费", "cing收费基本信息表_Click", Err.Number, Err.Description
End Sub



Private Sub ctabShoufei_Click(PreviousTab As Integer)
On Error GoTo errhandle
    cind字典(字典_收费标准).Visible = False
    cind字典(字典_收费项目).Visible = False
    cind字典(字典_主管科室).Visible = False
    

    If PreviousTab = 1 Then
        Ccbo收费项目大类.ListIndex = 0
        Ccbo收费项目大类.Refresh
    Else
        Ccbo门收费大类.ListIndex = 0
        Ccbo门收费大类.Refresh
    End If
    
    If ctabShoufei.Tab = 收费 And cchk内部收费.Value = 1 Then
        ctlb工具栏.Buttons("查询(&Q)").Enabled = True
        
        '修改:在从门诊收费界回来后,控制收费项目的控件应设为不可用 徐冀川 2002/11/26
        Ccbo收费项目大类.Enabled = False
        cinb收费输入(5).Enabled = False
    Else
        ctlb工具栏.Buttons("查询(&Q)").Enabled = False
    End If
    sub清除界面
Exit Sub
errhandle:
    sfsub错误处理 "收费界面", "ctabShoufei_Click", "cing收费基本信息表_Click", Err.Number, Err.Description
End Sub

Private Sub Ctim_Timer()
    Dim i As Long
    Dim j As Long
    Dim lobjRec As Object           '定义变量记录结果集
    Dim lstrSql As String              '定义变量记录SQL语句
    
    On Error GoTo errhandler
    
    '向cind字典(字典_收费项目)
    Ctim.Enabled = False
    Me.MousePointer = 11
    Me.cstuShoufei.Panels(1) = "正在加载基础数据，请稍后..."
    Me.Enabled = False
    
    '功能:向收费项目大类中,加入数据,在收费项目中根据收费项目大类来加载.
    '时间:2002/06/21
    '作者:徐冀川
    lstrSql = "select 收费项目编号,收费项目名称 from 收费管理_收费项目字典表 where len(收费项目编号)=3 " & _
           " order by 收费项目编号 "
    
    Set lobjRec = dafuncGetData(lstrSql)
        
    Do While Not lobjRec.EOF
        Ccbo收费项目大类.AddItem lobjRec("收费项目名称")
        Ccbo门收费大类.AddItem lobjRec("收费项目名称")
        lobjRec.MoveNext
    Loop
    
    Ccbo收费项目大类.ListIndex = 0
    Ccbo收费项目大类.Refresh
        
errhandler:
    Me.Enabled = True
    cind字典(字典_收费项目).Redraw = True
    subEnable收费输入
    Me.cstuShoufei.Panels(1) = ""
    Me.MousePointer = 0
    Exit Sub
    
End Sub

'功能:向收费项目网格中加载数据.
'输入:Para收费项目大类
'作者徐冀川:
'时间:2002/06/21

Private Sub sub收费项目填充(ByVal Para收费项目大类 As String)
On Error GoTo errhandler
    Dim lstrSql As String            '定义变量记录SQL语句
    Dim lobjRec As Object            '定义变量记录数据集
    Dim lstrtemp As String           '定义变量记录收费编号前缀码
    Dim i As Integer                 '定义循环变量
    Dim j As Integer                 '定义循环变量
    Dim lInt As Integer              '定义记录集个数
    Dim lobjRecCount As Object       '定义记录个数对象
    Dim lInt行数 As Integer
    Dim lbln标识 As Boolean
    
    '收费大类名称为空串,退出该过程
    If Para收费项目大类 = "" Then
        Exit Sub
    End If
    
    '根据收费项目大类名称,获取收费编号前缀
    lstrSql = "select 收费项目编号 from 收费管理_收费项目字典表 where 收费项目名称= '" & Para收费项目大类 & "'"
    Set lobjRec = dafuncGetData(lstrSql)
    lstrtemp = Left$(lobjRec("收费项目编号"), 3)
    
    '获取记录集个数
    lstrSql = "select count(*) as 记录集数 from 收费管理_收费项目字典表 where left(收费项目编号,3)='" & lstrtemp & "'"
    Set lobjRecCount = dafuncGetData(lstrSql)
    lInt = lobjRecCount("记录集数")
    
    '向网格中加入收费项目
    cind字典(字典_收费项目).Redraw = True
    cind字典(字典_收费项目).Clear
    
    mrds收费项目.MoveFirst
        
    If Not (mrds收费项目 Is Nothing) Then
        cind字典(字典_收费项目).Cols = mrds收费项目.Fields.Count

        cind字典(字典_收费项目).Refresh
        For i = 0 To mrds收费项目.Fields.Count - 1
            cind字典(字典_收费项目).TextMatrix(0, i) = mrds收费项目(i).Name
        Next
        
        lInt行数 = 1
        lbln标识 = False
        If (Not mrds收费项目.BOF) And (Not mrds收费项目.EOF) Then
            cind字典(字典_收费项目).Rows = lInt
            mrds收费项目.MoveFirst
            For i = 0 To mrds收费项目.RecordCount - 1
                
                If Left$(mrds收费项目("收费项目编号"), 3) = lstrtemp Then
                    lbln标识 = True
                    For j = 0 To mrds收费项目.Fields.Count - 1
                        '（功能：处理数据库可能存在的空值,将其转化为""，时间：2002/01/27,徐冀川）
                        cind字典(字典_收费项目).TextMatrix(lInt行数, j) = IIf(IsNull(mrds收费项目(j)), "", mrds收费项目(j))
                        cind字典(字典_收费项目).AutoSize j
                    Next j
                    
                End If
                If Not mrds收费项目.EOF Then mrds收费项目.MoveNext
                If lbln标识 = True Then
                    lInt行数 = lInt行数 + 1
                End If
            Next i
        End If
        
    End If
    cind字典(字典_收费项目).Width = 165
    For i = 0 To cind字典(字典_收费项目).Cols - 1
        cind字典(字典_收费项目).Width = cind字典(字典_收费项目).Width + cind字典(字典_收费项目).ColWidth(i)
    Next
    
    Exit Sub
errhandler:
    Me.Enabled = True
    cind字典(字典_收费项目).Redraw = True
    subEnable收费输入
    Me.cstuShoufei.Panels(1) = ""
    Me.MousePointer = 0
    Exit Sub
End Sub

Private Sub cupd修改打折比率_DownClick()
On Error Resume Next
    If Val(cinb收费输入(打折比率).Text) > 0 Then
        cinb收费输入(打折比率).Text = Format(CStr(Val(cinb收费输入(打折比率).Text) - 0.01), "0.00")
    Else
        cinb收费输入(打折比率).Text = "0.00"
    End If
End Sub
Private Sub cupd修改打折比率_GotFocus()
On Error Resume Next
    cind字典(字典_收费标准).Visible = False
    cind字典(字典_收费项目).Visible = False
    cind字典(字典_主管科室).Visible = False
End Sub

Private Sub cupd修改打折比率_UpClick()
On Error GoTo errhandle
    If Val(cinb收费输入(打折比率).Text) < 1 Then
        cinb收费输入(打折比率).Text = Format(CStr(Val(cinb收费输入(打折比率).Text) + 0.01), "0.00")
    Else
        cinb收费输入(打折比率).Text = "1.00"
    End If
Exit Sub
errhandle:
    sfsub错误处理 "收费管理界面对象", "frm收费", "Form_UpClick", Err.Number, Err.Description
End Sub

Private Sub Form_Activate()
On Error GoTo errhandle
    If mbln使用 Then
        ctabShoufei.Tab = 收费
    End If
Exit Sub
errhandle:
    sfsub错误处理 "收费管理界面对象", "frm收费", "Form_Activate", Err.Number, Err.Description
End Sub

Private Sub Form_Load()
Dim lcol工具栏 As Collection
On Error GoTo errhandle
    If pblnInUse Then Exit Sub
    pblnInUse = True
    mbln使用 = True
        
    Set mobj界面通用对象 = New cls界面通用对象
    Set mobj界面通用对象.Form = Me
    Set mobj界面通用对象.c工具栏 = ctlb工具栏
    
    Set lcol工具栏 = New Collection
    
    lcol工具栏.Add "查询(&Q)105"
    lcol工具栏.Add "|"
    lcol工具栏.Add "收费(&G)123"
    lcol工具栏.Add "|"
    lcol工具栏.Add "删除"
    lcol工具栏.Add "清空"
    lcol工具栏.Add "报废(&T)122"
    lcol工具栏.Add "|"
    lcol工具栏.Add "退出"
    
    mobj界面通用对象.subInitialize lcol工具栏, ""
    Set lcol工具栏 = Nothing
    
    '功能:在界面初始化时,默认为不是内部收费,隐藏cing收费基本信息表格
    '修改人:徐冀川
    '修改时间:2002/06/21
    
    cing收费基本信息表.Visible = False
    Frame5.Left = 120
    Frame5.Width = 10845
    cing费用清单(0).Width = 10600
    
    Clab业务分类.Enabled = False
    Ccbo业务分类.Enabled = False
    Ccbo业务分类.BackColor = &H80000000
    
    '以下为何嘉新增
    If Not func获取初始化数据 Then
        mbln使用 = False
        ctlb工具栏.Buttons(1).Enabled = False
        ctlb工具栏.Buttons(3).Enabled = False
        ctlb工具栏.Buttons(5).Enabled = False
        ctlb工具栏.Buttons(6).Enabled = False
        ctlb工具栏.Buttons(7).Enabled = False
        ctabShoufei.Visible = False
        Frame6.Visible = False
        Frame7.Visible = False
        Exit Sub
    End If
    
    
    sub初始化窗体
    '以上为何嘉新增
    
    ccmd选择.Enabled = False
    Copt今天.Enabled = False
    Copt本月.Enabled = False
    Copt所有.Enabled = False
    Cchk按基本信息查询.Enabled = False
    ctlb工具栏.Buttons(1).Enabled = False
    ctlb工具栏.Buttons(7).Enabled = False
    mstr收费编号 = ""
    mblntemp = False
    '余下初始化放入定时器。
    Ctim.Enabled = True
    
    '修改：2002-10-17（杨春）记忆打印前预览。
    On Error Resume Next
    Set mobj记忆 = New cls用户操作记忆
    mobj记忆.用户编号 = um用户编号
    mobj记忆.业务名 = "收费管理"
    If mobj记忆.记忆项值("打印前预览") = "是" Then
        cchk预览.Value = 1
    End If
    
    Exit Sub
errhandle:
    sfsub错误处理 "收费管理界面对象", "frm收费", "Form_Load", Err.Number, Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    pblnInUse = False
    Set mrds收费项目 = Nothing
    Set mrds收费标准 = Nothing
    Set mobj界面通用对象 = Nothing
    Set mobj收费管理 = Nothing
    Set mobj业务设置 = Nothing
    Set mobj单位档案 = Nothing
    
    '修改：2002-10-17（杨春）保存操作记忆项。
    mobj记忆.sub覆盖记忆值 "打印前预览", IIf(cchk预览.Value = 1, "是", "否")
    
End Sub


'&  ---------------| sub清除界面 |----------------------
'&  用途：  清除界面
'&  作者：  Shadow
'&  完成日期：2001/4/1
Private Sub sub清除界面()
    Dim i As Integer
    '徐冀川（2001/12/21）：增界面控件状态标识
    Dim l状态标识 As Integer
    
    On Error GoTo errhandle
    mcur总金额 = 0
    mcur基本信息总金额 = 0
    l状态标识 = 0
    
    '徐冀川（2001/12/21）:为了正常消除界面，改变界面控件属性，并记录标识值
    If Frame4.Enabled = False Then
        Frame4.Enabled = True
        l状态标识 = 1
    End If
  
  
    '在保存后需要清空收费编号;徐冀川;2002/9/30
    mstr交费单位编号 = ""
  
    If ctabShoufei.Tab = 收费 Then
        For i = 收费_收费编号 To 收费_收费标准
            cinb收费输入(i).Text = ""
        Next
        'cchk内部收费.Value = 0
        cchk同批收费.Value = 0
        cing收费基本信息表.Rows = 1
        cing费用清单(收费).Rows = 1
        If cinb收费输入(收费_收费编号).Enabled Then
            cinb收费输入(收费_收费编号).SetFocus
        Else
            cinb收费输入(收费_交费人).SetFocus
        End If
    Else
        For i = 门诊收费_交费人 To 门诊收费_收费标准
            cinb收费输入(i).Text = ""
        Next
        
        With cdtp日期(入院)
            .Year = Year(Date)
            .Month = Month(Date)
            .Day = Day(Date)
        End With
        
        With cdtp日期(出院)
            .Year = Year(Date)
            .Month = Month(Date)
            .Day = Day(Date)
        End With
        cing费用清单(门诊收费).Rows = 1
        If cinb收费输入(门诊收费_交费人).Enabled Then cinb收费输入(门诊收费_交费人).SetFocus
    End If
    
    cinb收费输入(打折比率).Text = "1.00"
    
    For i = 应收金额 To 交费日期
        cinb收费输入(i).Text = ""
    Next
    
    cind字典(字典_收费项目).Visible = False
    cind字典(字典_收费标准).Visible = False
    cind字典(字典_主管科室).Visible = False

    cinb收费输入(交费日期).Text = Date
    cinb收费输入(打折比率).Text = "1.00"
    cinb收费输入(门诊收费_主管科室).Text = um用户所属科室
    cinb收费输入(收费_主管科室).Text = um用户所属科室
    
    '徐冀川（2001/12/21）:根据标识值恢复界面控件属性
    If l状态标识 = 1 Then
        Frame4.Enabled = False
    Else
        Frame4.Enabled = True
    End If
Exit Sub
errhandle:
    sfsub错误处理 工程名, 模块名, "sub清除收费界面", Err.Number, Err.Description
End Sub


'&  ---------------| func查询费用信息 |----------------------
'&  用途：  按条件查询记录
'&  返回：  Recordset
'&  作者：  Shadow
'&  完成日期：2001/4/1
Private Function func查询费用信息() As Recordset
    Dim lstr查询条件 As String
    Dim i As Integer
    
    On Error GoTo errhandle
        
    If ctabShoufei.Tab = 收费 Then
        For i = 收费_收费编号 To 收费_主管科室
            If cinb收费输入(i).Text <> "" Then
                If lblCaption(i).Caption = "交费单位" Then
                    lstr查询条件 = lstr查询条件 & "交费单位编号" & "='" & mstr交费单位编号 & "' and "
                ElseIf lblCaption(i).Caption = "主管科室" Then
                    lstr查询条件 = lstr查询条件 & "主管科室编号" & "='" & mstr主管科室编号 & "' and "
                Else
                    lstr查询条件 = lstr查询条件 & lblCaption(i).Caption & "='" & cinb收费输入(i).Text & "' and "
                End If
            End If
        Next
    Else
        For i = 门诊收费_交费人 To 门诊收费_主管科室
            If cinb收费输入(i).Text <> "" Then lstr查询条件 = lstr查询条件 & lblCaption(i).Caption & "='" & cinb收费输入(i).Text & "' and "
        Next
    End If
    
    If (lstr查询条件 = vbNullString) Or (lstr查询条件 = vbNullChar) Then
        sffuncMsg "请在输入框中输入查询条件！", sf警告
        Exit Function
    Else
        lstr查询条件 = lstr查询条件 & "收费状态=0"
        Set func查询费用信息 = mobj收费管理.funcExecute("select a.收费批号,a.收费编号,a.收费项目编号,收费项目名称=(select 收费项目名称 from 收费管理_收费项目字典表 where 收费项目编号=a.收费项目编号),a.数量,计量单位=(select 计量单位 from 收费管理_收费项目字典表 where 收费项目编号=a.收费项目编号),a.单价,a.金额,a.收费状态,a.交费方式,a.交费人,a.交费单位编号,交费单位=(select 单位名称 from 单位档案_单位基本信息表 where 申请编号=a.交费单位编号),a.交费日期,a.退费日期,收费人编号=a.收费人,收费人=(select 姓名 from 系统管理_员工基本信息表 where 编号=a.收费人),退费人编号=a.退费人,退费人=(select 姓名 from 系统管理_员工基本信息表 where 编号=a.退费人) ,主管科室经手人编号=a.主管科室经手人,主管科室经手人=(select 姓名 from 系统管理_员工基本信息表 where 编号=a.主管科室经手人),主管科室编号,主管科室=(select 名称 from 系统管理_科室字典表 where 编号=a.主管科室编号),打折比率  from 收费管理_费用信息表 a where " & lstr查询条件, "cls费用信息")
                                            
    End If
    Exit Function
    
errhandle:
    Set func查询费用信息 = Nothing
    sfsub错误处理 工程名, 模块名, "func查询费用信息", Err.Number, Err.Description
End Function

'&  ---------------| func收集数据 |----------------------
'&  用途：  从界面上收集数据并组合成集合,输出到业务对象 "mobj收费管理"
'&  返回：  Collection
'&  作者：  Shadow
'&  完成日期：2001/4/1
Private Function func收集数据() As Collection
    Dim i As Long
    Dim j As Integer
    Dim lcol记录 As Collection
    Dim lcol数据 As Collection
    Dim lintTab As Integer              'ctabShoufei的当前页
    Dim lrds主管科室 As Recordset
    On Error GoTo errhandle
    lintTab = ctabShoufei.Tab           '取得ctabShoufei的当前页
    '组织数据
    If cing费用清单(lintTab).Rows = 1 Then
        sffuncMsg "无可用的费用信息！", sf警告
        GoTo WayOut
    Else
        Set lcol数据 = New Collection
        For i = 1 To cing费用清单(lintTab).Rows - 1
            Set lcol记录 = New Collection
            '将一条费用记录装入集合 "lcol记录"
            For j = 0 To cing费用清单(lintTab).Cols - 1
                '添加网格中的字段(“收费项目编号”、“单价”、“数量”、“金额”)
                lcol记录.Add cing费用清单(lintTab).TextMatrix(i, j), cing费用清单(lintTab).TextMatrix(0, j)
            Next
            If lintTab = 收费 Then
                '添加收费其他字段
                For j = 收费_交费人 To 收费_主管科室
                    If lblCaption(j).Caption = "交费单位" Then
                        lcol记录.Add mstr交费单位编号, "交费单位编号"
                        lcol记录.Add cinb收费输入(收费_交费单位).Text, "交费单位名称"
                    ElseIf lblCaption(j).Caption = "主管科室" Then
                        Set lrds主管科室 = mobj收费管理.funcExecute("select 编号 from 系统管理_科室字典表 where 名称='" & cinb收费输入(收费_主管科室).Text & "'", "cls费用信息")
                        If lrds主管科室.BOF And lrds主管科室.EOF Then
                            sffuncMsg "输入的主管科室不在配置的科室范围内！", sf警告
                            If cinb收费输入(收费_主管科室).Enabled Then cinb收费输入(收费_主管科室).SetFocus
                            Set lrds主管科室 = Nothing
                            Set func收集数据 = Nothing
                            Exit Function
                        Else
                            mstr主管科室编号 = lrds主管科室("编号").Value
                        End If
                        
                        lcol记录.Add mstr主管科室编号, "主管科室编号"
                    Else
                        lcol记录.Add cinb收费输入(j).Text, lblCaption(j).Caption
                    End If
                Next
                lcol记录.Add um用户编号, "主管科室经手人"                     '经手人要修改
                
                '*****************以下为何嘉修改——0808*******************
                mrds交费方式.Filter = "名称='" & cmb交费方式.Text & "'"
                mint交费方式编号 = mrds交费方式("编号").Value
                '*****************以上为何嘉修改——0808*******************
                
            Else
                '添加门诊收费其他字段
                For j = 门诊收费_交费人 To 门诊收费_主管科室
                    If lblCaption(j).Caption = "交费单位" Then
                        lcol记录.Add mstr交费单位编号, "交费单位编号"
                        lcol记录.Add cinb收费输入(门诊收费_交费单位).Text, "交费单位名称"
                    ElseIf lblCaption(j).Caption = "主管科室" Then
                        Set lrds主管科室 = mobj收费管理.funcExecute("select 编号 from 系统管理_科室字典表 where 名称='" & cinb收费输入(门诊收费_主管科室).Text & "'", "cls费用信息")
                        If lrds主管科室.BOF And lrds主管科室.EOF Then
                            sffuncMsg "输入的主管科室不在配置的科室范围内！", sf警告
                            If cinb收费输入(收费_主管科室).Enabled Then cinb收费输入(收费_主管科室).SetFocus
                            Set lrds主管科室 = Nothing
                            Set func收集数据 = Nothing
                            Exit Function
                        Else
                            mstr主管科室编号 = lrds主管科室("编号").Value
                        End If
                        
                        lcol记录.Add mstr主管科室编号, "主管科室编号"
                    Else
                        lcol记录.Add cinb收费输入(j).Text, lblCaption(j).Caption
                    End If
                Next
                
                lcol记录.Add um用户编号, "主管科室经手人"
                lcol记录.Add 0, "收费状态"
                lcol记录.Add cdtp日期(入院).Value 'CDate(cdtp日期(入院).Year & "/" & cdtp日期(入院).Month & "/" & cdtp日期(入院).Day), "入院日期"
                lcol记录.Add cdtp日期(出院).Value 'CDate(cdtp日期(出院).Year & "/" & cdtp日期(出院).Month & "/" & cdtp日期(出院).Day), "出院日期"
            End If
            '将所有费用记录装入集合 "lcol数据"
            lcol数据.Add lcol记录
        Next
    End If
    Set func收集数据 = lcol数据
    GoTo WayOut
    
errhandle:
    Set func收集数据 = Nothing
    sfsub错误处理 工程名, 模块名, "func收集数据", Err.Number, Err.Description, True
WayOut:
    Set lcol记录 = Nothing
    Set lcol数据 = Nothing
    Set lrds主管科室 = Nothing
End Function

'修改：2002-6-25（杨春）生成收据号。
Private Function func收集确认信息() As Collection
    Dim lstr收据号  As String
    
    On Error Resume Next
    Set func收集确认信息 = New Collection
    With func收集确认信息
        .Add mstr收费批号, "收费批号"
        .Add Val(cinb收费输入(打折比率).Text), "打折比率"
        .Add mint交费方式编号, "收费方式"
        .Add CDate(cinb收费输入(交费日期).Text), "交费日期"
        .Add um用户编号, "收费人"
        .Add "", "退费人"
        .Add CDate("1900/1/1"), "退费日期"
    
        '修改：2002-6-25（杨春）生成收据号。
        lstr收据号 = mobj收费管理.func生成收据号
        .Add lstr收据号, "收据号"
    End With
    
End Function


'&  ---------------| func获取初始化数据 |----------------------
'&  用途：  获取初始化数据
'&  作者：  Shadow
'&  完成日期：2001/4/3
Private Function func获取初始化数据() As Boolean
    Dim lobj业务配置 As Object
    On Error GoTo errhandle
    '初始化业务及单位档案对象
    Set mobj收费管理 = CreateObject("收费业务对象.cls收费管理")
    Set mobj业务设置 = CreateObject("收费业务对象.cls业务设置")
    Set mobj单位档案 = CreateObject("单位档案业务.ClsUnitInterface")
    Set lobj业务配置 = CreateObject("收费数据对象.cls业务配置")
    mint打折控制 = lobj业务配置.打折控制
    mint科目级数 = lobj业务配置.科目级数
    Set mrds收费项目 = mobj收费管理.func查询收费项目("datalength(收费项目编号)=" & CStr(3 * mint科目级数))
    If mrds收费项目 Is Nothing Then
        sffuncMsg "获取收费项目失败，请与系统管理员联系！", sf警告
        func获取初始化数据 = False
        Exit Function
    End If
    If mrds收费项目.BOF Or mrds收费项目.EOF Then
        sffuncMsg """收费项目""尚未配置！请配置好""收费项目""再使用收费功能。", sf警告
        func获取初始化数据 = False
        Exit Function
    End If
    Set mrds收费标准 = mobj收费管理.funcExecute("select 收费标准名称,助记符 from 收费管理_收费标准信息表 group by 助记符,收费标准名称", "cls费用信息")
    Set mrds主管科室 = dafuncGetData("select * from 系统管理_科室字典表")
    If mrds主管科室 Is Nothing Then
        sffuncMsg "无法从系统表： ""系统管理_科室字典表""获取初始化数据。" & vbCrLf & "收费功能将不可使用！请与系统管理员联系！", sf警告
        func获取初始化数据 = False
        Exit Function
    End If
    If mrds主管科室.BOF Or mrds主管科室.EOF Then
        sffuncMsg """系统管理_科室字典表"" 尚未配置,收费功能将不可使用！" & vbCrLf & "请与系统管理员联系！"
        func获取初始化数据 = False
        Exit Function
    End If
    Set mrds交费方式 = mobj业务设置.func提取字典表信息("收费方式字典表")
    If mrds交费方式 Is Nothing Then
        sffuncMsg "获取收费方式失败,请与系统管理员联系！", sf警告
        func获取初始化数据 = False
        Exit Function
    End If
    If mrds交费方式.BOF And mrds交费方式.EOF Then
        sffuncMsg "交费方式尚未配置，请先配置好交费方式再使用收费功能！", sf警告
        func获取初始化数据 = False
        Exit Function
    End If
    func获取初始化数据 = True
    GoTo WayOut
errhandle:
    If Err.Number = 9999 Then
        sffuncMsg "尚缺:" & Mid$(Err.Description, InStr(Err.Description, ":") + 1) & vbCrLf & "请配置好后再使用收费功能。"
    End If
    func获取初始化数据 = False
    sfsub错误处理 工程名, 模块名, "func获取初始化数据", Err.Number, Err.Description
WayOut:
    
End Function


'&  ---------------| sub初始化窗体 |----------------------
'&  用途：  初始化收费界面
'&  返回：  无
'&  作者：  Shadow
'&  完成日期：2001/4/3
Private Sub sub初始化窗体()
On Error GoTo errhandle
    mstrUndoCount = ""
    mstrUndoMoney = ""
    mintCurInput = 0
    mlngX = 0
    mlngY = 0
    mstr交费单位编号 = ""
    mstr主管科室编号 = ""
    mint交费方式编号 = 0
    mstr收费批号 = ""
    mcur总金额 = 0
    mcur基本信息总金额 = 0
    Dim i As Long
    Dim j As Long
    '初始化 "cing收费基本信息表"
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '修改人：徐冀川
    '功能：初始化cing收费基本信息表
    '时间：2001-12-20
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    With cing收费基本信息表
        .Cols = 6
        .Rows = 1
        .TextMatrix(0, 基本信息_选择) = "选择"
        .ColWidth(基本信息_选择) = 480
        .ColAlignment(基本信息_选择) = flexAlignCenterCenter
       
        .TextMatrix(0, 基本信息_收费批号) = "收费批号"
        .ColWidth(基本信息_收费批号) = 1440
'        .ColAlignment(基本信息_收费批号) = flexAlignCenterCenter
         
       
        .TextMatrix(0, 基本信息_收费编号) = "收费编号"
        .ColWidth(基本信息_收费编号) = 1440
'        .ColAlignment(基本信息_收费编号) = flexAlignCenterCenter
        
        .TextMatrix(0, 基本信息_交费人) = "交费人"
        .ColWidth(基本信息_交费人) = 880
'        .ColAlignment(基本信息_交费人) = flexAlignCenterCenter
        
        .TextMatrix(0, 基本信息_交费单位名称) = "交费单位名称"
        .ColWidth(基本信息_交费单位名称) = 1800
'        .ColAlignment(基本信息_交费单位名称) = flexAlignCenterCenter
          
        .TextMatrix(0, 基本信息_金额) = "金额"
        .ColWidth(基本信息_金额) = 1000
'        .ColAlignment(基本信息_金额) = flexAlignCenterCenter
    End With
    
    '初始化 "cing费用清单"
    With cing费用清单(收费)
        .Cols = 5
        .Rows = 1
        .TextMatrix(0, 费用清单_收费项目编号) = "收费项目编号"
        .ColWidth(费用清单_收费项目编号) = 1310
        .ColAlignment(费用清单_收费项目编号) = flexAlignCenterCenter
        
        .TextMatrix(0, 费用清单_收费项目名称) = "收费项目名称"
        .ColWidth(费用清单_收费项目名称) = 1320
'        .ColAlignment(费用清单_收费项目名称) = flexAlignCenterCenter
        
        .TextMatrix(0, 费用清单_单价) = "单价"
        .ColWidth(费用清单_单价) = 480
'        .ColAlignment(费用清单_单价) = flexAlignCenterCenter
        
        .TextMatrix(0, 费用清单_数量) = "数量"
        .ColWidth(费用清单_数量) = 500
'        .ColAlignment(费用清单_数量) = flexAlignCenterCenter
        
        
        '.TextMatrix(0, 费用清单_计量单位) = "计量单位"
        '.ColWidth(费用清单_计量单位) = 900
        '.ColAlignment(费用清单_计量单位) = flexAlignCenterCenter
        
        .TextMatrix(0, 费用清单_金额) = "金额"
        .ColWidth(费用清单_金额) = 570
'        .ColAlignment(费用清单_金额) = flexAlignCenterCenter
    End With
    '隐藏收费批号列
    cing收费基本信息表.ColHidden(1) = True
    With cing费用清单(门诊收费)
        .Cols = 5
        .Rows = 1
        .TextMatrix(0, 费用清单_收费项目编号) = "收费项目编号"
        .ColWidth(费用清单_收费项目编号) = 1410
        .ColAlignment(费用清单_收费项目编号) = flexAlignCenterCenter
        
        .TextMatrix(0, 费用清单_收费项目名称) = "收费项目名称"
        .ColWidth(费用清单_收费项目名称) = 1410
'        .ColAlignment(费用清单_收费项目名称) = flexAlignCenterCenter
        
        .TextMatrix(0, 费用清单_单价) = "单价"
        .ColWidth(费用清单_单价) = 480
'        .ColAlignment(费用清单_单价) = flexAlignCenterCenter
        
        .TextMatrix(0, 费用清单_数量) = "数量"
        .ColWidth(费用清单_数量) = 480
'        .ColAlignment(费用清单_数量) = flexAlignCenterCenter
        
        
        '.TextMatrix(0, 费用清单_计量单位) = "计量单位"
        '.ColWidth(费用清单_计量单位) = 900
        '.ColAlignment(费用清单_计量单位) = flexAlignCenterCenter
        
        .TextMatrix(0, 费用清单_金额) = "金额"
        .ColWidth(费用清单_金额) = 570
'        .ColAlignment(费用清单_金额) = flexAlignCenterCenter
    End With
    
    '初始化 "cind字典(字典_收费项目)"
        
    With cind字典(字典_收费项目)
        .Rows = 1
        .Cols = 1
    End With
    
'    '向cind字典(字典_收费项目)
'    If Not (mrds收费项目 Is Nothing) Then
'        cind字典(字典_收费项目).Cols = mrds收费项目.Fields.Count
'        cind字典(字典_收费项目).Refresh
'        For i = 0 To mrds收费项目.Fields.Count - 1
'            cind字典(字典_收费项目).TextMatrix(0, i) = mrds收费项目(i).Name
'        Next
'
'        If (Not mrds收费项目.BOF) And (Not mrds收费项目.EOF) Then
'            cind字典(字典_收费项目).Rows = mrds收费项目.RecordCount + 1
'            mrds收费项目.MoveFirst
'            For i = 0 To mrds收费项目.RecordCount - 1
'                For j = 0 To mrds收费项目.Fields.Count - 1
'                    '（功能：处理数据库可能存在的空值,将其转化为""，时间：2002/01/27,徐冀川）
'                    cind字典(字典_收费项目).TextMatrix(i + 1, j) = IIf(IsNull(mrds收费项目(j)), "", mrds收费项目(j))
'                    cind字典(字典_收费项目).AutoSize j
'                Next j
'                If Not mrds收费项目.EOF Then mrds收费项目.MoveNext
'            Next i
'        End If
'    End If
'    cind字典(字典_收费项目).Width = 165
'    For i = 0 To cind字典(字典_收费项目).Cols - 1
'        cind字典(字典_收费项目).Width = cind字典(字典_收费项目).Width + cind字典(字典_收费项目).ColWidth(i)
'    Next
    
    Dim llngStyle As Long
    
    '初始化 "cind字典(字典_收费标准)"
    If Not (mrds收费标准 Is Nothing) Then
        cind字典(字典_收费标准).Cols = mrds收费标准.Fields.Count + 1
        For i = 1 To mrds收费标准.Fields.Count
            cind字典(字典_收费标准).TextMatrix(0, i) = mrds收费标准(i - 1).Name
        Next
        cind字典(字典_收费标准).TextMatrix(0, 0) = "编号"
        
        If Not (mrds收费标准.BOF And mrds收费标准.EOF) Then
            cind字典(字典_收费标准).Rows = mrds收费标准.RecordCount + 1
            mrds收费标准.MoveFirst
            For i = 1 To mrds收费标准.RecordCount
                cind字典(字典_收费标准).TextMatrix(i, 0) = i
                For j = 1 To mrds收费标准.Fields.Count
                    '（功能：处理数据库可能存在的空值,将其转化为""，时间：2002/01/27,徐冀川）
                    cind字典(字典_收费标准).TextMatrix(i, j) = IIf(IsNull(mrds收费标准(j - 1)), "", mrds收费标准(j - 1))
                    cind字典(字典_收费标准).AutoSize j
                Next j
                If Not mrds收费标准.EOF Then mrds收费标准.MoveNext
            Next i
        End If
        
        cind字典(字典_收费标准).Width = 165
        For i = 0 To cind字典(字典_收费标准).Cols - 1
            cind字典(字典_收费标准).Width = cind字典(字典_收费标准).Width + cind字典(字典_收费标准).ColWidth(i)
        Next
    End If
    
    '初始化 "cind字典_主管科室"
    If Not (mrds主管科室 Is Nothing) Then
        cind字典(字典_主管科室).Cols = mrds主管科室.Fields.Count
        For i = 0 To mrds主管科室.Fields.Count - 1
            cind字典(字典_主管科室).TextMatrix(0, i) = mrds主管科室(i).Name
        Next

        If (Not mrds主管科室.BOF) And (Not mrds主管科室.EOF) Then
            cind字典(字典_主管科室).Rows = mrds主管科室.RecordCount + 1
            mrds主管科室.MoveFirst
            For i = 0 To mrds主管科室.RecordCount - 1
                For j = 0 To mrds主管科室.Fields.Count - 1
                    '（功能：处理数据库可能存在的空值,将其转化为""，时间：2002/01/27,徐冀川）
                    cind字典(字典_主管科室).TextMatrix(i + 1, j) = IIf(IsNull(mrds主管科室(j)), "", mrds主管科室(j))
                    cind字典(字典_主管科室).AutoSize j
                Next j
                If Not mrds主管科室.EOF Then mrds主管科室.MoveNext
            Next i
        End If
    End If
    
    llngStyle = GetWindowLong(cind字典(字典_收费项目).hwnd, GWL_STYLE)
    SetWindowLong cind字典(字典_收费项目).hwnd, GWL_STYLE, llngStyle Or WS_THICKFRAME
    SetWindowLong cind字典(字典_收费标准).hwnd, GWL_STYLE, llngStyle Or WS_THICKFRAME
    SetWindowLong cind字典(字典_主管科室).hwnd, GWL_STYLE, llngStyle Or WS_THICKFRAME
    
        '修改(徐冀川):
    '功能:根权限控制界面上打折控件的属性
    '时间:2001-12-19
    If umfunc校验用户权限("收费管理_打折") Then
        Frame6.Enabled = True
        Frame6.Caption = "打折情况"
        lblCaption(19).Enabled = True
        cinb收费输入(19).Enabled = True
        cupd修改打折比率.Enabled = True
        cchk打印打折比率.Enabled = True
        Select Case mint打折控制
            Case 0
                cupd修改打折比率.Enabled = False
                cinb收费输入(19).Enabled = False
            Case 1
                cupd修改打折比率.Enabled = True
                cinb收费输入(19).Enabled = True
            Case 2
                cupd修改打折比率.Enabled = False
                cinb收费输入(19).Enabled = False
            Case Else
        End Select
        
    Else
        Frame6.Caption = "打折情况(无权限)"
        Frame6.Enabled = False
        lblCaption(19).Enabled = False
        cinb收费输入(19).Enabled = False
        cupd修改打折比率.Enabled = False
        cchk打印打折比率.Enabled = False
    End If
        
    '初始化交费方式列表
    If Not (mrds交费方式 Is Nothing) Then
        If Not (mrds交费方式.BOF And mrds交费方式.EOF) Then
            mrds交费方式.MoveFirst
            For i = 0 To mrds交费方式.RecordCount - 1
                cmb交费方式.AddItem mrds交费方式("名称")
                If Not mrds交费方式.EOF Then mrds交费方式.MoveNext
            Next
        End If
        cmb交费方式.ListIndex = 0
    End If
    cinb收费输入(交费日期).Text = Date
    cinb收费输入(收费_主管科室).Text = um用户所属科室
    cinb收费输入(门诊收费_主管科室).Text = um用户所属科室
    
    Exit Sub
errhandle:
    sfsub错误处理 "收费管理界面对象", "frm收费", "sub初始化窗体", Err.Number, Err.Description
End Sub

'&  ---------------| subDisable收费输入 |----------------------
'&  用途：  使收费界面上的输入框和表格失效
'&  作者：  Shadow
'&  完成日期：2001/4/3
Private Sub subDisable收费输入()
On Error Resume Next
    Dim i As Long
    For i = 收费_收费编号 To 收费_收费标准
        cinb收费输入(i).Enabled = False
    Next
    cing收费基本信息表.Enabled = False
    cing费用清单(收费).Enabled = False
End Sub

'&  ---------------| subEnable收费输入 |----------------------
'&  用途：  使收费界面上的输入框和表格有效
'&  返回：  无
'&  作者：  Shadow
'&  完成日期：2001/4/3

Private Sub subEnable收费输入()
    Dim i As Long
    
    On Error Resume Next
    If cchk内部收费.Value = 1 Then
        For i = 收费_收费编号 To 收费_收费标准
            cinb收费输入(i).Enabled = True
        Next
        For i = 收费_收费项目 To 收费_收费标准
            cinb收费输入(i).Enabled = False
        Next
    Else
        For i = 收费_收费项目 To 收费_收费标准
            cinb收费输入(i).Enabled = True
        Next
        cinb收费输入(收费_收费编号).Enabled = False
    End If
    
    If cchk同批收费.Value = 1 Then cing收费基本信息表.Editable = True
    cing费用清单(收费).Enabled = True
    Err.Clear
End Sub

'&  ---------------| subDisable门诊收费输入 |----------------------
'&  用途：  使门诊收费界面上的输入框和表格失效
'&  作者：  Shadow
'&  完成日期：2001/4/3
Private Sub subDisable门诊收费输入()
On Error Resume Next
    Dim i As Long
    For i = 门诊收费_交费人 To 门诊收费_收费标准
        cinb收费输入(i).Enabled = False
    Next
    For i = 入院 To 出院
        cdtp日期(i).Enabled = False
    Next
    cing费用清单(门诊收费).Enabled = False
End Sub

'&  ---------------| subEnable门诊收费输入 |----------------------
'&  用途：  使门诊收费界面上的输入框和表格有效
'&  作者：  Shadow
'&  完成日期：2001/4/3
Private Sub subEnable门诊收费输入()
On Error Resume Next
    Dim i As Long
    For i = 门诊收费_交费人 To 门诊收费_收费标准
        cinb收费输入(i).Enabled = True
    Next
    For i = 入院 To 出院
        cdtp日期(i).Enabled = True
    Next
    cing费用清单(门诊收费).Enabled = True
End Sub

Private Sub mobj界面通用对象_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    Dim lrds费用信息 As Recordset
    Dim lcol数据 As Collection
    Dim lcol收费确认信息 As Collection
    Dim lstr划价返回 As String
    Dim lstr收费编号 As String
    Dim i As Long, j As Long
    Dim lstr收费编号组() As String
    Dim lint收费编号数量 As Integer
    
    On Error GoTo errhandle
    Select Case Operate

'&  =============================================| 收费 |============================================
        Case "收费"
            If Not ValidateData Then Exit Sub
            
            ctlb工具栏.Buttons(9).Enabled = False
            
            mrds交费方式.Filter = "名称='" & cmb交费方式.Text & "'"
            mint交费方式编号 = mrds交费方式("编号").Value
            
            If cing费用清单(ctabShoufei.Tab).Rows = 1 Then
                sffuncMsg "无可用费用信息", sf警告
                GoTo WayOut
            End If
            
            If IsNumeric(cinb收费输入(19).Text) Then
                If CDbl(cinb收费输入(19).Text) = 0 Then
                    MsgBox "打折比率不能为0！", vbOKOnly & vbExclamation, "系统提示"
                    Cancel = True
                    cinb收费输入(19).Text = "1.00"
                    Exit Sub
                End If
            Else
                MsgBox "打折比率录入不正确！", vbOKOnly & vbExclamation, "系统提示"
                Cancel = True
                Exit Sub
            End If
            
            If cchk内部收费.Value = 0 Or ctabShoufei.Tab = 门诊收费 Then
                
                '直接收费
                Set lcol数据 = func收集数据
                If lcol数据 Is Nothing Then
                    sffuncMsg "收集数据失败！", sf警告
                    GoTo WayOut
                End If
                lstr划价返回 = mobj收费管理.func划价_数据集合(lcol数据)
                
                '修改：加上对病人出入院日期的判断，时间：2002/02/05，徐冀川
                If cdtp日期(0).Value > cdtp日期(1).Value Then
                    sffuncMsg "出院日期不能小于入院日期,请重新输入时间！"
                    GoTo WayOut
                End If
                
                If lstr划价返回 = "" Then
                    sffuncMsg "执行划价操作时失败！", sf警告
                    GoTo WayOut
                Else
                    lstr收费编号 = Mid(lstr划价返回, InStr(lstr划价返回, ";") + 1)
                    mstr收费批号 = lstr收费编号
                End If
                
                '保存收费信息。
                '修改：2002-6-25（因为收据确认信息中要产生收据号，所有添加事务处理）。
                On Error GoTo errTransHandler
                dasubBeginTran
                
                Set lcol收费确认信息 = func收集确认信息 '修改：2002-6--25（杨春）会产生收据号。
                
                Call mobj收费管理.func收费(mstr收费批号, lcol收费确认信息)
                
                dasubCommitTran
                On Error GoTo errhandle
                
                ctlb工具栏.Buttons("收费(&G)").Enabled = False
                
                MsgBox "收费完成！请等待打印票据！", vbInformation, "系统提示"
                mcur总金额 = 0
                cinb收费输入(应收金额) = "0"
                If Not (cchk内部收费.Value = 1 And ctabShoufei.Tab = 收费) Then sub清除界面
                If ctabShoufei.Tab = 收费 Then
                    If cinb收费输入(收费_交费人).Enabled Then
                        cinb收费输入(收费_交费人).SetFocus
                    Else
                        cinb收费输入(收费_收费编号).SetFocus
                    End If
                Else
                    If cinb收费输入(门诊收费_交费人).Enabled Then cinb收费输入(门诊收费_交费人).SetFocus
                End If
            
            ElseIf (cchk内部收费.Value = 1) And (ctabShoufei.Tab = 收费) Then
                
                '内部收费
                For i = 1 To cing收费基本信息表.Rows - 1
                    If cing收费基本信息表.Cell(flexcpChecked, i, 0) = 1 Then
                        lint收费编号数量 = lint收费编号数量 + 1
                        ReDim Preserve lstr收费编号组(lint收费编号数量)
                        lstr收费编号组(lint收费编号数量) = cing收费基本信息表.TextMatrix(i, 2)
                    End If
                Next
                
                If lint收费编号数量 = 0 Then
                    sffuncMsg "无选中的收费信息！", sf警告
                    GoTo WayOut
                End If
                                                
                mstr收费批号 = lstr收费编号组(1)
                
                '保存收费信息。
                '修改：2002-6-25（因为收据确认信息中要产生收据号，所有添加事务处理）。
                On Error GoTo errTransHandler
                dasubBeginTran
                
                Set lcol收费确认信息 = func收集确认信息 '修改：2002-6--25（杨春）会产生收据号。
                
                For i = 1 To UBound(lstr收费编号组)
                    Call mobj收费管理.func收费(lstr收费编号组(i), lcol收费确认信息)
                Next
                
                dasubCommitTran
                On Error GoTo errhandle
                
                'ReDim lstr收费编号组(0)
                For i = cing收费基本信息表.Rows - 1 To 1 Step -1
                    If i <= cing收费基本信息表.Rows - 1 Then
                        If cing收费基本信息表.Cell(flexcpChecked, i, 0) = 1 Then
                            cing收费基本信息表.RemoveItem i
                        End If
                    End If
                Next
                cing费用清单(ctabShoufei.Tab).Rows = 1
                ctlb工具栏.Buttons("收费(&G)").Enabled = False
                MsgBox "收费完成！请等待打印票据！", vbInformation, "系统提示"
                mcur基本信息总金额 = 0
                cinb收费输入(应收金额) = "0"
                cinb收费输入(实收金额).Text = "0"
                If Not (cchk内部收费.Value = 1 And ctabShoufei.Tab = 收费) Then sub清除界面
                If cinb收费输入(收费_收费编号).Enabled Then cinb收费输入(收费_收费编号).SetFocus
            End If
'&  =====================================================| 打印 |=================================================
                Dim lrdsreturn As Recordset
                Dim llngFieldCounter As Long
                Dim llngRecordCounter As Long
                Dim lstr格式文件名 As String
                Dim lcol分类记录 As Collection
                Dim lcol分类费用 As Collection
                Dim lrds返回格式文件名 As Recordset
                Dim lobj汇总记录 As Object
                
                '取得收费记录中有的票据类型数量
                Set lrdsreturn = mobj收费管理.funcExecute("select b.票据类型编号 from 收费管理_收费项目字典表 b, 收费管理_费用信息表 c " & _
                                                                "Where b.收费项目编号 = c.收费项目编号 and c.收费批号 ='" & _
                                                                mstr收费批号 & "' group by b.票据类型编号", "cls费用信息")
                
                If lrdsreturn Is Nothing Then
                    sffuncMsg "未检索到收费项目的票据类型信息,无法进行打印！", sf警告
                    If Not (cchk内部收费.Value = 1 And ctabShoufei.Tab = 收费) Then sub清除界面
                    GoTo WayOut
                End If

                If (lrdsreturn.BOF And lrdsreturn.EOF) Then
                    sffuncMsg "未检索到收费项目的票据类型信息,无法进行打印！", sf警告
                    If Not (cchk内部收费.Value = 1 And ctabShoufei.Tab = 收费) Then sub清除界面
                Else
                    lrdsreturn.MoveFirst
                End If
                
                '按票据类型取出费用信息
                For i = 0 To lrdsreturn.RecordCount - 1
                                                            
                    Set lrds费用信息 = mobj收费管理.funcExecute("select * from 收费管理_打印费用信息 where 票据类型编号=" & lrdsreturn("票据类型编号") & " and 收费批号='" & mstr收费批号 & "'", "cls费用信息")

                    Set lcol分类费用 = New Collection
                    If lrds费用信息 Is Nothing Then
                        sffuncMsg "无可打印信息！", sf警告
                        If Not (cchk内部收费.Value = 1 And ctabShoufei.Tab = 收费) Then sub清除界面
                        GoTo WayOut
                    End If
                    
                    If lrds费用信息.BOF And lrds费用信息.EOF Then
                        sffuncMsg "无可打印信息！", sf警告
                        If Not (cchk内部收费.Value = 1 And ctabShoufei.Tab = 收费) Then sub清除界面
                        GoTo WayOut
                    End If
                    
                     '*******以下为何嘉新增-0823**********
                    Dim lstr交费单位 As String
                    Dim lstr交费人 As String
                    If IIf(IsNull(lrds费用信息("交费单位名称").Value), "", lrds费用信息("交费单位名称")) <> "" Then
                        lstr交费单位 = lrds费用信息("交费单位名称").Value
                    Else
                        lstr交费单位 = ""
                    End If
                    If IIf(IsNull(lrds费用信息("交费人").Value), "", lrds费用信息("交费人")) <> "" Then
                        lstr交费人 = lrds费用信息("交费人").Value
                    Else
                        lstr交费人 = ""
                    End If
                    '********以上为何嘉新增-0823**********
                    
                    
                    '修改：2002-9-29（杨春）合并打印。
                    Set lobj汇总记录 = mobj收费管理.funcExecute("select 收费项目编号,单价=avg(单价),数量=sum(数量),金额=sum(金额) from 收费管理_打印费用信息 " _
                                & "where 票据类型编号=" & lrdsreturn("票据类型编号") & " and 收费批号='" & mstr收费批号 _
                                & "' group by 收费批号,收费项目编号", "cls费用信息")
                    
                    For llngRecordCounter = 0 To lobj汇总记录.RecordCount - 1
                        
                        '修改：2002-9-29（杨春）获取当前项目的详细信息。
                        Set lrds费用信息 = mobj收费管理.funcExecute("select * from 收费管理_打印费用信息 where 票据类型编号=" & lrdsreturn("票据类型编号") & " and 收费批号='" & mstr收费批号 & "' AND 收费项目编号='" & lobj汇总记录("收费项目编号") & "'", "cls费用信息")
                        
                        Set lcol分类记录 = New Collection
                        
                        '将一条记录添加到一个集合(lcol分类记录)
                        For llngFieldCounter = 0 To lrds费用信息.Fields.Count - 1
                            If lrds费用信息.Fields(llngFieldCounter).Name = "交费单位名称" Or lrds费用信息.Fields(llngFieldCounter).Name = "交费人" Then
                                If lrds费用信息.Fields(llngFieldCounter).Name = "交费单位名称" Then lcol分类记录.Add lstr交费单位, "交费单位名称"
                                If lrds费用信息.Fields(llngFieldCounter).Name = "交费人" Then lcol分类记录.Add lstr交费人, "交费人"
                            ElseIf lrds费用信息.Fields(llngFieldCounter).Name <> "单价" And lrds费用信息.Fields(llngFieldCounter).Name <> "数量" And lrds费用信息.Fields(llngFieldCounter).Name <> "金额" Then
                                '修改：2002-9-29（杨春）单价、数量、金额显示汇总数据。
                                lcol分类记录.Add lrds费用信息(llngFieldCounter).Value, lrds费用信息.Fields(llngFieldCounter).Name
                            End If
                        Next
                        '修改：2002-9-29（杨春）单价、数量、金额显示汇总数据。
                        lcol分类记录.Add Format(lobj汇总记录("单价").Value, "0.00"), "单价"
                        lcol分类记录.Add lobj汇总记录("数量").Value, "数量"
                        lcol分类记录.Add Format(lobj汇总记录("金额").Value, "0.00"), "金额"
                        
'                        '添加门诊收费信息
                        lcol分类记录.Add cinb收费输入(门诊收费_年龄).Text, "年龄"
                        lcol分类记录.Add cinb收费输入(门诊收费_性别).Text, "性别"
                        lcol分类记录.Add cinb收费输入(门诊收费_住院号).Text, "住院号"
                        lcol分类记录.Add cinb收费输入(门诊收费_病种).Text, "病种"
                        lcol分类记录.Add CStr(cdtp日期(入院).Year) & CStr(cdtp日期(入院).Month) & CStr(cdtp日期(入院).Day), "入院日期"
                        lcol分类记录.Add CStr(cdtp日期(出院).Year) & CStr(cdtp日期(出院).Month) & CStr(cdtp日期(出院).Day), "出院日期"
                        lcol分类记录.Add cinb收费输入(门诊收费_入院操作员).Text, "入院操作员"
                        lcol分类记录.Add cinb收费输入(门诊收费_经治医生).Text, "经治医生"
                        
                        '将费用记录加入集合lcol分类费用
                        lcol分类费用.Add lcol分类记录
                        
                        '将记录向下移
'                        If Not lrds费用信息.EOF Then lrds费用信息.MoveNext
                        If Not lobj汇总记录.EOF Then lobj汇总记录.MoveNext
                        
                    Next
                    
                    '查找票据格式。
                    If ctabShoufei.Tab = 收费 Then
                        Set lrds返回格式文件名 = mobj收费管理.funcExecute("select * from 收费管理_票据设置信息表 where 票据类型编号='" & lrdsreturn("票据类型编号") & "' and 对应业务='一般'", "cls费用信息")
                    Else
                        Set lrds返回格式文件名 = mobj收费管理.funcExecute("select * from 收费管理_票据设置信息表 where 票据类型编号='" & lrdsreturn("票据类型编号") & "' and 对应业务='门诊'", "cls费用信息")
                    End If
                    If lrds返回格式文件名 Is Nothing Then
                        sffuncMsg "未查找到票据格式文件！", sf警告
                    End If
                        
                    '开始打印票据。
                    If lrds返回格式文件名.BOF And lrds返回格式文件名.EOF Then
                        sffuncMsg "未查找到票据格式文件！", sf警告
                    Else
                        lstr格式文件名 = lrds返回格式文件名("票据格式文件名称")
                        Call mobj收费管理.sub打印票据(lcol分类费用, App.Path & "\" & lstr格式文件名, IIf(cchk预览.Value = 1, True, False), cchk打印打折比率.Value, lrds返回格式文件名("最大项数").Value)
                    End If
                    If Not lrdsreturn.EOF Then lrdsreturn.MoveNext
                Next
                
                If Not (cchk内部收费.Value = 1 And ctabShoufei.Tab = 收费) Then sub清除界面
                ctlb工具栏.Buttons("收费(&G)").Enabled = True
                
                ' 恢复退出功能(徐冀川2002_1_10)
                ctlb工具栏.Buttons(9).Enabled = True
                Set lcol分类记录 = Nothing
                Set lcol分类费用 = Nothing
                           
'&  ======================================| 查询 |============================================
        Case "查询"

            '修改：2001-11-22（允许按交费单位、主管科室）查询。
            Dim lobjRec As Object
            Dim lstr待查收费编号 As String
           ' Dim lstr交费人 As String
            Dim lstr单位名称 As String
            Dim lstr主管科室编号 As String
            Dim lstrSql As String
            Dim lstr时间参数 As String
            Dim lstr今天 As String
            Dim lstr本月 As String
            Dim lstr所有 As String
            Dim lTime As String
            Dim lstr业务分类 As String          '定义变量记录业务分类
            
            '更新收项目大类
            Sub更新收费项目大类
            
            '清空内部收费网格中的内容
            cing收费基本信息表.Rows = 1
            '查询所需记录：收费编号、主管科室、交费单位编号。
            lstr待查收费编号 = cinb收费输入(收费_收费编号).Text
            If cinb收费输入(收费_交费单位) = "" Then
                lstr单位名称 = ""
            Else
                lstr单位名称 = Trim(cinb收费输入(收费_交费单位))
            End If
            
            If cinb收费输入(收费_交费人) = "" Then
                lstr交费人 = ""
            Else
                lstr交费人 = Trim(cinb收费输入(收费_交费人))
            End If
            
            If cinb收费输入(收费_主管科室) = "" Then
                lstr主管科室编号 = ""
            Else
                lstr主管科室编号 = mstr主管科室编号
            End If
            If lstr待查收费编号 = "" Then
                cing收费基本信息表.Rows = 1
            End If
            
            lstr业务分类 = Ccbo业务分类.Text
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '修改人：徐冀川
            '功能：按条件生成查语句
            '时间：2001/12/20
            '
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            If Cchk按基本信息查询.Value = Unchecked Then
                If Copt今天.Value = True Then lstr时间参数 = "今天"
                If Copt本月.Value = True Then lstr时间参数 = "本月"
                If Copt所有.Value = True Then lstr时间参数 = "所有"
                
                '增加对业务分类的查询
                
                If lstr业务分类 = "所有业务" Then
                    lstrSql = "select a.打折比率,a.收费批号,a.收费编号,a.收费项目编号,b.收费项目名称,a.单价,a.数量,b.计量单位,a.金额,a.交费单位名称,a.交费人,d.片区,c.名称 as 主管科室名称" & _
                        " from 收费管理_费用信息表 a left join 收费管理_收费项目字典表 b on a.收费项目编号 = b.收费项目编号 " & _
                        " left join 系统管理_科室字典表 c on a.主管科室编号=c.编号 " & _
                        " left join 单位档案_单位基本信息表 d  on a.交费单位编号=d.申请编号" & _
                        " where a.收费状态= 0 "
                Else
                    lstrSql = "select a.打折比率,a.收费批号,a.收费编号,a.收费项目编号,b.收费项目名称,a.单价,a.数量,b.计量单位,a.金额,a.交费单位名称,a.交费人,d.片区, c.名称 as 主管科室名称" & _
                        " from 收费管理_费用信息表 a left join 收费管理_收费项目字典表 b on a.收费项目编号 = b.收费项目编号 " & _
                        " left join 系统管理_科室字典表 c on a.主管科室编号=c.编号 " & _
                        " left join 单位档案_单位基本信息表 d  on a.交费单位编号=d.申请编号" & _
                        " where a.收费状态= 0  and a.业务分类= '" & lstr业务分类 & "'"
                End If

                Select Case lstr时间参数
                    Case "今天" '显示当天的收费记录！
                        lTime = Format(Now(), "yyyy-mm-dd")
                        lstrSql = lstrSql & " and left(convert(varchar(30),a.交费日期,120),10)='" & lTime & "'"
                    Case "本月" '显示本月的收费记录！
                        lTime = Format(Now(), "yyyy-mm")
                        lstrSql = lstrSql & " and left(convert(varchar(30),a.交费日期,120),7)='" & lTime & "'"
                    Case "所有" '显示所有的收费记录！
                        lstrSql = lstrSql
                End Select
                
            
            Else
            
                If lstr业务分类 = "所有业务" Then
                    lstrSql = "select a.打折比率,a.收费批号,a.收费编号,a.收费项目编号,b.收费项目名称,a.单价,a.数量,b.计量单位,a.金额,a.交费单位名称,a.交费人,d.片区,主管科室名称 = c.名称" & _
                        " from 收费管理_费用信息表 a left join 收费管理_收费项目字典表 b on a.收费项目编号 = b.收费项目编号 " & _
                        " left join 系统管理_科室字典表 c on a.主管科室编号=c.编号 " & _
                        " left join 单位档案_单位基本信息表 d on a.交费单位编号=d.申请编号" & _
                        " where a.收费状态= 0 and 收费编号=" & IIf(lstr待查收费编号 = "", "收费编号", "'" & lstr待查收费编号 & "'") & _
                        " and a.主管科室编号=" & IIf(lstr主管科室编号 = "" Or lstr待查收费编号 <> "", "a.主管科室编号", "'" & lstr主管科室编号 & "'") & _
                        " and a.交费单位名称=" & IIf(lstr单位名称 = "" Or lstr待查收费编号 <> "", "a.交费单位名称", "'" & lstr单位名称 & "'") & _
                        " and a.交费人= " & IIf(lstr交费人 = "" Or lstr待查收费编号 <> "", "a.交费人", "'" & lstr交费人 & "'")
                Else
                    lstrSql = "select a.打折比率,a.收费批号,a.收费编号,a.收费项目编号,b.收费项目名称,a.单价,a.数量,b.计量单位,a.金额,a.交费单位名称,a.交费人,d.片区,主管科室名称 = c.名称" & _
                        " from 收费管理_费用信息表 a left join 收费管理_收费项目字典表 b on a.收费项目编号 = b.收费项目编号 " & _
                        " left join 系统管理_科室字典表 c on a.主管科室编号=c.编号 " & _
                        " left join 单位档案_单位基本信息表 d on a.交费单位编号=d.申请编号" & _
                        " where a.业务分类= '" & lstr业务分类 & "'" & _
                        " and a.收费状态= 0 and 收费编号=" & IIf(lstr待查收费编号 = "", "收费编号", "'" & lstr待查收费编号 & "'") & _
                        " and a.主管科室编号=" & IIf(lstr主管科室编号 = "" Or lstr待查收费编号 <> "", "a.主管科室编号", "'" & lstr主管科室编号 & "'") & _
                        " and a.交费单位名称=" & IIf(lstr单位名称 = "" Or lstr待查收费编号 <> "", "a.交费单位名称", "'" & lstr单位名称 & "'") & _
                        " and a.交费人= " & IIf(lstr交费人 = "" Or lstr待查收费编号 <> "", "a.交费人", "'" & lstr交费人 & "'")
                End If
            End If
            
            '设置排序规则
            lstrSql = lstrSql + " order by  a.收费编号 desc"
            Set lrds费用信息 = mobj收费管理.funcExecute(lstrSql, "cls费用信息")
                        
            If (lrds费用信息 Is Nothing) Then
                'cing费用清单(0).Clear
                
                Clab收费项目大类.Enabled = False
                Ccbo收费项目大类.Enabled = False
                lblCaption(5).Enabled = False
                cinb收费输入(5).Enabled = False
            
                cing费用清单(0).Rows = 1
                sffuncMsg "无符合条件的费用信息！", sf警告
                GoTo WayOut
            ElseIf (lrds费用信息.BOF And lrds费用信息.EOF) Then
                'cing费用清单(0).Clear
                Clab收费项目大类.Enabled = False
                Ccbo收费项目大类.Enabled = False
                lblCaption(5).Enabled = False
                cinb收费输入(5).Enabled = False
                cing费用清单(0).Rows = 1
                sffuncMsg "无符合条件的费用信息！", sf警告
                GoTo WayOut
            Else
                lrds费用信息.MoveFirst
                If cchk内部收费.Value = 0 Then
                    cing收费基本信息表.Rows = 1
                End If
                cing费用清单(收费).Rows = 1
                If lrds费用信息("交费单位名称").Value <> vbNullString Then
                    cinb收费输入(收费_交费单位) = lrds费用信息("交费单位名称")
                End If
                
                cinb收费输入(收费_交费人) = lrds费用信息("交费人")
                cinb收费输入(收费_主管科室) = lrds费用信息("主管科室名称")
                cinb收费输入(打折比率).Text = lrds费用信息("打折比率")
                
                
                '现示片区信息
                If IIf(IsNull(lrds费用信息("片区")), "", lrds费用信息("片区")) = "" Then
                    Clab片区.Caption = "片区：(不详)"
                Else
                    Clab片区.Caption = "(" + lrds费用信息("片区") + ")"
                End If
                
                '修改：2001/12/20（徐冀川）界面增加收费编号的显示
                cinb收费输入(收费_收费编号).Text = lrds费用信息("收费编号")
                cinb收费输入(收费_主管科室).Text = lrds费用信息("主管科室名称")
                
                '向表格中添加项目
                For i = 0 To lrds费用信息.RecordCount - 1
                    'lrds费用信息("计量单位") & vbTab &
                    '如收费编号相同,向表格"cing费用清单"添加项目,并累计"cing收费基本信息表"中的金额
                    If lrds费用信息("收费编号") = cing收费基本信息表.TextMatrix(cing收费基本信息表.Rows - 1, 基本信息_收费编号) Then
                        cing费用清单(收费).AddItem lrds费用信息("收费项目编号") & vbTab & _
                                                   lrds费用信息("收费项目名称") & vbTab & _
                                                   lrds费用信息("单价") & vbTab & _
                                                   lrds费用信息("数量") & vbTab & _
                                                   lrds费用信息("金额")
                       '累计金额
                        cing收费基本信息表.TextMatrix(cing收费基本信息表.Rows - 1, 基本信息_金额) = _
                        CStr(Val(cing收费基本信息表.TextMatrix(cing收费基本信息表.Rows - 1, 基本信息_金额)) + lrds费用信息("金额"))
                        mcur总金额 = cing收费基本信息表.TextMatrix(cing收费基本信息表.Rows - 1, 基本信息_金额)
                    Else
                    '如收费编号不相同,向表格"cing收费基本信息表"添加项目,并清空"cing费用信息(收费)"中的项目,再重新添加。
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    '徐冀川（2001/12/20）修改
                    '功能：在表格中显示内部收费的信息
                    'lrds费用信息("收费批号")
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        cing收费基本信息表.AddItem vbTab & lrds费用信息("收费批号") & vbTab & lrds费用信息("收费编号") & vbTab & lrds费用信息("交费人") & _
                                                   vbTab & lrds费用信息("交费单位名称") & vbTab & lrds费用信息("金额")
                        cing收费基本信息表.Cell(flexcpChecked, cing收费基本信息表.Rows - 1, 0) = 2
                        cing费用清单(收费).Rows = 1
                        cing费用清单(收费).AddItem lrds费用信息("收费项目编号") & vbTab & _
                                                   lrds费用信息("收费项目名称") & vbTab & _
                                                   lrds费用信息("单价") & vbTab & _
                                                   lrds费用信息("数量") & vbTab & _
                                                   lrds费用信息("金额")
                        '隐藏显示收费批号列
                        cing收费基本信息表.ColHidden(1) = True
                                               
                    End If
                    If Not lrds费用信息.EOF Then lrds费用信息.MoveNext
                Next
                
                '记录当前的收费项编号
                mstr收费编号 = cing收费基本信息表.TextMatrix(cing收费基本信息表.RowSel, 基本信息_收费编号)
                If cing收费基本信息表.Rows > 1 Then
                    mcur基本信息总金额 = 0
                    For i = 1 To cing收费基本信息表.Rows - 1
                        If cing收费基本信息表.Cell(flexcpChecked, i, 0) = 1 Then
                            mcur基本信息总金额 = (mcur基本信息总金额 + cing收费基本信息表.TextMatrix(i, 基本信息_金额))
                        End If
                    Next
                    
                    '功能：界面元素控制,徐冀川,2002/07/22
                    Clab收费项目大类.Enabled = True
                    Ccbo收费项目大类.Enabled = True
                    lblCaption(5).Enabled = True
                    cinb收费输入(5).Enabled = True
                    
                End If
                cinb收费输入(应收金额).Text = mcur基本信息总金额 * Val(cinb收费输入(打折比率).Text)
            End If
                        
            If cinb收费输入(实收金额).Enabled Then cinb收费输入(实收金额).SetFocus
            
        Case "删除"
            
            If Not cing费用清单(ctabShoufei.Tab).Enabled Then Exit Sub
            If cing费用清单(ctabShoufei.Tab).RowSel < 1 Then Exit Sub
            ' 功能：控费用信息的删除既：内部收费才不能删除，门诊情况可以删除。时间：2002/02/06 徐冀川
            If cchk内部收费.Value = vbChecked And ctabShoufei.Tab = 0 Then
            
                If Not umfunc校验用户权限("收费管理_内部收费信息修改") Then
                    sffuncMsg "没有修改内部收费的权限！"
                    GoTo WayOut
                End If
                
                If cing费用清单(ctabShoufei.Tab).Rows - 1 = 1 Then
                     sffuncMsg "费用信息的最后一项收费项目不允许删除，要想删除它只有报废它的费用信息！"
                     Exit Sub
                Else
                    '获取总金额
                    Dim lcurMoney As Currency
                    For i = 1 To cing费用清单(ctabShoufei.Tab).Rows - 1
                        lcurMoney = lcurMoney + Val(cing费用清单(ctabShoufei.Tab).TextMatrix(i, 费用清单_金额))
                    Next
                    mcur总金额 = lcurMoney
                    mcur总金额 = mcur总金额 - Val(cing费用清单(ctabShoufei.Tab).TextMatrix(cing费用清单(ctabShoufei.Tab).RowSel, 费用清单_金额))
                    
                    Dim lstrtemp As String
                    lstrtemp = cing费用清单(ctabShoufei.Tab).TextMatrix(cing费用清单(ctabShoufei.Tab).RowSel, 0)
                    cing费用清单(ctabShoufei.Tab).RemoveItem cing费用清单(ctabShoufei.Tab).RowSel
                   
                    If mstr收费编号 = "" And lstrtemp = "" Then
                        sffuncMsg "费用信息错误无法删除！"
                        Exit Sub
                    Else
                        Dim lstr收费项目编号 As String
                        lstr收费项目编号 = lstrtemp
                        sub删除费用信息收费项目 mstr收费编号, lstr收费项目编号
                        
                        '修改：增加了对删除收费项详细信息的说明 列出收费项目名称
                        Dim lstr收费项目名称 As String
                        
                        lstr收费项目名称 = cing费用清单(ctabShoufei.Tab).TextMatrix(cing费用清单(ctabShoufei.Tab).RowSel, 1)
                        
                        '功能：将修改处理发消息形式发送 时间：2002/08/05 作者：徐冀川
                        sub消息发送 mstr收费编号, "收费编号为：" & mstr收费编号 & "的费用信息的收费项目：" & lstr收费项目名称 & " 已被删除！"
                    End If
                     
                    '更新界面显示
        
                    For i = 1 To cing收费基本信息表.Rows - 1
                        If cing收费基本信息表.Cell(flexcpText, i, 2) = mstr收费编号 Then
                            cing收费基本信息表.TextMatrix(i, 5) = mcur总金额 * Val(cinb收费输入(打折比率).Text)
                        End If
                    Next
                    
                    Dim lbln是否有选中数 As Boolean
                    lbln是否有选中数 = False
                    
                    For i = 1 To cing收费基本信息表.Rows - 1
                        If cing收费基本信息表.Cell(flexcpChecked, i, 0) = 1 Then
                            lbln是否有选中数 = True
                            Exit For
                        End If
                    Next
                    
                    If lbln是否有选中数 = True Then
                        sub界面数据刷新
                    End If
                    
                    'cinb收费输入(应收金额).Text = mcur总金额 * Val(cinb收费输入(打折比率).Text)
                End If
                
            Else
                mcur总金额 = mcur总金额 - Val(cing费用清单(ctabShoufei.Tab).TextMatrix(cing费用清单(ctabShoufei.Tab).RowSel, 费用清单_金额))
                cing费用清单(ctabShoufei.Tab).RemoveItem cing费用清单(ctabShoufei.Tab).RowSel
                cinb收费输入(应收金额).Text = mcur总金额 * Val(cinb收费输入(打折比率).Text)
            End If
        Case "清空"
            sub清除界面
            
        Case "报废"
            '功能:增加对内部收费信息的报废处理.
            '时间:2002/07/01
            '作者:徐冀川
            For i = 1 To cing收费基本信息表.Rows - 1
                If cing收费基本信息表.Cell(flexcpChecked, i, 0) = 1 Then
                    lint收费编号数量 = lint收费编号数量 + 1
                    ReDim Preserve lstr收费编号组(lint收费编号数量)
                    lstr收费编号组(lint收费编号数量) = cing收费基本信息表.TextMatrix(i, 2)
                End If
            Next
                
            If lint收费编号数量 = 0 Then
                sffuncMsg "无选中的收费信息！", sf警告
                GoTo WayOut
            End If
            
            If MsgBox("你确信要报废选项中的费用信息吗？", vbQuestion + vbYesNo + vbDefaultButton2, "系统提示") = vbYes Then
                
                
                '循环报废费用信息
                For i = 1 To UBound(lstr收费编号组)
                
                    '功能：将修改处理发消息形式发送 徐冀川 2002/08/05
                    sub消息发送 lstr收费编号组(i), "收费编号为：" & lstr收费编号组(i) & "的费用信息已被报废！"
                    sub报废费用信息 lstr收费编号组(i)
                Next

                mobj界面通用对象_BeforeOperate "查询", False
                
            End If
 
        Case "退出"
            Set lrds费用信息 = Nothing
        Case Else
    End Select
    GoTo WayOut
    
errhandle:
    If Err.Number = 94 Then Resume Next
    If Err.Number = 5 Then GoTo WayOut
    If Err.Number = 20513 Then
        MsgBox "打印机出错！", vbExclamation, "系统提示"
        GoTo WayOut
    End If
    
    sfsub错误处理 工程名, 模块名, "mobj界面通用对象_BeforeOperate", Err.Number, Err.Description
    
WayOut:
    ctlb工具栏.Buttons("收费(&G)").Enabled = True
    ctlb工具栏.Buttons(9).Enabled = True
    Set lcol数据 = Nothing
    Set lrds费用信息 = Nothing
    If cinb收费输入(收费_收费编号).Enabled And Operate <> "退出" Then cinb收费输入(收费_收费编号).SetFocus
    Exit Sub
'    Resume
    
errTransHandler:
    dasubRollBack
    GoTo errhandle
End Sub

'根据收费编号删除费用信息的收费项目.
'时间:2002/07/02
'作者:徐冀川
Private Sub sub删除费用信息收费项目(ByVal para收费编号 As String, ByVal Para项目编号 As String)
On Error GoTo errhandler
    Dim lstrSql As String     '定义变量记录SQL语句
    
    lstrSql = "delete from 收费管理_费用信息表 where 收费编号='" & para收费编号 & "'" & _
              " and 收费项目编号='" & Para项目编号 & "'"
    dafuncGetData (lstrSql)
Exit Sub
errhandler:
     sfsub错误处理 "收费界面", "frm收费", "sub删除费用信息收费项目", Err.Number, Err.Description
End Sub


'根据收费编号删除费用信息.
'时间:2002/0607/01
'作者:徐冀川
Public Sub sub报废费用信息(ByVal para收费编号)
On Error GoTo errhandler
    Dim lstrSql As String           '定义变量记录SQL语句
    
    lstrSql = "delete from 收费管理_费用信息表 where 收费编号='" & para收费编号 & "'"
    dafuncGetData (lstrSql)
Exit Sub
errhandler:
    sfsub错误处理 "收费界面", "frm收费", "sub报废费用信息", Err.Number, Err.Description
End Sub


'功能: 将给定金额转换为人民币的大写字符串
'输入: money       金额
'输出: FuncConvertToCapsStr     转换的大写字符串
'最后修改时间: 96.6.11
'--------------------------------------------------
Public Function FuncConvertToCapsStr(Money As Currency) As String
On Error GoTo errhandle
    Const digit_str = "零壹贰叁肆伍陆柒捌玖"
    Const unit_str = "仟佰拾万仟佰拾元角分"
    Dim money_str As String
    
    If Money > 99999999.99 Then
        FuncConvertToCapsStr = ""
    ElseIf Money = 0 Then
        FuncConvertToCapsStr = "零元整"
    Else
        Dim temp_str As String
        Dim i, j As Integer
        
        If Money < 0 Then
            money_str = "负"
            Money = -Money
        Else
            money_str = ""
        End If
        
        temp_str = Format(Money, "00000000.00")
        
        '转换整数部分
        For i = 1 To 8
            If Mid(temp_str, i, 1) <> "0" Then Exit For
        Next
        For i = i To 8
            j = CInt(Mid(temp_str, i, 1))
            If j > 0 Then
                money_str = money_str & Mid(digit_str, j + 1, 1) & Mid(unit_str, i, 1)
            Else
                If i = 4 Then
                    money_str = money_str & "万"
                ElseIf i = 8 Then
                    money_str = money_str & "元"
                ElseIf Mid(temp_str, i + 1, 1) <> "0" Then
                    money_str = money_str & Mid(digit_str, j + 1, 1)
                End If
            End If
        Next
        
        '转换小数部分
        If Right(temp_str, 2) = "00" Then
            money_str = money_str & "整"
        Else
            '转换角
            j = CInt(Mid(temp_str, 10, 1))
            If j > 0 Then
                money_str = money_str & Mid(digit_str, j + 1, 1) & "角"
            Else
                money_str = money_str & "零"
            End If
            '转换分
            j = CInt(Mid(temp_str, 11, 1))
            If j > 0 Then
                money_str = money_str & Mid(digit_str, j + 1, 1) & "分"
            Else
                money_str = money_str & "整"
            End If
        End If
        
        FuncConvertToCapsStr = money_str
    End If
Exit Function
errhandle:
    sfsub错误处理 "收费界面", "frm收费", " FuncConvertToCapsStr()", Err.Number, Err.Description
End Function

Private Function func检查项目是否已选(ByVal Para收费项目编号 As String) As Boolean
On Error GoTo errhandle
    Dim i As Long
    func检查项目是否已选 = False
    If cing费用清单(ctabShoufei.Tab).Rows = 1 Then
        func检查项目是否已选 = False
        Exit Function
    End If
    
    For i = 1 To cing费用清单(ctabShoufei.Tab).Rows - 1
        If Para收费项目编号 = cing费用清单(ctabShoufei.Tab).TextMatrix(i, 费用清单_收费项目编号) Then
            func检查项目是否已选 = True
            Exit Function
        End If
    Next
Exit Function
errhandle:
    sfsub错误处理 "收费界面", "frm收费", " func检查项目是否已选()", Err.Number, Err.Description
End Function

Private Function ValidateData() As Boolean
On Error GoTo errhandle
    Select Case ctabShoufei.Tab
        Case 收费
            
            If cchk内部收费.Value = 0 Then
                If cinb收费输入(收费_交费人).Text = vbNullString And cinb收费输入(收费_交费单位) = vbNullString Then
                    ValidateData = False
                    sffuncMsg """交费人"" 和 ""交费单位"" 必须输入其中之一！", sf警告
                Else
                    ValidateData = True
                End If
                If cinb收费输入(收费_主管科室).Text = "" Then
                    ValidateData = False
                    sffuncMsg "请输入主管科室！", sf警告
                End If
            Else
                ValidateData = True
            End If
        Case 门诊收费
            If cinb收费输入(门诊收费_交费人).Text = vbNullString And cinb收费输入(门诊收费_交费单位) = vbNullString Then
                ValidateData = False
                sffuncMsg """交费人"" 和 ""交费单位"" 必须输入其中之一！", sf警告
            Else
                ValidateData = True
            End If
            If cinb收费输入(门诊收费_主管科室).Text = "" Then
                ValidateData = False
                sffuncMsg "请输入主管科室！", sf警告
            End If
    End Select
Exit Function
errhandle:
    sfsub错误处理 "收费界面", "frm收费", " ValidateData()", Err.Number, Err.Description
End Function


Private Function func匹配收费项目(ByVal paraValue As String) As Long
    Dim i As Long
    
    On Error GoTo errhandle
    
    func匹配收费项目 = 0
    
    If paraValue = vbNullString Then Exit Function
    '如 paraValue 首位为数字,则匹配编号
    If Asc(paraValue) >= Asc("0") And Asc(paraValue) <= Asc("9") Then
        For i = 1 To cind字典(字典_收费项目).Rows - 1
            If Left(cind字典(字典_收费项目).TextMatrix(i, 字典_收费项目编号), Len(paraValue)) = paraValue Then
                cind字典(字典_收费项目).TopRow = i
                cind字典(字典_收费项目).Select i, 0
                func匹配收费项目 = i
                Exit Function
            End If
        Next
    End If
    
    '如 paraValue 首位为字母,则匹配助记符
    If Asc(paraValue) >= Asc("A") And Asc(paraValue) <= Asc("z") Then
        For i = 1 To cind字典(字典_收费项目).Rows - 1
            If Left(cind字典(字典_收费项目).TextMatrix(i, 字典_助记符), Len(paraValue)) = paraValue Then
                cind字典(字典_收费项目).TopRow = i
                cind字典(字典_收费项目).Select i, 0
                func匹配收费项目 = i
                Exit Function
            End If
        Next
    End If
    
    '其它情况匹配名称

    For i = 1 To cind字典(字典_收费项目).Rows - 1
        If Left(cind字典(字典_收费项目).TextMatrix(i, 字典_收费项目名称), Len(paraValue)) = paraValue Then
            cind字典(字典_收费项目).TopRow = i
            cind字典(字典_收费项目).Select i, 0
            func匹配收费项目 = i
            Exit Function
        End If
    Next
    
errhandle:
    If Err.Number = 0 Then Exit Function
    func匹配收费项目 = 0
    sfsub错误处理 "收费界面", "frm收费", "func匹配收费项目", Err.Number, Err.Description
End Function

Private Function func匹配收费标准(ByVal paraValue As String) As Long
On Error GoTo errhandle
    Dim i As Long
    func匹配收费标准 = 0
    If paraValue = vbNullString Then Exit Function
    '如 paraValue 首位为数字,则匹配编号
    If Asc(paraValue) >= "0" And Asc(paraValue) <= Asc("9") Then
        For i = 1 To cind字典(字典_收费标准).Rows - 1
            If Left(cind字典(字典_收费标准).TextMatrix(i, 字典_收费标准编号), Len(paraValue)) = paraValue Then
                cind字典(字典_收费标准).TopRow = i
                cind字典(字典_收费标准).Select i, 0
                func匹配收费标准 = i
                Exit Function
            End If
        Next
    End If
    
    '如 paraValue 首位为字母,则匹配助记符
    If Asc(paraValue) >= Asc("A") And Asc(paraValue) <= Asc("z") Then
        For i = 1 To cind字典(字典_收费标准).Rows - 1
            If Left(cind字典(字典_收费标准).TextMatrix(i, 字典_助记符), Len(paraValue)) = paraValue Then
                cind字典(字典_收费标准).TopRow = i
                cind字典(字典_收费标准).Select i, 0
                func匹配收费标准 = i
                Exit Function
            End If
        Next
    End If
    '其它匹配名称
    For i = 1 To cind字典(字典_收费标准).Rows - 1
        If Left(cind字典(字典_收费标准).TextMatrix(i, 字典_收费标准名称), Len(paraValue)) = paraValue Then
            cind字典(字典_收费标准).TopRow = i
            cind字典(字典_收费标准).Select i, 0
            func匹配收费标准 = i
            Exit Function
        End If
    Next
Exit Function
errhandle:
    sfsub错误处理 "收费界面", "frm收费", "func匹配收费标准", Err.Number, Err.Description
End Function

Private Sub Disable()
On Error Resume Next
    Dim i As Control
    For Each i In Controls
        i.Enabled = False
    Next
    ctlb工具栏.Buttons("退出(ESC)").Enabled = True
End Sub


'功能：通过信使客户端，向指定的科室发送消息
'输入：Para目的科室编号,Para消息内容
'时间：2002/08/05
'作者：徐冀川

Private Sub sub消息发送(ByVal para收费编号 As String, ByVal Para消息内容 As String)
On Error GoTo errhander
    Dim lstr目的科室号 As String        '定义变量记录目的科室号
    Dim lstrSql As String               '定义变量记录SqL语句
    Dim lobjRec As Object               '定义对象记录临数据集
    Dim lstr交费人 As String            '定义变量记录交费人
    Dim lstr交费单位 As String          '定义变量记录交费单位
    
    
    '如果用户科室编号或是没有消息内容就退出过程
    If para收费编号 = "" Or Para消息内容 = "" Then
        Exit Sub
    Else
        
        lstrSql = "select distinct(收费编号),交费人,交费单位名称,主管科室编号 from  收费管理_费用信息表 where 收费编号='" & para收费编号 & "'"
        Set lobjRec = dafuncGetData(lstrSql)
        If lobjRec.RecordCount > 0 Then
            lstr目的科室号 = IIf(IsNull(lobjRec("主管科室编号")), "", lobjRec("主管科室编号"))
            lstr交费人 = IIf(IsNull(lobjRec("交费人")), "不详", lobjRec("交费人"))
            lstr交费单位 = IIf(IsNull(lobjRec("交费单位名称")), "不详", lobjRec("交费单位名称"))
            Para消息内容 = Para消息内容 & "　交费人" & lstr交费人 & "，交费单位：" & lstr交费单位 & "。"
            
            '由于信息服务是可选择安装,所以在发送时要检验对象是否存在 徐冀川 2002/09/17
            If Not um信使客户端 Is Nothing Then
                um信使客户端.sub发送消息 um用户所属科室编号, lstr目的科室号, Para消息内容, "费用修改"
            End If
        End If
    End If
Exit Sub
errhander:
    'sfsub错误处理 "收费界面", "frm收费", "sub消息发送", Err.Number, Err.Description
End Sub

