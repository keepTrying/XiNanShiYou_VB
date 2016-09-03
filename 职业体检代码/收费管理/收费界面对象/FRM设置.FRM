VERSION 5.00
Object = "{8099FCC2-0A81-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F48D4DEC-1198-11D5-91BE-0050BA06B70C}#5.8#0"; "录入控件.ocx"
Begin VB.Form frm设置 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "业务设置"
   ClientHeight    =   7800
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11055
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   11055
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.ImageList cimg按钮图标 
      Left            =   2640
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   21
      ImageHeight     =   21
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar ctlb设置 
      Align           =   1  'Align Top
      Height          =   615
      Left            =   0
      TabIndex        =   51
      Top             =   0
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   1085
      ButtonWidth     =   714
      ButtonHeight    =   926
      Appearance      =   1
      ImageList       =   "cimg按钮图标"
      _Version        =   393216
   End
   Begin TabDlg.SSTab cstbShezhi 
      Height          =   6765
      Left            =   30
      TabIndex        =   0
      Top             =   840
      Width           =   10905
      _ExtentX        =   19235
      _ExtentY        =   11933
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "收费项目设置"
      TabPicture(0)   =   "FRM设置.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "clbl使用说明"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "票据设置"
      TabPicture(1)   =   "FRM设置.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "收费标准设置"
      TabPicture(2)   =   "FRM设置.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame5"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "打折设置"
      TabPicture(3)   =   "FRM设置.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame1"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "其它设置"
      TabPicture(4)   =   "FRM设置.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame9"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      Begin VB.Frame Frame9 
         Height          =   5895
         Left            =   -74880
         TabIndex        =   66
         Top             =   480
         Width           =   10215
         Begin VB.TextBox Ctxt收据号值 
            Height          =   285
            Left            =   2040
            MaxLength       =   6
            TabIndex        =   68
            Top             =   1200
            Width           =   1815
         End
         Begin VB.Label Label20 
            Caption         =   "　　在新的一个财政年开始时，用户可以使用此功能，设置系统收据号当前值为０，表示收据号从１开始。"
            ForeColor       =   &H80000017&
            Height          =   495
            Left            =   4320
            TabIndex        =   71
            Top             =   1080
            Width           =   4575
         End
         Begin VB.Label Label19 
            Caption         =   "说明:"
            ForeColor       =   &H80000017&
            Height          =   255
            Left            =   4320
            TabIndex        =   70
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "系统收据号初始化"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000017&
            Height          =   240
            Left            =   360
            TabIndex        =   69
            Top             =   600
            Width           =   2040
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "系统收据号当前值："
            ForeColor       =   &H80000017&
            Height          =   180
            Left            =   360
            TabIndex        =   67
            Top             =   1200
            Width           =   1620
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   6135
         Left            =   -74880
         TabIndex        =   32
         Top             =   480
         Width           =   10700
         Begin 录入控件.ctlInputBox cinp交费单位 
            Height          =   360
            Left            =   5400
            TabIndex        =   50
            Top             =   360
            Width           =   3525
            _ExtentX        =   6218
            _ExtentY        =   635
            BackColor       =   16777215
            ForeColor       =   0
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            LeftOfTextbox   =   760
            Text            =   ""
            Label           =   "交费单位"
            Enabled         =   0   'False
            名称            =   ""
            长度            =   0
            允许等于最大值  =   0   'False
            允许等于最小值  =   0   'False
            允许多选        =   0   'False
         End
         Begin VB.VScrollBar cvsl打折 
            Enabled         =   0   'False
            Height          =   360
            Left            =   7320
            Max             =   1
            Min             =   100
            TabIndex        =   20
            Top             =   960
            Value           =   80
            Width           =   675
         End
         Begin VB.TextBox ctxt打折 
            Enabled         =   0   'False
            Height          =   360
            Left            =   6240
            TabIndex        =   19
            Top             =   960
            Width           =   1065
         End
         Begin VB.CommandButton ccmd单位定位 
            Caption         =   "单位定位"
            Enabled         =   0   'False
            Height          =   375
            Left            =   9000
            TabIndex        =   18
            Top             =   360
            Width           =   975
         End
         Begin VSFlex6Ctl.vsFlexGrid cFlg打折 
            Height          =   5685
            Left            =   210
            TabIndex        =   24
            Top             =   300
            Width           =   4995
            _cx             =   4203115
            _cy             =   4204332
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
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   1
            Cols            =   4
            FixedRows       =   1
            FixedCols       =   1
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
            DataMember      =   ""
         End
         Begin VB.Frame Frame8 
            Caption         =   "打折控制项"
            Height          =   1755
            Left            =   5400
            TabIndex        =   33
            Top             =   4080
            Width           =   4680
            Begin VB.OptionButton copt打折 
               Caption         =   "不打折"
               Height          =   180
               Index           =   0
               Left            =   150
               TabIndex        =   21
               Top             =   300
               Width           =   840
            End
            Begin VB.OptionButton copt打折 
               Caption         =   "可以打折，但不严格控制"
               Height          =   180
               Index           =   1
               Left            =   150
               TabIndex        =   22
               Top             =   735
               Width           =   2280
            End
            Begin VB.OptionButton copt打折 
               Caption         =   "可以打折，但严格控制"
               Height          =   180
               Index           =   2
               Left            =   150
               TabIndex        =   23
               Top             =   1185
               Width           =   2100
            End
         End
         Begin 录入控件.ctlInputBox cinb单位编号 
            Height          =   360
            Left            =   5460
            TabIndex        =   48
            Top             =   390
            Visible         =   0   'False
            Width           =   2325
            _ExtentX        =   4101
            _ExtentY        =   635
            BackColor       =   16777215
            ForeColor       =   0
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            LeftOfTextbox   =   760
            Text            =   ""
            Label           =   "单位编号"
            Enabled         =   0   'False
            名称            =   ""
            长度            =   0
            允许等于最大值  =   0   'False
            允许等于最小值  =   0   'False
            允许多选        =   0   'False
         End
         Begin VB.Label clbl打折说明 
            Height          =   1785
            Left            =   5400
            TabIndex        =   49
            Top             =   2280
            Width           =   5085
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "打折比率"
            Height          =   180
            Left            =   5400
            TabIndex        =   47
            Top             =   1080
            Width           =   720
         End
      End
      Begin VB.Frame Frame5 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   6255
         Left            =   -74880
         TabIndex        =   30
         Top             =   360
         Width           =   10605
         Begin VB.ListBox clst收费标准名称 
            Height          =   1860
            ItemData        =   "FRM设置.frx":008C
            Left            =   7680
            List            =   "FRM设置.frx":008E
            TabIndex        =   39
            Top             =   1080
            Visible         =   0   'False
            Width           =   2445
         End
         Begin VB.TextBox ctxt最大单价 
            Enabled         =   0   'False
            Height          =   405
            Left            =   8400
            TabIndex        =   44
            Top             =   1890
            Width           =   1395
         End
         Begin VB.TextBox ctxt最小单价 
            Enabled         =   0   'False
            Height          =   405
            Left            =   5760
            TabIndex        =   43
            Top             =   1890
            Width           =   1395
         End
         Begin VB.CommandButton Ccmd移除 
            Caption         =   "<"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3780
            TabIndex        =   13
            Top             =   3450
            Width           =   360
         End
         Begin VB.CommandButton Ccmd添加 
            Caption         =   ">"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3780
            TabIndex        =   12
            Top             =   2970
            Width           =   360
         End
         Begin VB.Frame Frame7 
            Height          =   1635
            Left            =   4200
            TabIndex        =   31
            Top             =   150
            Width           =   6225
            Begin VB.TextBox cinb收费标准名称 
               Enabled         =   0   'False
               Height          =   360
               Left            =   1245
               TabIndex        =   65
               Top             =   660
               Width           =   1845
            End
            Begin VB.TextBox cinb助记符 
               Enabled         =   0   'False
               Height          =   360
               Left            =   1245
               MaxLength       =   10
               TabIndex        =   63
               Top             =   1200
               Width           =   1845
            End
            Begin VB.CheckBox Cchk收费标准列表 
               Caption         =   "收费标准列表"
               Enabled         =   0   'False
               Height          =   345
               Left            =   3480
               Style           =   1  'Graphical
               TabIndex        =   17
               Top             =   600
               UseMaskColor    =   -1  'True
               Width           =   2460
            End
            Begin VB.TextBox ctxt子系统编号 
               Height          =   285
               Left            =   3540
               TabIndex        =   42
               Top             =   570
               Visible         =   0   'False
               Width           =   1065
            End
            Begin VB.ComboBox Ccmb使用范围 
               Enabled         =   0   'False
               Height          =   300
               Left            =   4410
               Style           =   2  'Dropdown List
               TabIndex        =   15
               Top             =   240
               Width           =   1575
            End
            Begin VB.ComboBox Ccmb子系统名 
               Enabled         =   0   'False
               Height          =   300
               Left            =   1245
               Style           =   2  'Dropdown List
               TabIndex        =   14
               Top             =   240
               Width           =   1845
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               Caption         =   "收费标准名称"
               Height          =   180
               Left            =   120
               TabIndex        =   64
               Top             =   750
               Width           =   1080
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "助记符"
               Height          =   180
               Left            =   120
               TabIndex        =   62
               Top             =   1290
               Width           =   540
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "使用范围"
               Height          =   180
               Left            =   3480
               TabIndex        =   41
               Top             =   300
               Width           =   720
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "子系统名"
               Height          =   180
               Left            =   120
               TabIndex        =   40
               Top             =   330
               Width           =   720
            End
         End
         Begin MSComctlLib.TreeView ctvw收费项目2 
            Height          =   5865
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   3570
            _ExtentX        =   6297
            _ExtentY        =   10345
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   353
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            Appearance      =   1
         End
         Begin VSFlex6Ctl.vsFlexGrid cflg收费标准 
            Height          =   3615
            Left            =   4200
            TabIndex        =   16
            Top             =   2400
            Width           =   6345
            _cx             =   4205496
            _cy             =   4200680
            _ConvInfo       =   1
            Appearance      =   1
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
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   1
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   1
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
            DataMember      =   ""
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "最大单价"
            Height          =   180
            Left            =   7530
            TabIndex        =   46
            Top             =   2010
            Width           =   720
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "当前项目最小单价"
            Height          =   180
            Left            =   4290
            TabIndex        =   45
            Top             =   2010
            Width           =   1440
         End
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   6210
         Left            =   -74880
         TabIndex        =   29
         Top             =   360
         Width           =   10700
         Begin VSFlex6Ctl.vsFlexGrid cFlg票据信息 
            Height          =   5955
            Left            =   120
            TabIndex        =   57
            Top             =   150
            Width           =   5505
            _cx             =   4204014
            _cy             =   4204808
            _ConvInfo       =   1
            Appearance      =   1
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
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   20
            Cols            =   7
            FixedRows       =   1
            FixedCols       =   1
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
            DataMember      =   ""
         End
         Begin VB.Frame Frame10 
            Appearance      =   0  'Flat
            Caption         =   "票据信息"
            ForeColor       =   &H80000008&
            Height          =   3930
            Left            =   5760
            TabIndex        =   34
            Top             =   120
            Width           =   4695
            Begin VB.TextBox ctxt最大项数 
               Enabled         =   0   'False
               Height          =   360
               Left            =   240
               MaxLength       =   2
               TabIndex        =   60
               Top             =   2280
               Width           =   1245
            End
            Begin VB.ComboBox ccmb票据信息票据类型 
               Enabled         =   0   'False
               Height          =   300
               Left            =   240
               Style           =   2  'Dropdown List
               TabIndex        =   59
               Top             =   1440
               Width           =   1305
            End
            Begin VB.TextBox cinb票据编号 
               Enabled         =   0   'False
               Height          =   360
               Left            =   240
               MaxLength       =   2
               TabIndex        =   58
               Top             =   600
               Width           =   1290
            End
            Begin MSComDlg.CommonDialog Ccmn票据 
               Left            =   2730
               Top             =   2250
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin VB.TextBox cinb票据名称 
               Enabled         =   0   'False
               Height          =   360
               Left            =   1800
               MaxLength       =   25
               TabIndex        =   55
               Top             =   600
               Width           =   2595
            End
            Begin VB.TextBox cinb票据格式文件名称 
               Enabled         =   0   'False
               Height          =   375
               Left            =   240
               MaxLength       =   24
               TabIndex        =   52
               Top             =   3120
               Width           =   4185
            End
            Begin VB.ComboBox ccmb对应业务 
               Enabled         =   0   'False
               Height          =   300
               Left            =   1800
               Style           =   2  'Dropdown List
               TabIndex        =   9
               Top             =   1440
               Width           =   2625
            End
            Begin VB.CommandButton Ccmd浏览 
               Caption         =   "浏览…"
               Enabled         =   0   'False
               Height          =   390
               Left            =   3360
               Style           =   1  'Graphical
               TabIndex        =   10
               Top             =   2640
               Width           =   1125
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "最大项数"
               Height          =   180
               Left            =   240
               TabIndex        =   61
               Top             =   2040
               Width           =   720
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "票据名称"
               Height          =   180
               Left            =   1800
               TabIndex        =   56
               Top             =   360
               Width           =   720
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "编号"
               Height          =   180
               Left            =   240
               TabIndex        =   54
               Top             =   360
               Width           =   360
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "格式文件名"
               Height          =   180
               Left            =   240
               TabIndex        =   53
               Top             =   2880
               Width           =   900
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "对应业务"
               Height          =   180
               Left            =   1800
               TabIndex        =   37
               Top             =   1200
               Width           =   720
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "票据类型"
               Height          =   180
               Left            =   240
               TabIndex        =   35
               Top             =   1200
               Width           =   720
            End
         End
         Begin VB.Label clbl票据设置说明 
            Caption         =   "Label4"
            Height          =   1725
            Left            =   5760
            TabIndex        =   38
            Top             =   4200
            Width           =   4545
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   5205
         Left            =   90
         TabIndex        =   25
         Top             =   1440
         Width           =   10575
         Begin VB.Frame Frame3 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   4965
            Left            =   6000
            TabIndex        =   26
            Top             =   120
            Width           =   4365
            Begin VB.ComboBox ccmb收费项目票据类型 
               Enabled         =   0   'False
               Height          =   300
               ItemData        =   "FRM设置.frx":0090
               Left            =   1380
               List            =   "FRM设置.frx":0092
               Style           =   2  'Dropdown List
               TabIndex        =   8
               Top             =   4260
               Width           =   2835
            End
            Begin 录入控件.ctlInputBox cinb收费项目编号 
               Height          =   360
               Left            =   240
               TabIndex        =   27
               Top             =   255
               Width           =   3975
               _ExtentX        =   7011
               _ExtentY        =   635
               BackColor       =   16777215
               ForeColor       =   0
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               LeftOfTextbox   =   1120
               Text            =   ""
               Label           =   "收费项目编号"
               Enabled         =   0   'False
               名称            =   ""
               长度            =   0
               允许等于最大值  =   0   'False
               允许等于最小值  =   0   'False
               允许多选        =   0   'False
            End
            Begin 录入控件.ctlInputBox cinb收费项目名称 
               Height          =   360
               Left            =   255
               TabIndex        =   2
               Top             =   827
               Width           =   3960
               _ExtentX        =   6985
               _ExtentY        =   635
               BackColor       =   16777215
               ForeColor       =   0
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               LeftOfTextbox   =   1120
               Text            =   ""
               Label           =   "收费项目名称"
               Enabled         =   0   'False
               名称            =   ""
               长度            =   10
               允许等于最大值  =   0   'False
               允许等于最小值  =   0   'False
               允许多选        =   0   'False
            End
            Begin 录入控件.ctlInputBox cinb收费项目助记符 
               Height          =   360
               Left            =   795
               TabIndex        =   3
               Top             =   1399
               Width           =   3420
               _ExtentX        =   6033
               _ExtentY        =   635
               BackColor       =   16777215
               ForeColor       =   0
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               LeftOfTextbox   =   580
               Text            =   ""
               Label           =   "助记符"
               Enabled         =   0   'False
               名称            =   ""
               长度            =   10
               允许等于最大值  =   0   'False
               允许等于最小值  =   0   'False
               允许多选        =   0   'False
            End
            Begin 录入控件.ctlInputBox cinb收费项目单价 
               Height          =   360
               Left            =   990
               TabIndex        =   4
               Top             =   1965
               Width           =   3240
               _ExtentX        =   5715
               _ExtentY        =   635
               BackColor       =   16777215
               ForeColor       =   0
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               LeftOfTextbox   =   400
               Text            =   ""
               Label           =   "单价"
               Enabled         =   0   'False
               名称            =   ""
               长度            =   5
               允许等于最大值  =   0   'False
               允许等于最小值  =   0   'False
               允许多选        =   0   'False
            End
            Begin 录入控件.ctlInputBox cinb收费项目最小单价 
               Height          =   360
               Left            =   615
               TabIndex        =   5
               Top             =   2550
               Width           =   3600
               _ExtentX        =   6350
               _ExtentY        =   635
               BackColor       =   16777215
               ForeColor       =   0
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               LeftOfTextbox   =   760
               Text            =   ""
               Label           =   "最小单价"
               Enabled         =   0   'False
               名称            =   ""
               长度            =   5
               允许等于最大值  =   0   'False
               允许等于最小值  =   0   'False
               允许多选        =   0   'False
            End
            Begin 录入控件.ctlInputBox cinb收费项目最大单价 
               Height          =   360
               Left            =   615
               TabIndex        =   6
               Top             =   3115
               Width           =   3600
               _ExtentX        =   6350
               _ExtentY        =   635
               BackColor       =   16777215
               ForeColor       =   0
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               LeftOfTextbox   =   760
               Text            =   ""
               Label           =   "最大单价"
               Enabled         =   0   'False
               名称            =   ""
               长度            =   5
               允许等于最大值  =   0   'False
               允许等于最小值  =   0   'False
               允许多选        =   0   'False
            End
            Begin 录入控件.ctlInputBox cinb收费项目计量单位 
               Height          =   360
               Left            =   615
               TabIndex        =   7
               Top             =   3690
               Width           =   3600
               _ExtentX        =   6350
               _ExtentY        =   635
               BackColor       =   16777215
               ForeColor       =   0
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               LeftOfTextbox   =   760
               Text            =   ""
               Label           =   "计量单位"
               Enabled         =   0   'False
               名称            =   ""
               长度            =   5
               允许等于最大值  =   0   'False
               允许等于最小值  =   0   'False
               允许多选        =   0   'False
            End
            Begin VB.Label Label2 
               Caption         =   "所属票据类型"
               Height          =   255
               Left            =   240
               TabIndex        =   36
               Top             =   4320
               Width           =   1095
            End
         End
         Begin MSComctlLib.TreeView cTvw收费项目1 
            Height          =   4815
            Left            =   120
            TabIndex        =   1
            Top             =   240
            Width           =   5775
            _ExtentX        =   10186
            _ExtentY        =   8493
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   353
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   6
            Appearance      =   1
         End
      End
      Begin VB.Label clbl使用说明 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   1065
         Left            =   105
         TabIndex        =   28
         Top             =   360
         Width           =   10305
      End
   End
End
Attribute VB_Name = "frm设置"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1
Public pblnInUse As Boolean

'定义权限控制数
Dim mint权限数 As Integer
Dim mbln退出 As Boolean

'定义操作收费项目所需要的变量
Dim mcol收费项目添加 As Collection
Dim mcol收费项目修改 As Collection
Dim mcol收费项目删除 As Collection
Dim mstr收费项目操作标志 As String

'定义操作票据信息所需要的变量
Dim mcol票据信息添加 As Collection
Dim mcol票据信息修改 As Collection
Dim mcol票据信息删除 As Collection
Dim mint票据信息行数 As Integer
Dim mstr票据信息操作标志 As String
Dim mstr编号() As Integer


'定义操作收费标准所需要的变量
Dim mrsd收费标准 As Recordset
Dim mstr收费标准操作方式标志 As String

'定义操作打折信息所需要的变量
Dim mcol打折信息添加 As Collection
Dim mcol打折信息修改 As Collection
Dim mcol打折信息删除 As Collection
Dim mstr打折比率 As String
Dim mstr打折信息操作标志 As String

'定义操作业务配置所需要的变量
Dim mcol业务配置 As Collection
Dim mint打折控制     As Integer
Dim mint项目级数     As Integer

'定义操作票据类型所需要的变量
Dim mrsd票据类型 As Recordset

'定义打折控制情况变量
Dim mstr打折控制操作标志 As String

'定义操作子系统条件的变量
Dim mrsd子系统条件 As Recordset
Private WithEvents mobjGUI As cls界面通用对象
Attribute mobjGUI.VB_VarHelpID = -1

Private pint打折控制 As Integer '保持业务配置的打折控制信息
Private pint科目级数 As Integer '保持业务配置的科目级数

Private Sub Cchk收费标准列表_Click()
    On Error GoTo errHandler
    clst收费标准名称.Visible = Cchk收费标准列表.Value
    If Cchk收费标准列表.Value Then
        clst收费标准名称.SetFocus
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "收费界面", "frm设置", "Cchk收费标准列表_Click", Err.Number, Err.Description, False
End Sub


Private Sub ccmb对应业务_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        ctxt最大项数.SetFocus
    End If
End Sub

Private Sub ccmb票据信息票据类型_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        ccmb对应业务.SetFocus
    End If
End Sub


Private Sub Ccmb使用范围_Click()
    On Error GoTo errHandler
    If mstr收费标准操作方式标志 = "" Then
        Exit Sub
    End If
    If Ccmb使用范围.Text = "" Then
        MsgBox "请选择使用范围", vbExclamation, "收费标准设置"
        
        Exit Sub
    End If
    mrsd子系统条件.Filter = "系统名='" & Ccmb子系统名.Text & "' and " & "条件名='" & Ccmb使用范围.Text & "'"
    If Not mrsd子系统条件.EOF And Not mrsd子系统条件.BOF Then
        ctxt子系统编号.Text = mrsd子系统条件("编号").Value
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "收费界面", "frm设置", "Ccmb使用范围_Change", Err.Number, Err.Description, False
End Sub

Private Sub Ccmb使用范围_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandler
    If KeyAscii = vbKeyReturn Then
        mrsd子系统条件.Filter = "系统名='" & Ccmb子系统名.Text & "' and " & "条件名='" & Ccmb使用范围.Text & "'"
        ctxt子系统编号.Text = mrsd子系统条件("编号").Value
        cinb收费标准名称.SetFocus
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "收费界面", "frm设置", "Ccmb使用范围_Change", Err.Number, Err.Description, False
End Sub

Private Sub Ccmb子系统名_Click()
    Dim lrsdtemp As Recordset
    On Error GoTo errHandler
    mrsd子系统条件.Filter = "系统名='" & Ccmb子系统名.Text & "'"
    Set lrsdtemp = mrsd子系统条件
    If (Not lrsdtemp.EOF) And (Not lrsdtemp.BOF) Then
    mrsd子系统条件.MoveFirst
    Ccmb使用范围.Clear
    Do While (Not mrsd子系统条件.EOF)
        Ccmb使用范围.AddItem mrsd子系统条件("条件名").Value
        mrsd子系统条件.MoveNext
    Loop
    End If
    Set lrsdtemp = Nothing
    If mstr收费标准操作方式标志 = "修改" Or mstr收费标准操作方式标志 = "添加" Then
        Ccmb使用范围.SetFocus
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "收费界面", "frm设置", "Ccmb子系统名_Change", Err.Number, Err.Description, False

End Sub



Private Sub Ccmb子系统名_KeyPress(KeyAscii As Integer)
    Dim lrsdtemp As Recordset
    On Error GoTo errHandler
    If KeyAscii = vbKeyReturn Then
        mrsd子系统条件.Filter = "系统名='" & Ccmb子系统名.Text & "'"
        Set lrsdtemp = mrsd子系统条件
        If (Not lrsdtemp.EOF) And (Not lrsdtemp.BOF) Then
            mrsd子系统条件.MoveFirst
            Do While (Not mrsd子系统条件.EOF)
                Ccmb使用范围.AddItem mrsd子系统条件("条件名").Value
                mrsd子系统条件.MoveNext
            Loop
        End If
        Set lrsdtemp = Nothing
        Ccmb使用范围.SetFocus
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "收费界面", "frm设置", "Ccmb子系统名_Change", Err.Number, Err.Description, False
End Sub

'显示单位定位窗口
Private Sub ccmd单位定位_Click()
    On Error GoTo errHandler
    
    Dim lobj单位信息 As Object
    '本功能弹出一个简单定位窗体，出现的坐标位置由参数决定
    Set lobj单位信息 = pobj单位定位.func单位简单定位(8600, 1000)
    If Not (lobj单位信息 Is Nothing) Then   '已获取单位信息
        cinp交费单位.Text = lobj单位信息.Fields!单位名称
        '处理空值,时间:2002/01/29,徐冀川
        cinb单位编号.Text = IIf(IsNull(lobj单位信息.Fields!申请编号), "", lobj单位信息.Fields!申请编号)
    Else                                                           '未获取单位信息
    
    End If

    Set lobj单位信息 = Nothing
    ctxt打折.SetFocus
    Exit Sub
errHandler:
    Set lobj单位信息 = Nothing
    sfsub错误处理 "收费界面", "frm设置", "ccmd单位定位_Click", Err.Number, Err.Description, False
End Sub


'向收费标准中添加一条记录
Private Sub Ccmd添加_Click()
    On Error GoTo errHandler
    Dim lstr收费项目编号
    Dim lrsdtemp As Recordset
    Dim lrsd收费项目名称 As Recordset
    Dim linttemp As Integer
    Dim lintcount As Integer
    If ctvw收费项目2.SelectedItem Is Nothing Then
        MsgBox "请选择收费项目", vbExclamation, "收费标准设置"
        Exit Sub
    End If
    lstr收费项目编号 = Right(ctvw收费项目2.SelectedItem.Key, Len(ctvw收费项目2.SelectedItem.Key) - 1)
    If Len(lstr收费项目编号) / 3 <> pint科目级数 Then
        MsgBox "只有具体的收费项目才能添加到收费标准中", vbExclamation, "收费标准设置"
        Exit Sub
    End If
    lintcount = cflg收费标准.Rows - 1
    For linttemp = 1 To lintcount
        If lstr收费项目编号 = cflg收费标准.Cell(flexcpText, linttemp, 2) Then
            MsgBox "收费项目不能重复", vbExclamation, "收费标准设置"
            Exit Sub
        End If
    Next
    Set lrsdtemp = pobj收费管理.func查询收费项目("收费项目编号='" & lstr收费项目编号 & "'")
    If (Not lrsdtemp.BOF) And (Not lrsdtemp.EOF) Then
        cflg收费标准.AddItem ""
        cflg收费标准.Cell(flexcpText, cflg收费标准.Rows - 1, 2) = lrsdtemp("收费项目编号").Value
        Set lrsd收费项目名称 = pobj收费管理.func查询收费项目("收费项目编号='" & lrsdtemp("收费项目编号").Value & "'")
        If (Not lrsd收费项目名称.EOF) And (Not lrsd收费项目名称.BOF) Then
            cflg收费标准.Cell(flexcpText, cflg收费标准.Rows - 1, 3) = lrsd收费项目名称("收费项目名称").Value
        End If
        cflg收费标准.Cell(flexcpText, cflg收费标准.Rows - 1, 4) = lrsdtemp("单价").Value
        cflg收费标准.Cell(flexcpText, cflg收费标准.Rows - 1, 5) = 1
        ctxt最小单价.Text = lrsdtemp("最小单价").Value
        ctxt最大单价.Text = lrsdtemp("最大单价").Value
        'cflg收费标准.Cell(flexcpChecked, cflg收费标准.Rows - 1, 6) = 1
        Ccmb子系统名.SetFocus
    End If
    ctvw收费项目2.SetFocus
    Exit Sub
errHandler:
    sfsub错误处理 "收费界面", "frm设置", "Ccmd添加_Click", Err.Number, Err.Description, False
End Sub

'从收费标准中移除一条项目
Private Sub Ccmd移除_Click()
    On Error GoTo errHandler
    If cflg收费标准.RowSel = 0 Then
        MsgBox "请选择项目", vbExclamation, "收费标准设置"
        Exit Sub
    End If
    cflg收费标准.RemoveItem (cflg收费标准.RowSel)
    Exit Sub
errHandler:
    sfsub错误处理 "收费界面", "frm设置", "Ccmd移除_Click", Err.Number, Err.Description, False
End Sub





'验证单价设置的有效性
Private Sub cflg收费标准_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    On Error GoTo errHandler
    If Col = 4 Then
        If IsNumeric(cflg收费标准.Cell(flexcpText, Row, Col)) Then
            If (Not ctxt最大单价.Text <> "") And (Not ctxt最小单价.Text <> "") Then
                If IsNumeric(ctxt最大单价.Text) And IsNumeric(ctxt最小单价.Text) Then
                    If CDbl(cflg收费标准.Cell(flexcpText, Row, Col)) > CDbl(ctxt最大单价.Text) Or CDbl(cflg收费标准.Cell(flexcpText, Row, Col)) < CDbl(ctxt最小单价.Text) Then
                        MsgBox "单价必须在最大单价和最小单价之间", vbExclamation, "收费标准设置"
                        cflg收费标准.Cell(flexcpText, Row, Col) = ""
                    End If
                End If
            End If
        Else
            MsgBox "单价必须为数值，请修改", vbExclamation, "收费标准设置"
            cflg收费标准.Cell(flexcpText, Row, Col) = ""
        End If
    End If
    
    If Col = 5 Then
        If IsNumeric(cflg收费标准.Cell(flexcpText, Row, Col)) Then
            If cflg收费标准.Cell(flexcpText, Row, Col) < 1 Then
                MsgBox "数量必须大于零，请修改", vbExclamation, "收费标准设置"
                cflg收费标准.Cell(flexcpText, Row, Col) = ""
            End If
        Else
            MsgBox "数量必须为数值，请修改", vbExclamation, "收费标准设置"
            cflg收费标准.Cell(flexcpText, Row, Col) = ""
        End If
    End If
    cflg收费标准.SetFocus
    cflg收费标准.Refresh
    Exit Sub
errHandler:
    sfsub错误处理 "收费界面", "frm设置", "cflg收费标准_AfterEdit", Err.Number, Err.Description, False
End Sub

'设置收费标准允许修改的列
Private Sub cflg收费标准_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    On Error GoTo errHandler
    If Col = 4 Or Col = 5 Then
        Cancel = False
    Else
        Cancel = True
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "收费界面", "frm设置", "cflg收费标准_BeforeEdit", Err.Number, Err.Description, False
End Sub



'点击收费标准网格，将所选的项目显示到相应的位置上
Private Sub cflg收费标准_Click()
    On Error GoTo errHandler
    Dim lrsd收费项目 As Recordset
    Set lrsd收费项目 = pobj收费管理.func查询收费项目("收费项目编号='" & cflg收费标准.Cell(flexcpText, cflg收费标准.RowSel, 2) & "'")
    If (Not lrsd收费项目.BOF) And (Not lrsd收费项目.EOF) Then
        ctxt最小单价.Text = lrsd收费项目("最小单价").Value
        ctxt最大单价.Text = lrsd收费项目("最大单价").Value
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "收费界面", "frm设置", "cflg收费标准_Click", Err.Number, Err.Description, False
End Sub




'Private Sub cinb票据格式文件名称_Change()
'
'End Sub

Private Sub cinb票据名称_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = 32 Then
        KeyAscii = 0
    End If
    If KeyAscii = vbKeyReturn Then
        ccmb票据信息票据类型.SetFocus
    End If
End Sub


Private Sub cinb收费标准名称_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = 32 Then
        KeyAscii = 0
    End If
    If KeyAscii = vbKeyReturn Then
        cinb助记符.SetFocus
    End If
End Sub



'设置焦点移动
Private Sub cinb收费项目单价_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And (mstr收费项目操作标志 = "修改" Or mstr收费项目操作标志 = "添加") Then
        cinb收费项目最小单价.SetFocus
    End If
End Sub

'设置焦点移动
Private Sub cinb收费项目计量单位_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = 32 Then
        KeyAscii = 0
    End If
    If KeyAscii = vbKeyReturn And (mstr收费项目操作标志 = "修改" Or mstr收费项目操作标志 = "添加") Then
        ccmb收费项目票据类型.SetFocus
    End If
End Sub





'设置焦点移动
Private Sub cinb收费项目名称_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = 32 Then
        KeyAscii = 0
    End If
    If KeyAscii = vbKeyReturn And (mstr收费项目操作标志 = "修改" Or mstr收费项目操作标志 = "添加") Then
        cinb收费项目助记符.SetFocus
    End If
End Sub

'设置焦点移动
Private Sub cinb收费项目助记符_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = 32 Then
        KeyAscii = 0
    End If
    If KeyAscii = vbKeyReturn And (mstr收费项目操作标志 = "修改" Or mstr收费项目操作标志 = "添加") Then
        cinb收费项目单价.SetFocus
    End If
End Sub

'设置焦点移动
Private Sub cinb收费项目最大单价_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And (mstr收费项目操作标志 = "修改" Or mstr收费项目操作标志 = "添加") Then
        cinb收费项目计量单位.SetFocus
    End If
End Sub

'设置焦点移动
Private Sub cinb收费项目最小单价_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And (mstr收费项目操作标志 = "修改" Or mstr收费项目操作标志 = "添加") Then
        cinb收费项目最大单价.SetFocus
    End If
End Sub

Private Sub cinb助记符_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = 32 Then
        KeyAscii = 0
    End If
End Sub

'显示所有的信息到相应的位置上
Private Sub clst收费标准名称_Click()
    On Error GoTo errHandler
    Dim lrsdtemp As Recordset
    Dim lrsd子系统条件 As Recordset
    Dim lrsd收费项目名称 As Recordset
    Dim linttemp As Integer
    cinb收费标准名称.Text = clst收费标准名称.Text
    Set lrsdtemp = pobj业务设置.func查询收费标准("收费标准名称='" & clst收费标准名称.Text & "'")
'    Set mrsd收费标准 = lrsdtemp
    If (Not lrsdtemp.BOF) And (Not lrsdtemp.EOF) Then
        lrsdtemp.MoveFirst
        Set lrsd子系统条件 = pobj业务设置.func查询子系统条件("编号=" & lrsdtemp("过滤条件").Value)
        cinb助记符.Text = lrsdtemp("助记符").Value
        'lrsd子系统条件.Filter="innerID="
        If (Not lrsd子系统条件.EOF) And (Not lrsd子系统条件.BOF) Then
            lrsd子系统条件.MoveFirst
            Ccmb子系统名.Text = lrsd子系统条件("系统名").Value
            Dim lstr子系统名 As String
            If Not lrsd子系统条件("条件名") Is Nothing Then
            lstr子系统名 = lrsd子系统条件("条件名").Value
            If Not lstr子系统名 = vbNullString Then
                Ccmb使用范围.Text = lrsd子系统条件("条件名").Value
            End If
            End If
            ctxt子系统编号.Text = lrsd子系统条件("编号").Value
            Set lrsd子系统条件 = Nothing
        Else
            MsgBox "子系统条件设置有问题，请重新选择", vbExclamation, "收费标准设置"
        End If
        'mrsd子系统条件.Filter = ""
        
        linttemp = 1
        cflg收费标准.Rows = lrsdtemp.RecordCount + 1
        Do While (Not lrsdtemp.EOF)
            cflg收费标准.Cell(flexcpText, linttemp, 1) = lrsdtemp("编号").Value
            cflg收费标准.Cell(flexcpText, linttemp, 2) = lrsdtemp("收费项目编号").Value
            Set lrsd收费项目名称 = pobj收费管理.func查询收费项目("收费项目编号='" & lrsdtemp("收费项目编号").Value & "'")
            If (Not lrsd收费项目名称.EOF) And (Not lrsd收费项目名称.BOF) Then
                cflg收费标准.Cell(flexcpText, linttemp, 3) = lrsd收费项目名称("收费项目名称").Value
                ctxt最大单价.Text = lrsd收费项目名称("最小单价").Value
                ctxt最大单价.Text = lrsd收费项目名称("最大单价").Value
            End If
            cflg收费标准.Cell(flexcpText, linttemp, 4) = lrsdtemp("单价").Value
            cflg收费标准.Cell(flexcpText, linttemp, 5) = lrsdtemp("数量").Value
            'cflg收费标准.Cell(flexcpChecked, linttemp, 6) = lrsdtemp("允许修改数量").Value
            linttemp = linttemp + 1
            lrsdtemp.MoveNext
        Loop
        cflg收费标准.AutoSize 0, cflg收费标准.Cols - 1
        Set lrsdtemp = Nothing
    End If
    Cchk收费标准列表.Value = 0
    clst收费标准名称.Visible = False
    Exit Sub
errHandler:
    sfsub错误处理 "收费界面", "frm设置", "clst收费标准名称_Click", Err.Number, Err.Description, False
End Sub


'记录用户当前操作设置界面

'*************************************************************************
'*********************设置收费项目时的事件********************************
'*************************************************************************

'状态初始化为添加收费项目,只做3方面工作
'1、将一些项变为可使用
'2、将输入框置为空
'3、删除和修改不可用
Private Sub sub收费项目添加()
    On Error GoTo errHandler
    '如下这些项可录入
    'cinb收费项目编号.Enabled = True
    cinb收费项目名称.Enabled = True
    cinb收费项目单价.Enabled = True
    cinb收费项目计量单位.Enabled = True
    cinb收费项目助记符.Enabled = True
    cinb收费项目最大单价.Enabled = True
    cinb收费项目最小单价.Enabled = True
    ccmb收费项目票据类型.Enabled = True
    
    '如下这些项置为空
    cinb收费项目编号.Text = ""
    cinb收费项目名称.Text = ""
    cinb收费项目单价.Text = ""
    cinb收费项目计量单位.Text = ""
    cinb收费项目助记符.Text = ""
    cinb收费项目最大单价.Text = ""
    cinb收费项目最小单价.Text = ""
    'ccmb收费项目票据类型.Text = ""
    
    '只有保存和清除可用
    ctlb设置.Buttons("修改(&U)").Enabled = False
    ctlb设置.Buttons("删除(&D)").Enabled = False
    ctlb设置.Buttons("添加(&A)").Enabled = False
    mstr收费项目操作标志 = "添加"
    Exit Sub
errHandler:
    sfsub错误处理 "收费界面", "frm设置", "cmenuadd_Click", Err.Number, Err.Description, False
End Sub
'Private Sub cmenuadd_Click()
'    On Error GoTo errHandler
'    '如下这些项可录入
'    'cinb收费项目编号.Enabled = True
'    cinb收费项目名称.Enabled = True
'    cinb收费项目单价.Enabled = True
'    cinb收费项目计量单位.Enabled = True
'    cinb收费项目助记符.Enabled = True
'    cinb收费项目最大单价.Enabled = True
'    cinb收费项目最小单价.Enabled = True
'    ccmb收费项目票据类型.Enabled = True
'
'    '如下这些项置为空
'    cinb收费项目编号.Text = ""
'    cinb收费项目名称.Text = ""
'    cinb收费项目单价.Text = ""
'    cinb收费项目计量单位.Text = ""
'    cinb收费项目助记符.Text = ""
'    cinb收费项目最大单价.Text = ""
'    cinb收费项目最小单价.Text = ""
'    'ccmb收费项目票据类型.Text = ""
'
'    '只有保存和清除可用
'    ctlb设置.Buttons("修改(&U)").Enabled = False
'    ctlb设置.Buttons("删除(&D)").Enabled = False
'    ctlb设置.Buttons("添加(&A)").Enabled = False
'    mstr收费项目操作标志 = "添加"
'    Exit Sub
'errHandler:
'    sfsub错误处理 "收费界面", "frm设置", "cmenuadd_Click", Err.Number, Err.Description, False
'End Sub

'删除收费项目
Private Sub sub删除收费项目()
    On Error GoTo errHandler
    '只有保存和清除可用
    ctlb设置.Buttons("修改(&U)").Enabled = False
    ctlb设置.Buttons("删除(&D)").Enabled = False
    ctlb设置.Buttons("添加(&A)").Enabled = False
    cTvw收费项目1.Enabled = False
    mstr收费项目操作标志 = "删除"
    Exit Sub
errHandler:
    sfsub错误处理 "收费界面", "frm设置", "sub删除收费项目", Err.Number, Err.Description, False

End Sub


'修改收费项目
Private Sub sub修改收费项目()
    On Error GoTo errHandler
      '如下这些项可录入
    'cinb收费项目编号.Enabled = True
    cinb收费项目名称.Enabled = True
    cinb收费项目单价.Enabled = True
    cinb收费项目计量单位.Enabled = True
    cinb收费项目助记符.Enabled = True
    cinb收费项目最大单价.Enabled = True
    cinb收费项目最小单价.Enabled = True
    ccmb收费项目票据类型.Enabled = True
    
    '只有保存和清除可用
    ctlb设置.Buttons("修改(&U)").Enabled = False
    ctlb设置.Buttons("删除(&D)").Enabled = False
    ctlb设置.Buttons("添加(&A)").Enabled = False
    mstr收费项目操作标志 = "修改"
    Exit Sub
errHandler:
    sfsub错误处理 "收费界面", "frm设置", "cmenuupdate_Click", Err.Number, Err.Description, False
End Sub

Private Sub copt打折_Click(Index As Integer)
    If cstbShezhi.TabEnabled(3) = False Or mbln退出 Then
        Exit Sub
    End If
    mint打折控制 = Index
    ctlb设置.Buttons("添加(&A)").Enabled = False
    ctlb设置.Buttons("修改(&U)").Enabled = False
    ctlb设置.Buttons("删除(&D)").Enabled = False
    mstr打折控制操作标志 = "添加"
End Sub


'设置焦点移动
Private Sub cstbShezhi_KeyPress(KeyAscii As Integer)
    If mbln退出 Then
        Exit Sub
    End If
    If KeyAscii = vbKeyReturn Then
        Select Case cstbShezhi.Caption
            Case "收费项目设置"
                cTvw收费项目1.SetFocus
            Case "票据设置"
                cFlg票据信息.SetFocus
            Case "收费标准设置"
               ctvw收费项目2.SetFocus
            Case "打折设置"
                ccmd单位定位.SetFocus
            
        End Select
    End If
End Sub







Private Sub cTvw收费项目1_Click()
    On Error GoTo errHandler
    If (cstbShezhi.TabEnabled(0) = True) And (Not mbln退出) Then
        subtvw收费项目点击
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "收费界面", "frm设置", "cTvw收费项目1_click", Err.Number, Err.Description, False
End Sub

'调用subtvw收费项目点击
Private Sub cTvw收费项目1_KeyDown(KeyCode As Integer, Shift As Integer)
    On Err GoTo errhandle
    If KeyCode = vbKeyReturn Then
        subtvw收费项目点击
        If mstr收费项目操作标志 = "修改" Or mstr收费项目操作标志 = "添加" Then
            cinb收费项目名称.SetFocus
        End If
    End If
    Exit Sub
errhandle:
    sfsub错误处理 "收费界面", "frm设置", "ctvw收费项目1_keydown", Err.Number, Err.Description, False
End Sub

'功能：将收费项目信息显示在录入框中
'输入：无
'处理过程：
'   1、查找记录
'   2、显示
'输出：无
'返回：无
'创建人：何嘉
'创建时间：2001-04-04
Private Sub subtvw收费项目点击()
    Dim lintcount As Integer
    Dim lrsd收费项目 As Recordset
    Dim lrsd票据类型 As Recordset
    On Error GoTo errHandler
    If ((Not cTvw收费项目1.SelectedItem Is Nothing) And (Not cTvw收费项目1.SelectedItem.Key = "s")) Then
        Set lrsd收费项目 = pobj收费管理.func查询收费项目("收费项目编号='" & Right(LTrim(RTrim(cTvw收费项目1.SelectedItem.Key)), Len(LTrim(RTrim(cTvw收费项目1.SelectedItem.Key))) - 1) & "'")
        If (Not lrsd收费项目.EOF) And (Not lrsd收费项目.BOF) Then
            cinb收费项目编号.Text = lrsd收费项目.Fields("收费项目编号").Value
            cinb收费项目名称.Text = lrsd收费项目.Fields("收费项目名称").Value
            cinb收费项目单价.Text = lrsd收费项目.Fields("单价").Value
            cinb收费项目计量单位.Text = lrsd收费项目.Fields("计量单位").Value
            cinb收费项目最小单价.Text = lrsd收费项目.Fields("最小单价").Value
            cinb收费项目助记符.Text = lrsd收费项目.Fields("助记符").Value
            cinb收费项目最大单价.Text = lrsd收费项目.Fields("最大单价").Value
'            If (Not mrsd票据类型.EOF) And (Not mrsd票据类型.BOF) Then
'                mrsd票据类型.MoveFirst
'            End If
            Set lrsd票据类型 = pobj业务设置.func提取字典表信息("票据类型字典表")
            If (Not lrsd票据类型.BOF) And (Not lrsd票据类型.EOF) Then
                lrsd票据类型.MoveFirst
                Do While (Not lrsd票据类型.EOF)
                    If lrsd票据类型("InnerID").Value = lrsd收费项目("票据类型编号").Value Then
                        Exit Do
                    Else
                        lrsd票据类型.MoveNext
                    End If
                    If lrsd票据类型.EOF Then
                        If (Len(LTrim(RTrim(cTvw收费项目1.SelectedItem.Key))) - 1) / 3 = pint科目级数 Then
                            MsgBox "如果修改了票据类型字典表，请重新录入该项目的票据类型,该项目票据类型有问题", vbExclamation, "票据类型设置"
                        End If
                        Set lrsd票据类型 = Nothing
                        Exit Sub
                    End If
                Loop
                If Not lrsd票据类型("名称") Is Nothing Then
                    If lrsd票据类型("名称").Value <> "" Then
                        ccmb收费项目票据类型.Text = lrsd票据类型("名称").Value
                    End If
                End If
            End If

        End If
    End If
    Set lrsd票据类型 = Nothing
    Exit Sub
errHandler:
    sfsub错误处理 "收费界面", "frm设置", "cTvw收费项目1_MouseDown", Err.Number, Err.Description, True
End Sub


'*************************************************************************
'******************************以下为点击stab时的事件*********************
'*************************************************************************
Private Sub cstbShezhi_Click(PreviousTab As Integer)
    On Error GoTo errHandler
    If mbln退出 Then
        Exit Sub
    End If
    '对前一个tab页的处理
    If PreviousTab <> cstbShezhi.Tab Then
        Select Case cstbShezhi.TabCaption(PreviousTab)
            Case "收费项目设置"
                cinb收费项目编号.Enabled = False
                cinb收费项目名称.Enabled = False
                cinb收费项目单价.Enabled = False
                cinb收费项目计量单位.Enabled = False
                cinb收费项目助记符.Enabled = False
                cinb收费项目最大单价.Enabled = False
                cinb收费项目最小单价.Enabled = False
                ccmb收费项目票据类型.Enabled = False
                
                '如下这些项置为空
                cinb收费项目编号.Text = ""
                cinb收费项目名称.Text = ""
                cinb收费项目单价.Text = ""
                cinb收费项目计量单位.Text = ""
                cinb收费项目助记符.Text = ""
                cinb收费项目最大单价.Text = ""
                cinb收费项目最小单价.Text = ""
                
                mstr收费项目操作标志 = ""
            Case "票据设置"
                sub票据信息清除
            Case "收费标准设置"
                sub收费标准清除
            Case "打折设置"
                sub打折信息清除
            'Case "其它设置"
                
        End Select
    End If
    

    ctlb设置.Buttons(1).Enabled = True
    ctlb设置.Buttons(2).Enabled = True
    ctlb设置.Buttons(3).Enabled = True
    ctlb设置.Buttons(4).Enabled = True
    ctlb设置.Buttons(6).Enabled = True
    ctlb设置.Buttons(8).Enabled = True
    ctlb设置.Refresh
    Exit Sub
errHandler:
    sfsub错误处理 "收费界面", "frm设置", "cstbShezhi_Click", Err.Number, Err.Description, False
End Sub
'功能：　获取当前系统中的收据号
'作者：　徐冀川
'时间：2002/07/17
Private Sub sub获取当前收据号()
On Error GoTo errHandler
    Dim lint当前收据号 As Long          '定义变量记录当前收据号
    Dim lstrSql As String               '定义变量记录SQL语句
    Dim lobjRec As Object               '定义对象记录临时结果集
    
    lstrSql = "select 当前值 from 系统管理_系统编号生成记录表 where " & _
              " 业务名称= '收费管理' and 编号名称='收据号'"
    Set lobjRec = dafuncGetData(lstrSql)
    
    lint当前收据号 = IIf(IsNull(lobjRec("当前值")), 1, lobjRec("当前值"))
    Ctxt收据号值.Text = lint当前收据号
        
Exit Sub
errHandler:
    sfsub错误处理 "收费界面", "frm设置", "sub获取当前收据号", Err.Number, Err.Description, False
End Sub


Private Sub cstbShezhi_GotFocus()
    On Error GoTo errHandler
    If mbln退出 Then
        Exit Sub
    End If
    If mint权限数 = 0 Then
        Exit Sub
    End If

    Dim lbutton As MSComctlLib.Button
    For Each lbutton In ctlb设置.Buttons
        lbutton.Enabled = True
    Next
    Set lbutton = Nothing
'    End Select
    Exit Sub
errHandler:
    sfsub错误处理 "收费界面", "frm设置", "cstbShezhi_GotFocus", Err.Number, Err.Description, False
End Sub
'*************************************************************************

Private Sub ctvw收费项目2_Click()
    If cstbShezhi.TabEnabled(2) = False Or mbln退出 Then
        Exit Sub
    End If
    If (Not ctvw收费项目2.SelectedItem Is Nothing) And ((Len(ctvw收费项目2.SelectedItem.Key) - 1) / 3 = pint科目级数) And (mstr收费标准操作方式标志 = "添加" Or mstr收费标准操作方式标志 = "修改") Then
        Ccmd添加.SetFocus
    End If
End Sub

Private Sub ctvw收费项目2_KeyPress(KeyAscii As Integer)
    If cstbShezhi.TabEnabled(2) = False Or mbln退出 Then
        Exit Sub
    End If
    If KeyAscii = vbKeyReturn And (Not ctvw收费项目2.SelectedItem Is Nothing) And ((Len(ctvw收费项目2.SelectedItem.Key) - 1) / 3 = pint科目级数) And (mstr收费标准操作方式标志 = "添加" Or mstr收费标准操作方式标志 = "修改") Then
        Ccmd添加.SetFocus
    End If
End Sub

Private Sub ctxt打折_GotFocus()
    On Error GoTo errHandler
    If ctxt打折.Text <> "" Then
    mstr打折比率 = ctxt打折.Text
    Else
        mstr打折比率 = ""
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "收费界面", "frm设置", "ctxt打折_GotFocus", Err.Number, Err.Description, False
End Sub

Private Sub ctxt打折_LostFocus()
    On Error GoTo errHandler
    If ctxt打折 <> "" Then
        If ctxt打折.Text = "0" Then
            MsgBox "打折比率不能为空"
            ctxt打折.Text = ""
        End If
        mstr打折比率 = ""
    Else
        mstr打折比率 = ""
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "收费界面", "frm设置", "ctxt打折_LostFocus", Err.Number, Err.Description, False
End Sub


Private Sub Ctxt收据号值_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If Len(Trim(Ctxt收据号值.Text)) < 6 Then
        If (KeyAscii <= 47 Or KeyAscii > 57) And KeyAscii <> 13 And KeyAscii <> 8 Then
           KeyAscii = 0
        End If
    Else
        KeyAscii = 8
    End If
End Sub

Private Sub ctxt最大项数_Change()
    If Not ctxt最大项数 = "" Then
        If Not IsNumeric(ctxt最大项数.Text) Then
            ctxt最大项数.Text = ""
        Else
            If CInt(ctxt最大项数.Text) < 1 Then
                ctxt最大项数.Text = ""
            End If
        End If
    End If
End Sub

Private Sub ctxt最大项数_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Ccmd浏览.SetFocus
    End If
End Sub
Private Sub cvsl打折_Change()
    On Error GoTo errHandler
    ctxt打折.Text = cvsl打折.Value / 100
    Exit Sub
errHandler:
    sfsub错误处理 "收费界面", "frm设置", "cvsl打折_Change", Err.Number, Err.Description, False
End Sub

'******************************************************************************
'***********************************   初始化   *******************************
'******************************************************************************

'初始化界面，有如下内容，使用说明，收费项目级数，打折控制，收费项目treeview1,2
'票据信息，打折信息
Private Sub Form_Load()
'    Dim lint权限数 As Integer
    
    Dim lobj业务配置 As Object
    '用于初始化收费项目
    Dim lrsd收费项目 As Recordset
    Dim lrsd子系统条件 As Recordset
    Dim lrsd收费标准名称 As Recordset
    Dim lint级数 As Integer
    
    '以下变量用于初始化ctvw收费项目
    Dim lnodtemp As Node
    Dim lnodparent As Node
    Dim lstrtvwparent As String
    Dim lstrtvwmeposition As String
    Dim lstrtvwme As String
    
    '以下变量用于初始化票据类型
    Dim lrsd票据类型 As Recordset
    Dim lcol票据类型 As Collection
    
    Dim lrsdtemp As Recordset
    '以下变量用于初始化票据信息
    Dim lrsd票据信息 As Recordset
    
    '工具栏上的按钮初始化集合
    Dim lcol工具栏按钮 As New Collection
    On Error GoTo errHandler
    
    mbln退出 = False
    If pblnInUse = True Then Exit Sub
    
    pblnInUse = True
    
'    Set pobj收费管理 = CreateObject("收费业务对象.cls收费管理")
'    Set pobj业务设置 = CreateObject("收费业务对象.cls业务设置")
'    Set pobj单位定位 = CreateObject("单位档案业务.ClsUnitInterface")
    Set lobj业务配置 = CreateObject("收费数据对象.cls业务配置")
    
    pint打折控制 = lobj业务配置.打折控制
    pint科目级数 = lobj业务配置.科目级数
    
    Set lobj业务配置 = Nothing
     
   
     '权限验证
    mint权限数 = 0
    cstbShezhi.TabEnabled(0) = False
    cstbShezhi.TabEnabled(1) = False
    cstbShezhi.TabEnabled(2) = False
    cstbShezhi.TabEnabled(3) = False


    If umfunc校验用户权限("收费管理_打折设置") Then
        cstbShezhi.TabEnabled(3) = True
        mint权限数 = mint权限数 + 1
        cstbShezhi.Tab = 3

    End If

    If umfunc校验用户权限("收费管理_收费标准设置") Then
        cstbShezhi.TabEnabled(2) = True
        mint权限数 = mint权限数 + 1
        cstbShezhi.Tab = 2
    End If

    If umfunc校验用户权限("收费管理_票据设置") Then
        cstbShezhi.TabEnabled(1) = True
        mint权限数 = mint权限数 + 1
        cstbShezhi.Tab = 1
    End If

    If umfunc校验用户权限("收费管理_收费项目设置") Then
        cstbShezhi.TabEnabled(0) = True
        mint权限数 = mint权限数 + 1
        cstbShezhi.Tab = 0
    End If


    If mint权限数 = 0 Then
        Set mobjGUI = New cls界面通用对象
        Set mobjGUI.Form = Me
        Set mobjGUI.c工具栏 = ctlb设置
        lcol工具栏按钮.Add "退出"
        mobjGUI.subInitialize lcol工具栏按钮, ""
        sffuncMsg "请点击退出,检查是否配置配置好票据类型字典表或者收费系统各项权限再使用！", sf警告
        MsgBox "你无权使用该界面", vbExclamation, "设置"
        cstbShezhi.Visible = False
        mbln退出 = True
        GoTo WayOut
    End If
    
    Set mrsd票据类型 = pobj业务设置.func提取字典表信息("票据类型字典表")
    If Not mrsd票据类型 Is Nothing Then
        If mrsd票据类型.RecordCount = 0 Then
            Set mobjGUI = New cls界面通用对象
            Set mobjGUI.Form = Me
            Set mobjGUI.c工具栏 = ctlb设置
            lcol工具栏按钮.Add "退出"
            mobjGUI.subInitialize lcol工具栏按钮, ""
            sffuncMsg "请点击退出,配置好票据类型字典表再使用", sf警告
            copt打折(0).Enabled = False
            copt打折(1).Enabled = False
            copt打折(2).Enabled = False
            mbln退出 = True
            cstbShezhi.Visible = False
            GoTo WayOut
        End If
    Else
        Set mobjGUI = New cls界面通用对象
        Set mobjGUI.Form = Me
        Set mobjGUI.c工具栏 = ctlb设置
        lcol工具栏按钮.Add "退出"
        mobjGUI.subInitialize lcol工具栏按钮, ""
        sffuncMsg "请点击退出,执行收费系统安装程序再使用", sf警告
        copt打折(0).Enabled = False
        copt打折(1).Enabled = False
        copt打折(2).Enabled = False
        mbln退出 = True
        cstbShezhi.Visible = False
        GoTo WayOut
    End If
    
    '初始化工具栏
    Set mobjGUI = New cls界面通用对象
    Set mobjGUI.Form = Me
    Set mobjGUI.c工具栏 = ctlb设置
    'Set mobjGUI.c状态栏 = cstuShoufei

    lcol工具栏按钮.Add "添加"
    lcol工具栏按钮.Add "修改"
    lcol工具栏按钮.Add "删除"
    lcol工具栏按钮.Add "清空"
    lcol工具栏按钮.Add "|"
    lcol工具栏按钮.Add "保存"
    lcol工具栏按钮.Add "|"
    lcol工具栏按钮.Add "退出"
    
    mobjGUI.subInitialize lcol工具栏按钮, ""
    
    frm提示.Show , frm设置
    
    '初始化使用说明
    
    clbl使用说明.Caption = "使用说明：" & vbCrLf & vbTab & "1、添加子项：选择“添加”，在左边树状列表中点击左键选中项目，输入内容后选择保存。" & vbCrLf & vbTab & "2、删除子项：选择“删除”，在左边树状列表中选中项目后再点击保存。" & vbCrLf & vbTab & "3、修改子项：选择修改，在左边树状列表中选中一个项目，在右边输入框中修改后选择保存。" & vbCrLf & vbTab & "4、当前最大科目级数为" & pint科目级数
'    clbl收费项目级数.Caption = "注意：1、收费项目级数最大只能到""5""级。" & vbCrLf & vbCrLf & String(6, vbTab) & "2、在已有收费项目的情况下，" & vbCrLf & String(9, vbTab) & "如更改收费项目级数，" & vbCrLf & String(9, vbTab) & "将导致以前的收费项目丢失！"
    clbl票据设置说明.Caption = vbCrLf & vbTab & "使用说明：" & vbCrLf & vbCrLf & vbTab & "1、添加票据信息：点击添加，输入相应的项目。" & vbCrLf & vbTab & "2、修改：在左边列表中单击修改项，修改相应项目。" & vbCrLf & vbTab & "3、删除：点击删除项。" & vbCrLf & vbTab & "4、保存：所有操作必须保存后有效。"
    clbl打折说明.Caption = vbCrLf & vbTab & "使用说明：" & vbCrLf & vbCrLf & vbTab & "1、添加打折信息：点击添加，输入相应的项目。" & vbCrLf & vbTab & "2、修改：在左边列表中单击修改项，修改相应项目。" & vbCrLf & vbTab & "3、删除：点击删除按钮，保存后生效。" & vbCrLf & vbTab & "4、保存：所有操作必须保存后有效。" & vbCrLf & vbTab & "5、打折控制：打折控制的设置并不针对具体的单位,       具体单位的打折情况只有打折控制项为第3种选项       时才有效。"
    '初始化收费项目级数
'    cinp收费项目级数.Text = pint科目级数
    
    '初始化打折控制
    copt打折(pint打折控制).Value = True
    mint打折控制 = pint打折控制
    frm提示.cprg提示.Value = frm提示.cprg提示.Value + 1
    frm提示.Refresh
    '初始化打折信息
    Set lrsdtemp = pobj业务设置.func查询打折信息
    cFlg打折.Cell(flexcpText, 0, 1) = "单位编号"
    cFlg打折.Cell(flexcpText, 0, 2) = "单位名称"
    cFlg打折.Cell(flexcpText, 0, 3) = "打折比率"
    cFlg打折.AutoSize 0, cFlg打折.Cols - 1
    If (Not lrsdtemp.EOF) And (Not lrsdtemp.BOF) Then
        Dim linttemp As Integer
        linttemp = 1
        lrsdtemp.MoveFirst
        Do While (Not lrsdtemp.EOF)
            cFlg打折.AddItem ""
            cFlg打折.Cell(flexcpText, linttemp, 1) = lrsdtemp("单位编号").Value
            cFlg打折.Cell(flexcpText, linttemp, 2) = lrsdtemp("单位名称").Value
            cFlg打折.Cell(flexcpText, linttemp, 3) = lrsdtemp("打折比率").Value
            lrsdtemp.MoveNext
            linttemp = linttemp + 1
            
        Loop
        For linttemp = 1 To 3
            cFlg打折.AutoSize (linttemp)
        Next
    End If
   frm提示.cprg提示.Value = frm提示.cprg提示.Value + 4
   frm提示.Refresh
    '初始化对应业务
    ccmb对应业务.AddItem "一般", 0
    ccmb对应业务.AddItem "门诊", 1
    
    '初始化子系统条件
    Set mrsd子系统条件 = pobj业务设置.func查询子系统条件
    Set lrsd子系统条件 = pobj业务设置.func查询子系统条件(, "select 子系统名 from 收费管理_子系统过滤条件表")
    If (Not lrsd子系统条件.BOF) And (Not lrsd子系统条件.EOF) Then
        lrsd子系统条件.MoveFirst
        Do While (Not lrsd子系统条件.EOF)
            Ccmb子系统名.AddItem lrsd子系统条件("系统名").Value
            lrsd子系统条件.MoveNext
        Loop
    End If
    frm提示.cprg提示.Value = frm提示.cprg提示.Value + 3
    frm提示.Refresh
    '初始化cflg收费标准列表
    cflg收费标准.Cell(flexcpText, 0, 1) = "编号"
    cflg收费标准.Cell(flexcpText, 0, 2) = "收费项目编号"
    cflg收费标准.Cell(flexcpText, 0, 3) = "收费项目名称"
    cflg收费标准.Cell(flexcpText, 0, 4) = "单价"
    cflg收费标准.Cell(flexcpText, 0, 5) = "数量"
'    cflg收费标准.Cell(flexcpText, 0, 6) = "允许修改数量"
'    cflg收费标准.ColWidth(6) = 1
    
    '初始化收费标准名称列表
    Set lrsd收费标准名称 = pobj业务设置.func查询收费标准(, "select 收费标准名称 from 收费管理_收费标准信息表 group by 收费标准名称")
    If (Not lrsd收费标准名称.BOF) And (Not lrsd收费标准名称.EOF) Then
        lrsd收费标准名称.MoveFirst
        Do While (Not lrsd收费标准名称.EOF)
            clst收费标准名称.AddItem lrsd收费标准名称("收费标准名称").Value
            lrsd收费标准名称.MoveNext
        Loop
    End If
    frm提示.cprg提示.Value = frm提示.cprg提示.Value + 3
    frm提示.Refresh
     '初始化票据信息
    Set lrsd票据信息 = pobj业务设置.func查询票据信息
    cFlg票据信息.Cell(flexcpText, 0, 1) = "编号"
    cFlg票据信息.Cell(flexcpText, 0, 2) = "票据类型编号"
    cFlg票据信息.Cell(flexcpText, 0, 3) = "票据名称"
    cFlg票据信息.Cell(flexcpText, 0, 4) = "票据格式文件名称"
    cFlg票据信息.Cell(flexcpText, 0, 5) = "对应业务"
    cFlg票据信息.Cell(flexcpText, 0, 6) = "最大项数"
    Set cFlg票据信息.DataSource = lrsd票据信息
    cFlg票据信息.Rows = lrsd票据信息.RecordCount + 1
    mint票据信息行数 = lrsd票据信息.RecordCount + 1
    cFlg票据信息.Refresh
    cFlg票据信息.AutoSize 0, cFlg票据信息.Cols - 1
    '初始化票据类型
    
    If (Not mrsd票据类型.EOF) And (Not mrsd票据类型.BOF) Then
        mrsd票据类型.MoveFirst
        Do While (Not mrsd票据类型.EOF)
            ccmb收费项目票据类型.AddItem mrsd票据类型.Fields("名称").Value
            ccmb收费项目票据类型.ItemData(ccmb收费项目票据类型.NewIndex) = mrsd票据类型.Fields("innerId").Value
            ccmb票据信息票据类型.AddItem mrsd票据类型.Fields("名称").Value
            ccmb票据信息票据类型.ItemData(ccmb票据信息票据类型.NewIndex) = mrsd票据类型.Fields("innerId").Value
            mrsd票据类型.MoveNext
        Loop
        mrsd票据类型.MoveFirst
    End If
    
    frm提示.cprg提示.Value = frm提示.cprg提示.Value + 2
    frm提示.Refresh
    
     '初始化收费项目
    Set lnodparent = cTvw收费项目1.Nodes.Add(, , "s", "收费项目")
    'Set lnodparent = Nothing
    For lint级数 = 1 To pint科目级数
        Set lrsd收费项目 = pobj收费管理.func查询收费项目("Len(收费项目编号) =" & lint级数 * 3 & " order by 收费项目编号")
        If (Not lrsd收费项目.BOF) And (Not lrsd收费项目.EOF) Then
            lrsd收费项目.MoveFirst
            Do While (Not lrsd收费项目.EOF)
                lstrtvwme = "s" & lrsd收费项目("收费项目编号").Value
                Set lnodtemp = cTvw收费项目1.Nodes.Add("s" & Mid(lstrtvwme, 2, ((lint级数 - 1) * 3)), tvwChild, lstrtvwme, lrsd收费项目("收费项目名称").Value)
                lrsd收费项目.MoveNext
                Set lnodtemp = Nothing
                cTvw收费项目1.Refresh
            Loop
        End If
    Next
    Set lnodparent = ctvw收费项目2.Nodes.Add(, , "s", "收费项目")
    Set lnodparent = Nothing
    If Not lrsd收费项目 Is Nothing Then
        For lint级数 = 1 To pint科目级数
            Set lrsd收费项目 = pobj收费管理.func查询收费项目("Len(收费项目编号) =" & lint级数 * 3 & " order by 收费项目编号")
            If (Not lrsd收费项目.BOF) And (Not lrsd收费项目.EOF) Then
                lrsd收费项目.MoveFirst
                Do While (Not lrsd收费项目.EOF)
                    lstrtvwme = "s" & lrsd收费项目("收费项目编号").Value
                    Set lnodtemp = ctvw收费项目2.Nodes.Add("s" & Mid(lstrtvwme, 2, ((lint级数 - 1) * 3)), tvwChild, lstrtvwme, lrsd收费项目("收费项目名称").Value)
                    lrsd收费项目.MoveNext
                    Set lnodtemp = Nothing
                Loop
            End If
        Next
    End If
    
    '初始化操作标志
    mstr票据信息操作标志 = ""
    mstr收费项目操作标志 = ""
    mstr收费标准操作方式标志 = ""
    mstr打折信息操作标志 = ""
    mstr打折控制操作标志 = ""
    
    '功能：初始化收据号
    '作者：徐冀川
    '时间：2002/07/17
    sub获取当前收据号
    Ctxt收据号值.Enabled = False
    
    frm提示.cprg提示.Value = frm提示.cprg提示.Value + 7
    frm提示.Refresh
    Unload frm提示
WayOut:
    Exit Sub
errHandler:
    If Err.Number = 9999 Then
        MsgBox "请先执行收费管理数据库脚本", vbExclamation, "设置"
        Exit Sub
    Else
        sfsub错误处理 "收费界面", "frm设置", "Form_load", Err.Number, Err.Description, False
    End If
End Sub

'查找票据文件
Private Sub ccmd浏览_Click()
    On Error GoTo errHandler
    Dim lstrfilename As String
    Ccmn票据.InitDir = App.Path
    Ccmn票据.ShowOpen
'    cFlg票据信息.Enabled = False
    lstrfilename = funcgetfilename(Ccmn票据.filename)
    If Len(lstrfilename) > 14 Then
        MsgBox "对不起，文件名称过长，必须少于14个字"
    Else
        cinb票据格式文件名称.Text = lstrfilename
    End If
'    cFlg票据信息.Enabled = True
    Exit Sub
errHandler:
    sfsub错误处理 "收费界面", "frm设置", "Ccmd浏览_Click", Err.Number, Err.Description, False
End Sub

'点击票据信息，信息显示在相应位置上。
Private Sub cFlg票据信息_Click()
    Dim lrsdtemp As Recordset
    Dim lstr编号 As String
    On Error GoTo errHandler
    If cstbShezhi.TabEnabled(1) = False Or mbln退出 Then
        Exit Sub
    End If
    If cFlg票据信息.RowSel = 0 Then
        Exit Sub
    End If
    Set lrsdtemp = pobj业务设置.func提取字典表信息("票据类型字典表")
    cinb票据编号.Text = cFlg票据信息.Cell(flexcpText, cFlg票据信息.RowSel, 1)
    cinb票据名称.Text = cFlg票据信息.Cell(flexcpText, cFlg票据信息.RowSel, 3)
    cinb票据格式文件名称.Text = cFlg票据信息.Cell(flexcpText, cFlg票据信息.RowSel, 4)
    ctxt最大项数.Text = cFlg票据信息.Cell(flexcpText, cFlg票据信息.RowSel, 6)
    If (Not lrsdtemp.EOF) And (Not lrsdtemp.BOF) Then
        lrsdtemp.MoveFirst
       Do While (Not lrsdtemp.EOF)
        lstr编号 = lrsdtemp("innerID").Value
            If lstr编号 = cFlg票据信息.Cell(flexcpText, cFlg票据信息.RowSel, 2) Then
               
               ccmb票据信息票据类型.Text = lrsdtemp("名称")
               
               Exit Do
                'ccmb
            Else
                lrsdtemp.MoveNext
            End If
            If lrsdtemp.EOF Then
                MsgBox "如果修改了票据类型字典表，请重新配置该项目的票据类型,该项目票据类型有问题", vbExclamation, "票据类型设置"
                Exit Do
            End If
       Loop
    End If
    
    If cFlg票据信息.Cell(flexcpText, cFlg票据信息.RowSel, 5) = "一般" Then
        ccmb对应业务.ListIndex = 0
    Else
        ccmb对应业务.ListIndex = 1
    End If
    If mstr票据信息操作标志 = "添加" Or mstr票据信息操作标志 = "修改" Then
        cinb票据名称.SetFocus
    End If
    Set lrsdtemp = Nothing
    Exit Sub
errHandler:
    'If Err.Number = 5 Then Exit Sub
    sfsub错误处理 "收费界面", "frm设置", "cFlg票据信息_Click", Err.Number, Err.Description, False
End Sub

Private Sub cFlg票据信息_KeyPress(KeyAscii As Integer)
    Dim lrsdtemp As Recordset
    On Error GoTo errHandler
    If cstbShezhi.TabEnabled(1) = False Or mbln退出 Then
        Exit Sub
    End If
    If KeyAscii = vbKeyReturn Then
        If cFlg票据信息.RowSel = 0 Then
            Exit Sub
        End If
        Set lrsdtemp = pobj业务设置.func提取字典表信息("票据类型字典表")
        cinb票据编号.Text = cFlg票据信息.Cell(flexcpText, cFlg票据信息.RowSel, 1)
        cinb票据名称.Text = cFlg票据信息.Cell(flexcpText, cFlg票据信息.RowSel, 3)
        cinb票据格式文件名称.Text = cFlg票据信息.Cell(flexcpText, cFlg票据信息.RowSel, 4)
        If (Not lrsdtemp.EOF) And (Not lrsdtemp.BOF) Then
            lrsdtemp.MoveFirst
           Do While (Not lrsdtemp.EOF)
                If CStr(lrsdtemp("innerID")) = cFlg票据信息.Cell(flexcpText, cFlg票据信息.RowSel, 2) Then
                   
                   ccmb票据信息票据类型.Text = lrsdtemp("名称")
                   
                   Exit Do
                    'ccmb
                Else
                    lrsdtemp.MoveNext
                End If
           Loop
        End If
        '
        If cFlg票据信息.Cell(flexcpText, cFlg票据信息.RowSel, 5) = "一般" Then
            ccmb对应业务.ListIndex = 0
        Else
            ccmb对应业务.ListIndex = 1
        End If
        Set lrsdtemp = Nothing
        If mstr票据信息操作标志 = "修改" Or mstr票据信息操作标志 = "添加" Then
            cinb票据名称.SetFocus
        End If
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "收费界面", "frm设置", "cFlg票据信息_KeyPress", Err.Number, Err.Description, False

End Sub

'功能：将票据信息设置状态设置为添加
'输入：无
'处理过程：
'   1、设置录入框和下拉框的使用状态
'   2、将这些录入框和下拉框置为空
'   3、设置工具栏上的按钮的使用状态
'   4、设置票据信息操作方式
'输出：无
'返回：无
'创建人：何嘉
'创建时间：2001-04-04

Private Sub sub票据信息添加()
    On Error GoTo errHandler
    
    '设置使用状态
    cinb票据名称.Enabled = True
    'cinb票据格式文件名称.Enabled = True
    ccmb票据信息票据类型.Enabled = True
    ccmb对应业务.Enabled = True
    Ccmd浏览.Enabled = True
    ctxt最大项数.Enabled = True
    
    '置空录入框和下拉框
    cinb票据名称.Text = ""
    cinb票据格式文件名称.Text = ""
    ctxt最大项数.Text = ""
'    ccmb票据信息票据类型.Text = ""
'    ccmb对应业务.Text = ""
    
    '设置工具栏上的按钮的使用状态
    ctlb设置.Buttons("添加(&A)").Enabled = False
    ctlb设置.Buttons("删除(&D)").Enabled = False
    ctlb设置.Buttons("修改(&U)").Enabled = False
    
    '设置票据信息操作方式
    mstr票据信息操作标志 = "添加"
    
    Exit Sub
errHandler:
    sfsub错误处理 "收费界面", "frm设置", "sub票据信息添加", Err.Number, Err.Description, True
End Sub


'功能：将票据信息设置状态设置为修改
'输入：无
'处理过程：
'   1、设置录入框和下拉框的使用状态
'   2、将这些录入框和下拉框置为空
'   3、设置工具栏上的按钮的使用状态
'   4、设置票据信息操作方式
'输出：无
'返回：无
'创建人：何嘉
'创建时间：2001-04-04

Private Sub sub票据信息修改()
    On Error GoTo errHandler

    '设置使用状态
    cinb票据名称.Enabled = True
    'cinb票据格式文件名称.Enabled = True
    ccmb票据信息票据类型.Enabled = True
    ccmb对应业务.Enabled = True
    Ccmd浏览.Enabled = True
    ctxt最大项数.Enabled = True
    
    '设置工具栏上的按钮的使用状态
    ctlb设置.Buttons("添加(&A)").Enabled = False
    ctlb设置.Buttons("删除(&D)").Enabled = False
    ctlb设置.Buttons("修改(&U)").Enabled = False
    
    '设置票据信息操作方式
    mstr票据信息操作标志 = "修改"
    
    Exit Sub
errHandler:
    sfsub错误处理 "收费界面", "frm设置", "sub票据信息修改", Err.Number, Err.Description, True
End Sub


'功能：将票据信息设置状态设置为删除
'输入：无
'处理过程：
'   1、设置录入框和下拉框的使用状态
'   2、设置工具栏上的按钮的使用状态
'   3、设置票据信息操作方式
'输出：无
'返回：无
'创建人：何嘉
'创建时间：2001-04-04

Private Sub sub票据信息删除()
    On Error GoTo errHandler
    '设置工具栏上的按钮的使用状态
    ctlb设置.Buttons("添加(&A)").Enabled = False
    ctlb设置.Buttons("删除(&D)").Enabled = False
    ctlb设置.Buttons("修改(&U)").Enabled = False
    'cFlg票据信息.Enabled = False
    '设置票据信息操作方式
    mstr票据信息操作标志 = "删除"
    Exit Sub
errHandler:
    sfsub错误处理 "收费界面", "frm设置", "sub票据信息删除", Err.Number, Err.Description, True
End Sub



'功能：清除当前设置，恢复未设置状态
'输入：无
'处理过程：
'   1、设置录入框和下拉框的使用状态
'   2、将这些录入框和下拉框置为空
'   3、设置工具栏上的按钮的使用状态
'   4、设置票据信息操作方式
'输出：无
'返回：无
'创建人：何嘉
'创建时间：2001-04-04

Private Sub sub票据信息清除()
    On Error GoTo errHandler
    
    '置空录入框和下拉框
    cinb票据编号.Text = ""
    cinb票据名称.Text = ""
    cinb票据格式文件名称.Text = ""
    ccmb票据信息票据类型.ListIndex = -1
    ccmb对应业务.ListIndex = -1
    ctxt最大项数.Text = ""
    '设置使用状态
    cinb票据名称.Enabled = False
    cinb票据格式文件名称.Enabled = False
    ccmb票据信息票据类型.Enabled = False
    ccmb对应业务.Enabled = False
    Ccmd浏览.Enabled = False
    ctxt最大项数.Enabled = False
    cFlg票据信息.Enabled = True
    '设置工具栏上的按钮的使用状态
    ctlb设置.Buttons("添加(&A)").Enabled = True
    ctlb设置.Buttons("删除(&D)").Enabled = True
    ctlb设置.Buttons("修改(&U)").Enabled = True
    
    '设置票据信息操作方式
    mstr票据信息操作标志 = ""
    Exit Sub
errHandler:
    sfsub错误处理 "收费界面", "frm设置", "sub票据信息清除", Err.Number, Err.Description, True
End Sub



'功能：保存信息
'输入：无
'处理过程：
'   1、判断那种操作
'   2、判断用户录入是否合理
'   3、合理则保存信息
'   4、更新界面
'输出：无
'返回：无
'创建人：何嘉
'创建时间：2001-04-04

Private Function func票据信息保存() As Boolean
    On Error GoTo errHandler
    Dim linttemp As Integer
    Dim lstrme编号 As String
    Dim lstrtemp As String
    '判断用户操作方式
    Select Case mstr票据信息操作标志
        Case "添加"
            '自动生成编号
            ReDim mstr编号(cFlg票据信息.Rows)
            If cFlg票据信息.Rows > 1 Then
                For linttemp = 1 To cFlg票据信息.Rows - 1
                    mstr编号(linttemp) = cFlg票据信息.Cell(flexcpText, linttemp, 1)
                Next
                
                lstrtemp = mstr编号(1)
                For linttemp = 1 To cFlg票据信息.Rows
                    If lstrtemp < mstr编号(linttemp) Then
                        lstrtemp = mstr编号(linttemp)
                    End If
                Next
                lstrme编号 = Format(CStr(CInt(lstrtemp) + 1), "00")
            Else
                lstrme编号 = "01"
            End If
            
            
            
            
            '判断用户录入是否合理
            If cinb票据格式文件名称.Text = "" Or cinb票据名称.Text = "" Or ccmb对应业务.Text = "" Or ccmb票据信息票据类型.Text = "" Then
                MsgBox "录入不完整，请重新录入", vbExclamation, "票据信息设置"
                func票据信息保存 = False
                Exit Function
            End If
            If ctxt最大项数.Text = "" Then
                MsgBox "请填上最大项数！", vbExclamation, "票据设置"
                func票据信息保存 = False
                Exit Function
            End If
            If Not IsNumeric(ctxt最大项数.Text) Then
                If ctxt最大项数.Text < 1 Then
                    MsgBox "最大项数录入非法！", vbExclamation, "票据设置"
                    func票据信息保存 = False
                    Exit Function
                End If
            End If
            
            '保存信息
            Set mcol票据信息添加 = New Collection
            mcol票据信息添加.Add lstrme编号, "编号"
            mcol票据信息添加.Add cinb票据名称.Text, "票据名称"
            mcol票据信息添加.Add cinb票据格式文件名称.Text, "票据格式文件名称"
            mcol票据信息添加.Add ccmb票据信息票据类型.ItemData(ccmb票据信息票据类型.ListIndex), "票据类型编号"
            mcol票据信息添加.Add ccmb对应业务.Text, "对应业务"
            mcol票据信息添加.Add ctxt最大项数.Text, "最大项数"
            
            '判断该记录是否存在
            If cFlg票据信息.Rows > 0 Then
                For linttemp = 1 To cFlg票据信息.Rows - 1
                    'cFlg票据信息.Cell(flexcpText, linttemp, 3) = cinb票据名称.Text And
                    If cFlg票据信息.Cell(flexcpText, linttemp, 2) = ccmb票据信息票据类型.ItemData(ccmb票据信息票据类型.ListIndex) And _
                    cFlg票据信息.Cell(flexcpText, linttemp, 4) = cinb票据格式文件名称.Text And _
                    cFlg票据信息.Cell(flexcpText, linttemp, 5) = ccmb对应业务.Text And _
                    cFlg票据信息.Cell(flexcpText, linttemp, 6) = ctxt最大项数.Text Then
                        MsgBox "该票据设置信息已存在，请修改", vbExclamation, "票据设置"
                        func票据信息保存 = False
                        Exit Function
                    End If
                Next
                cFlg票据信息.AutoSize 0, cFlg票据信息.Cols - 1
            End If
            
            If pobj业务设置.func添加票据信息(mcol票据信息添加) Then
                           
                '界面上添加一条记录
                cFlg票据信息.AddItem ""
                cFlg票据信息.Cell(flexcpText, cFlg票据信息.Rows - 1, 1) = lstrme编号
                cFlg票据信息.Cell(flexcpText, cFlg票据信息.Rows - 1, 2) = ccmb票据信息票据类型.ItemData(ccmb票据信息票据类型.ListIndex)
                cFlg票据信息.Cell(flexcpText, cFlg票据信息.Rows - 1, 3) = cinb票据名称.Text
                cFlg票据信息.Cell(flexcpText, cFlg票据信息.Rows - 1, 4) = cinb票据格式文件名称.Text
                cFlg票据信息.Cell(flexcpText, cFlg票据信息.Rows - 1, 5) = ccmb对应业务.Text
                cFlg票据信息.Cell(flexcpText, cFlg票据信息.Rows - 1, 6) = ctxt最大项数.Text
                cFlg票据信息.AutoSize 0, cFlg票据信息.Cols - 1
            Else
                func票据信息保存 = False
                Exit Function
            End If
            Set mcol票据信息添加 = Nothing
        Case "删除"
                        
            '保存信息
            If cinb票据编号.Text = "" Then Exit Function
            
            If pobj业务设置.func删除票据信息(cinb票据编号.Text) Then
            
                '界面删除一条记录
                cFlg票据信息.RemoveItem (cFlg票据信息.RowSel)
            Else
                func票据信息保存 = False
                Exit Function
            End If
        Case "修改"
            '判断用户录入是否合理
            If cinb票据编号.Text = "" Then
                MsgBox "请选择将要修改的票据设置信息", vbExclamation, "票据信息设置"
                func票据信息保存 = False
                Exit Function
            End If
            If cinb票据格式文件名称.Text = "" Or cinb票据名称.Text = "" Or ccmb对应业务.Text = "" Or ccmb票据信息票据类型.Text = "" Then
                MsgBox "录入不完整，请重新录入", vbExclamation, "票据信息设置"
                func票据信息保存 = False
                Exit Function
            End If
            If ctxt最大项数.Text = "" Then
                MsgBox "请填上最大项数！", vbExclamation, "票据设置"
                func票据信息保存 = False
                Exit Function
            End If
            If Not IsNumeric(ctxt最大项数.Text) Then
                If ctxt最大项数.Text < 1 Then
                    MsgBox "最大项数录入非法！", vbExclamation, "票据设置"
                    func票据信息保存 = False
                    Exit Function
                End If
            End If
            '保存信息
            Set mcol票据信息修改 = New Collection
            mcol票据信息修改.Add cinb票据编号.Text, "编号"
            mcol票据信息修改.Add cinb票据名称.Text, "票据名称"
            mcol票据信息修改.Add cinb票据格式文件名称.Text, "票据格式文件名称"
            mcol票据信息修改.Add ccmb票据信息票据类型.ItemData(ccmb票据信息票据类型.ListIndex), "票据类型编号"
            mcol票据信息修改.Add ccmb对应业务.Text, "对应业务"
            mcol票据信息修改.Add ctxt最大项数.Text, "最大项数"
            If pobj业务设置.func修改票据信息(mcol票据信息修改) Then
                
                '界面上更新一条记录
                cFlg票据信息.Cell(flexcpText, cFlg票据信息.RowSel, 1) = cinb票据编号.Text
                cFlg票据信息.Cell(flexcpText, cFlg票据信息.RowSel, 2) = ccmb票据信息票据类型.ItemData(ccmb票据信息票据类型.ListIndex)
                cFlg票据信息.Cell(flexcpText, cFlg票据信息.RowSel, 3) = cinb票据名称.Text
                cFlg票据信息.Cell(flexcpText, cFlg票据信息.RowSel, 4) = cinb票据格式文件名称.Text
                cFlg票据信息.Cell(flexcpText, cFlg票据信息.RowSel, 5) = ccmb对应业务.Text
                cFlg票据信息.Cell(flexcpText, cFlg票据信息.RowSel, 6) = ctxt最大项数.Text
                cFlg票据信息.AutoSize 0, cFlg票据信息.Cols - 1
            Else
                func票据信息保存 = False
                Exit Function
            End If
            Set mcol票据信息修改 = Nothing
                        
        Case Else
            
    End Select
    
    '更新界面及操作标志
    sub票据信息清除
    func票据信息保存 = True
    Exit Function
errHandler:
    func票据信息保存 = False
    sfsub错误处理 "收费界面", "frm设置", "sub票据信息保存", Err.Number, Err.Description, True
End Function

'功能：将设置状态置为添加状态
'输入：无
'处理过程：
'   1、设置录入框和下拉框的使用状态
'   2、将这些录入框和下拉框置为空
'   3、设置工具栏上的按钮的使用状态
'   4、设置操作方式
'输出：无
'返回：无
'创建人：何嘉
'创建时间：2001-04-04
Private Sub sub收费标准添加()
    On Error GoTo errHandler

    '设置录入框和下拉框的使用状态
    cinb收费标准名称.Enabled = True
    cinb助记符.Enabled = True
    Ccmb使用范围.Enabled = True
    Ccmb子系统名.Enabled = True
    cflg收费标准.Editable = True
    cflg收费标准.Rows = 1
    Cchk收费标准列表.Value = 0
    Cchk收费标准列表.Enabled = False
    clst收费标准名称.Visible = False
    Ccmd添加.Enabled = True
    Ccmd移除.Enabled = True
    
    '将这些录入框和下拉框置为空
    cinb收费标准名称.Text = ""
    cinb助记符.Text = ""
'    Ccmb使用范围.Text = ""
'    Ccmb子系统名.Text = ""
    
    '设置工具栏上的按钮的使用状态
    ctlb设置.Buttons("添加(&A)").Enabled = False
    ctlb设置.Buttons("修改(&U)").Enabled = False
    ctlb设置.Buttons("删除(&D)").Enabled = False
    
    '设置票据信息操作方式
    mstr收费标准操作方式标志 = "添加"
    
    Exit Sub
errHandler:
    sfsub错误处理 "收费界面", "frm设置", "sub收费标准添加", Err.Number, Err.Description, True
End Sub

'功能：将设置状态设置为修改
'输入：无
'处理过程：
'   1、设置录入框和下拉框的使用状态
'   2、设置工具栏上的按钮的使用状态
'   3、设置操作方式
'输出：无
'返回：无
'创建人：何嘉
'创建时间：2001-04-04
Private Sub sub收费标准修改()
    Dim lrsd收费标准名称 As Recordset
    On Error GoTo errHandler
'    Set lrsd收费标准名称 = pobj业务设置.func查询收费标准(, "select 收费标准名称 from 收费管理_收费标准信息表")
'    If (Not lrsd收费标准名称.BOF) And (Not lrsd收费标准名称.EOF) Then
'        clst收费标准名称.Clear
'        lrsd收费标准名称.MoveFirst
'        Do While (Not lrsd收费标准名称.EOF)
'            clst收费标准名称.AddItem lrsd收费标准名称("收费标准名称").Value
'            lrsd收费标准名称.MoveNext
'        Loop
'    End If
'    cinb收费标准名称.Enabled = True
    cinb助记符.Enabled = True
    Ccmb使用范围.Enabled = True
    Ccmb子系统名.Enabled = True
    cflg收费标准.Editable = True
    Ccmd添加.Enabled = True
    Ccmd移除.Enabled = True
    Cchk收费标准列表.Enabled = True
    
    '设置工具栏上的按钮的使用状态
    ctlb设置.Buttons("添加(&A)").Enabled = False
    ctlb设置.Buttons("修改(&U)").Enabled = False
    ctlb设置.Buttons("删除(&D)").Enabled = False
    
    '设置票据信息操作方式
    mstr收费标准操作方式标志 = "修改"
    Exit Sub
errHandler:
    sfsub错误处理 "收费界面", "frm设置", "sub收费标准修改", Err.Number, Err.Description, True
End Sub

'功能：将设置状态设置为删除
'输入：无
'处理过程：
'   1、设置操作方式
'输出：无
'返回：无
'创建人：何嘉
'创建时间：2001-04-04
Private Sub sub收费标准删除()
    On Error GoTo errHandler
     '设置工具栏上的按钮的使用状态
    ctlb设置.Buttons("添加(&A)").Enabled = False
    ctlb设置.Buttons("删除(&D)").Enabled = False
    ctlb设置.Buttons("修改(&U)").Enabled = False
    cFlg票据信息.Editable = False
    Cchk收费标准列表.Enabled = True
    '设置操作方式
    mstr收费标准操作方式标志 = "删除"
    Exit Sub
errHandler:
    sfsub错误处理 "收费界面", "frm设置", "sub收费标准删除", Err.Number, Err.Description, True
End Sub

'功能：清除当前设置，恢复未设置状态
'输入：无
'处理过程：
'   1、设置录入框和下拉框的使用状态
'   2、将这些录入框和下拉框置为空
'   3、设置工具栏上的按钮的使用状态
'   4、设置操作方式
'输出：无
'返回：无
'创建人：何嘉
'创建时间：2001-04-04
Private Sub sub收费标准清除()
    On Error GoTo errHandler
      '设置票据信息操作方式
    mstr收费标准操作方式标志 = ""
       '设置录入框和下拉框的使用状态
    cinb收费标准名称.Enabled = False
    cinb助记符.Enabled = False
    Ccmb使用范围.ListIndex = -1
    Ccmb使用范围.Enabled = False
    Ccmb子系统名.ListIndex = -1
    Ccmb子系统名.Enabled = False
    cflg收费标准.Editable = False
    cflg收费标准.Clear
    cflg收费标准.Rows = 1
    cflg收费标准.Cell(flexcpText, 0, 1) = "编号"
    cflg收费标准.Cell(flexcpText, 0, 2) = "收费项目编号"
    cflg收费标准.Cell(flexcpText, 0, 3) = "收费项目名称"
    cflg收费标准.Cell(flexcpText, 0, 4) = "单价"
    cflg收费标准.Cell(flexcpText, 0, 5) = "数量"
    cflg收费标准.AutoSize 0, cflg收费标准.Cols - 1
    'cflg收费标准.Cell(flexcpText, 0, 6) = "允许修改数量"
    Cchk收费标准列表.Value = 0
    Cchk收费标准列表.Enabled = False
    clst收费标准名称.Visible = False
    Ccmd添加.Enabled = False
    Ccmd移除.Enabled = False
    '将这些录入框和下拉框置为空
    cinb收费标准名称.Text = ""
    cinb助记符.Text = ""
'    Ccmb使用范围.Text = ""
'    Ccmb子系统名.Text = ""
    Ccmb使用范围.Clear
    '设置工具栏上的按钮的使用状态
    ctlb设置.Buttons("添加(&A)").Enabled = True
    ctlb设置.Buttons("修改(&U)").Enabled = True
    ctlb设置.Buttons("删除(&D)").Enabled = True
    
  
    
    
    
    
    Exit Sub
errHandler:
    sfsub错误处理 "收费界面", "frm设置", "sub收费标准清除", Err.Number, Err.Description, True
End Sub

'功能：保存当前设置，恢复未设置状态
'输入：无
'处理过程：
'   1、判断那种操作
'   2、判断用户录入是否合理
'   3、合理则保存信息
'   4、更新界面
'输出：无
'返回：无
'创建人：何嘉
'创建时间：2001-04-04
Private Function func收费标准保存() As Boolean
    On Error GoTo errHandler
    Dim lintlistindex As Integer
    Dim lcol收费标准 As Collection
    Dim lintcountitem As Integer
    Dim linttemp As Integer
    Dim lrsdtemp As Recordset
    Dim lintprecountitem As Integer
    Dim linttemp编号() As Integer
    Dim lrsd收费标准名称 As Recordset
    Dim lstr最小单价 As String
    Dim lstr最大单价 As String
    Dim lstr助记符 As String

    If cflg收费标准.Rows = 1 Then
        MsgBox "对不起，请选择收费项目！", vbOKOnly & vbExclamation, "系统提示"
        func收费标准保存 = False
        Exit Function
    End If
    '判断用户的操作方式
    Select Case mstr收费标准操作方式标志
        
        Case "添加"
            '判断用户录入是否合理
            If cinb收费标准名称.Text <> "" Then
                If clst收费标准名称.ListCount <> -1 Then
                    For lintlistindex = 0 To clst收费标准名称.ListCount
                        If LTrim(RTrim(cinb收费标准名称.Text)) = clst收费标准名称.List(lintlistindex) Then
                            MsgBox "不能重名", vbExclamation, "收费标准设置"
                            func收费标准保存 = False
                            Exit Function
                        End If
                    Next
                Else
                End If
                If Ccmb子系统名.Text = "" Then
                    mrsd子系统条件.Filter = "系统名='收费管理'"
                    ctxt子系统编号.Text = mrsd子系统条件("编号")
                    mrsd子系统条件.Filter = ""
                Else
                    If Ccmb使用范围 = "" Then
                        MsgBox "请选择使用范围", vbExclamation, "收费标准设置"
                        func收费标准保存 = False
                        Exit Function
                    End If
                End If
                
            Else
                MsgBox "收费标准名称不能为空", vbExclamation, "收费标准设置"
                func收费标准保存 = False
                Exit Function
            End If
            lintcountitem = cflg收费标准.Rows - 1
            lstr助记符 = cinb助记符.Text
            cflg收费标准.Select 0, 0
            cinb助记符.Text = ""
            cinb助记符.Text = lstr助记符
            For linttemp = 1 To lintcountitem
                cflg收费标准.Refresh
                If cflg收费标准.Cell(flexcpText, linttemp, 4) = "" Then
                    MsgBox "单价录入不能为空，请检查后修改", vbExclamation, "收费标准设置"
                    func收费标准保存 = False
                    Exit Function
                Else
                    If IsNumeric(cflg收费标准.Cell(flexcpText, linttemp, 4)) Then
                    Set lrsdtemp = pobj收费管理.func查询收费项目("收费项目编号='" & cflg收费标准.Cell(flexcpText, linttemp, 2) & "'")
            
                    If (Not lrsdtemp.BOF) And (Not lrsdtemp.EOF) Then
            
                        lstr最大单价 = lrsdtemp("最大单价").Value
                        lstr最小单价 = lrsdtemp("最小单价").Value
                        cflg收费标准.Refresh
            
                        If CDbl(cflg收费标准.Cell(flexcpText, linttemp, 4)) < lstr最小单价 Or CDbl(cflg收费标准.Cell(flexcpText, linttemp, 4)) > CDbl(lstr最大单价) Then
                            MsgBox "单价必须在最小单价和最大单价，请检查后修改", vbExclamation, "收费标准设置"
                            func收费标准保存 = False
                            Exit Function
                        End If
                                             
                    End If
                    Set lrsdtemp = Nothing
                    Else
                        MsgBox "单价必须为数字，请检查后修改", vbExclamation, "收费标准设置"
                        func收费标准保存 = False
                        Exit Function
                    End If
                End If
                If cflg收费标准.Cell(flexcpText, linttemp, 5) = "" Then
                    MsgBox "数量录入不能为空，请检查后修改", vbExclamation, "收费标准设置"
                    func收费标准保存 = False
                    Exit Function
                Else
                    If Not IsNumeric(cflg收费标准.Cell(flexcpText, linttemp, 5)) Then
                        MsgBox "数量必须为数字，请检查后修改", vbExclamation, "收费标准设置"
                        func收费标准保存 = False
                        Exit Function
                    End If
                End If
            Next
            dasubBeginTran
            For linttemp = 1 To lintcountitem
                Set lcol收费标准 = New Collection
                lcol收费标准.Add cinb收费标准名称.Text, "收费标准名称"
                lcol收费标准.Add ctxt子系统编号.Text, "过滤条件"
                lcol收费标准.Add cinb助记符.Text, "助记符"
                lcol收费标准.Add cflg收费标准.Cell(flexcpText, linttemp, 2), "收费项目编号"
                lcol收费标准.Add cflg收费标准.Cell(flexcpText, linttemp, 4), "单价"
                lcol收费标准.Add cflg收费标准.Cell(flexcpText, linttemp, 5), "数量"
                lcol收费标准.Add 1, "允许修改数量"
                'Debug.Print lcol收费标准("允许修改数量")
                If pobj业务设置.func添加收费标准(lcol收费标准) Then
                    Set lcol收费标准 = Nothing
                Else
                    Err.Raise 33000
                End If
            Next
            dasubCommitTran
            clst收费标准名称.AddItem cinb收费标准名称.Text
            
        Case "修改"
            '判断用户录入是否合理
            If cinb收费标准名称.Text <> "" Then
                If clst收费标准名称.ListCount <> -1 Then
'                    For lintlistindex = 0 To clst收费标准名称.ListCount
'                        If LTrim(RTrim(clst收费标准名称)) = clst收费标准名称.List(lintlistindex) Then
'                            MsgBox "不能重名", vbExclamation, "收费标准设置"
'                            Exit Sub
'                        End If
'                    Next
                Else
                End If
                If Ccmb子系统名.Text = "" Then
                    mrsd子系统条件.Filter = "系统名='收费管理'"
                    ctxt子系统编号.Text = mrsd子系统条件("编号")
                    mrsd子系统条件.Filter = ""
                Else
                    If Ccmb使用范围 = "" Then
                        MsgBox "请选择使用范围", vbExclamation, "收费标准设置"
                        func收费标准保存 = False
                        Exit Function
                    End If
                End If
                
            Else
                MsgBox "收费标准名称不能为空", vbExclamation, "收费标准设置"
                func收费标准保存 = False
                Exit Function
            End If
            cflg收费标准.Refresh
            lintcountitem = cflg收费标准.Rows - 1
            lstr助记符 = cinb助记符.Text
            cflg收费标准.Select 0, 0
            cinb助记符.Text = ""
            cinb助记符.Text = lstr助记符
            For linttemp = 1 To lintcountitem
                cflg收费标准.Refresh
                If cflg收费标准.Cell(flexcpText, linttemp, 4) = "" Then
                    MsgBox "单价录入不能为空，请检查后修改", vbExclamation, "收费标准设置"
                    func收费标准保存 = False
                    Exit Function
                Else
                    If IsNumeric(cflg收费标准.Cell(flexcpText, linttemp, 4)) Then
                    Set lrsdtemp = pobj收费管理.func查询收费项目("收费项目编号='" & cflg收费标准.Cell(flexcpText, linttemp, 2) & "'")
                    If (Not lrsdtemp.BOF) And (Not lrsdtemp.EOF) Then
                        lstr最大单价 = lrsdtemp("最大单价").Value
                        lstr最小单价 = lrsdtemp("最小单价").Value
                        If CDbl(cflg收费标准.Cell(flexcpText, linttemp, 4)) < lstr最小单价 Or CDbl(cflg收费标准.Cell(flexcpText, linttemp, 4)) > CDbl(lstr最大单价) Then
                            MsgBox "单价必须在最小单价和最大单价，请检查后修改", vbExclamation, "收费标准设置"
                            func收费标准保存 = False
                            Exit Function
                        End If
                    End If
                    Set lrsdtemp = Nothing
                    Else
                        MsgBox "单价必须为数字，请检查后修改", vbExclamation, "收费标准设置"
                        func收费标准保存 = False
                        Exit Function
                    End If
                End If
                If cflg收费标准.Cell(flexcpText, linttemp, 5) = "" Then
                    MsgBox "数量录入不能为空，请检查后修改", vbExclamation, "收费标准设置"
                    func收费标准保存 = False
                    Exit Function
                Else
                    If Not IsNumeric(cflg收费标准.Cell(flexcpText, linttemp, 5)) Then
                        MsgBox "数量必须为数字，请检查后修改", vbExclamation, "收费标准设置"
                        func收费标准保存 = False
                        Exit Function
                    End If
                End If
            Next
            
            dasubBeginTran
            '删除原来的信息
            Set lrsd收费标准名称 = pobj业务设置.func查询收费标准("收费标准名称='" & cinb收费标准名称.Text & "'")
            If (Not lrsd收费标准名称.BOF) And (Not lrsd收费标准名称.EOF) Then
                lrsd收费标准名称.MoveFirst
                Do While (Not lrsd收费标准名称.EOF)
                    Dim lstr编号 As String
                    lstr编号 = lrsd收费标准名称("编号").Value
                    If pobj业务设置.func删除收费标准(lstr编号) Then
                        lrsd收费标准名称.MoveNext
                    Else
                        Set lrsd收费标准名称 = Nothing
                        Err.Raise 33001
                    End If
                Loop
                Set lrsd收费标准名称 = Nothing
            End If
            
            '添加信息
            For linttemp = 1 To lintcountitem
                Set lcol收费标准 = New Collection
                lcol收费标准.Add cinb收费标准名称.Text, "收费标准名称"
                lcol收费标准.Add ctxt子系统编号.Text, "过滤条件"
                lcol收费标准.Add cinb助记符.Text, "助记符"
                lcol收费标准.Add cflg收费标准.Cell(flexcpText, linttemp, 2), "收费项目编号"
                lcol收费标准.Add cflg收费标准.Cell(flexcpText, linttemp, 4), "单价"
                lcol收费标准.Add cflg收费标准.Cell(flexcpText, linttemp, 5), "数量"
                
                lcol收费标准.Add 1, "允许修改数量"
                If pobj业务设置.func添加收费标准(lcol收费标准) Then
                    Set lcol收费标准 = Nothing
                Else
                    Err.Raise 33001
                End If
            Next
            dasubCommitTran
            '处理界面
            
            Set lrsd收费标准名称 = pobj业务设置.func查询收费标准(, "select 收费标准名称 from 收费管理_收费标准信息表 group by 收费标准名称")
            If (Not lrsd收费标准名称.BOF) And (Not lrsd收费标准名称.EOF) Then
                clst收费标准名称.Clear
                lrsd收费标准名称.MoveFirst
                Do While (Not lrsd收费标准名称.EOF)
                    clst收费标准名称.AddItem lrsd收费标准名称("收费标准名称").Value
                    lrsd收费标准名称.MoveNext
                Loop
            End If
            
            
            
        Case "删除"
             '删除原来的信息
            If cflg收费标准.Rows > 1 Then
                lintprecountitem = cflg收费标准.Rows - 1
                ReDim linttemp编号(lintprecountitem)
                linttemp = 1
                If lintprecountitem > 0 Then
                    
                    For linttemp = 1 To lintprecountitem
                        linttemp编号(linttemp) = cflg收费标准.Cell(flexcpText, linttemp, 1)
                    Next
                    For linttemp = 1 To lintprecountitem
                        pobj业务设置.func删除收费标准 linttemp编号(linttemp)
                    Next
    
                End If
            End If
                      
            '处理界面
            
            Set lrsd收费标准名称 = pobj业务设置.func查询收费标准(, "select 收费标准名称 from 收费管理_收费标准信息表 group by 收费标准名称")
            If (Not lrsd收费标准名称.BOF) And (Not lrsd收费标准名称.EOF) Then
                clst收费标准名称.Clear
                lrsd收费标准名称.MoveFirst
                Do While (Not lrsd收费标准名称.EOF)
                    clst收费标准名称.AddItem lrsd收费标准名称("收费标准名称").Value
                    lrsd收费标准名称.MoveNext
                Loop
            End If
               
    End Select
    sub收费标准清除
    func收费标准保存 = True
    Exit Function
errHandler:
    func收费标准保存 = False
    If Err.Number = 33000 Then
        MsgBox "添加失败，请检查数据项是否合法，或是否发生网络故障，或是否数据库使用正常", vbExclamation, "收费标准设置"
        sub收费标准清除
        Exit Function
    End If
    If Err.Number = 33001 Then
        MsgBox "修改失败，请检查数据项是否合法，或是否发生网络故障，或是否数据库使用正常", vbExclamation, "收费标准设置"
        sub收费标准清除
        Exit Function
    End If
    sfsub错误处理 "收费界面", "frm设置", "func收费标准保存", Err.Number, Err.Description, True
End Function

'功能：将设置状态置为添加状态
'输入：无
'处理过程：
'   1、设置录入框的使用状态
'   2、将这些录入框置为空
'   3、设置工具栏上的按钮的使用状态
'   4、设置操作方式
'输出：无
'返回：无
'创建人：何嘉
'创建时间：2001-04-05
Private Sub sub打折信息添加()
    On Error GoTo errHandler
    '设置录入框的使用状态
    ctxt打折.Enabled = True
    ccmd单位定位.Enabled = True
    cvsl打折.Enabled = True
    
    '将这些录入框置为空
    cinb单位编号.Text = ""
    ctxt打折.Text = ""
    
    '设置工具栏上的按钮的使用状态
    ctlb设置.Buttons("添加(&A)").Enabled = False
    ctlb设置.Buttons("删除(&D)").Enabled = False
    ctlb设置.Buttons("修改(&U)").Enabled = False
    
    '设置操作方式
    mstr打折信息操作标志 = "添加"
    

    Exit Sub
errHandler:
    sfsub错误处理 "收费界面", "frm设置", "sub打折信息添加", Err.Number, Err.Description, True
End Sub

'功能：将设置状态设置为修改
'输入：无
'处理过程：
'   1、设置录入框的使用状态
'   2、设置工具栏上的按钮的使用状态
'   3、设置操作方式
'输出：无
'返回：无
'创建人：何嘉
'创建时间：2001-04-05
Private Sub sub打折信息修改()
    On Error GoTo errHandler
        '设置录入框的使用状态
    ctxt打折.Enabled = True
    ccmd单位定位.Enabled = False
    cvsl打折.Enabled = True
    
    '设置工具栏上的按钮的使用状态
     ctlb设置.Buttons("添加(&A)").Enabled = False
    ctlb设置.Buttons("删除(&D)").Enabled = False
    ctlb设置.Buttons("修改(&U)").Enabled = False
    '设置操作方式
    mstr打折信息操作标志 = "修改"
    

    Exit Sub
errHandler:
    sfsub错误处理 "收费界面", "frm设置", "sub打折信息修改", Err.Number, Err.Description, True
End Sub

'功能：将设置状态设置为删除
'输入：无
'处理过程：
'   1、设置操作方式
'   2、设置按钮状态
'输出：无
'返回：无
'创建人：何嘉
'创建时间：2001-04-05
Private Sub sub打折信息删除()
    On Error GoTo errHandler

    '设置工具栏上的按钮的使用状态
    ctlb设置.Buttons("添加(&A)").Enabled = False
    ctlb设置.Buttons("删除(&D)").Enabled = False
    ctlb设置.Buttons("修改(&U)").Enabled = False
    cFlg打折.Enabled = True
    '设置操作方式
    mstr打折信息操作标志 = "删除"

    Exit Sub
errHandler:
    sfsub错误处理 "收费界面", "frm设置", "sub打折信息删除", Err.Number, Err.Description, True
End Sub

'功能：清除当前设置，恢复未设置状态
'输入：无
'处理过程：
'   1、设置录入框的使用状态
'   2、将这些录入框置为空
'   3、设置工具栏上的按钮的使用状态
'   4、设置操作方式
'输出：无
'返回：无
'创建人：何嘉
'创建时间：2001-04-05
Private Sub sub打折信息清除()
    On Error GoTo errHandler
    '设置录入框的使用状态
    ctxt打折.Enabled = False
    ccmd单位定位.Enabled = False
    cvsl打折.Enabled = False
    '将这些录入框置为空
    cinb单位编号.Text = ""
    ctxt打折.Text = ""
    ' 清除交费单位名称 时间：2002/01/029 徐冀川
    cinp交费单位.Text = ""
    '设置操作方式
    mstr打折信息操作标志 = ""
    
    ctlb设置.Buttons("添加(&A)").Enabled = True
    ctlb设置.Buttons("修改(&U)").Enabled = True
    ctlb设置.Buttons("删除(&D)").Enabled = True
    
    '恢复打折控制
    copt打折(pint打折控制).Value = True
    
    mstr打折控制操作标志 = ""
    
    '设置工具栏上的按钮的使用状态

    cFlg打折.Enabled = True
    Exit Sub
errHandler:
    sfsub错误处理 "收费界面", "frm设置", "sub打折信息清除", Err.Number, Err.Description, True
End Sub

'功能：保存当前设置，恢复未设置状态
'输入：无
'处理过程：
'   1、判断那种操作
'   2、判断用户录入是否合理
'   3、合理则保存信息
'   4、更新界面
'输出：无
'返回：无
'创建人：何嘉
'创建时间：2001-04-05
Private Function func打折信息保存() As Boolean
    Dim lstr单位编号 As String
    Dim lint单位编号 As Integer
    Dim linttemp As Integer
    On Error GoTo errHandler
    
    '判断那种操作
    Select Case mstr打折信息操作标志
        Case "添加"
            Set mcol打折信息添加 = New Collection
            
            lstr单位编号 = LTrim(RTrim(cinb单位编号.Text)) '(RTrim(cinb单位编号.Text)), "单位编号"
            lint单位编号 = cFlg打折.Rows - 1
            For linttemp = 1 To lint单位编号
                If lstr单位编号 = cFlg打折.Cell(flexcpText, linttemp, 1) Then
                    MsgBox "该单位已进行了修改，如有变动请修改", vbExclamation, "打折设置"
                    func打折信息保存 = False
                    Exit Function
                End If
            Next
            If lstr单位编号 = "" Then
                sffuncMsg "单位信息输入有错，请重新单位定位选择单位!"
                func打折信息保存 = False
            Exit Function
            End If
            
            If ctxt打折.Text <> "" Then
                If IsNumeric(ctxt打折.Text) Then
                    If ctxt打折.Text <= 0 Or ctxt打折.Text > 1 Then
                        MsgBox "打折比率必须在0-1之间！", vbExclamation, "打折设置"
                        func打折信息保存 = False
                        Exit Function
                    Else
                        If (Len(ctxt打折.Text) - InStr(1, ctxt打折.Text, ".", vbTextCompare)) > 2 Then
                            MsgBox "打折比率只能有两位小数，请修改！", vbExclamation, "打折设置"
                            func打折信息保存 = False
                            Exit Function
                        End If
                    End If
                Else
                    MsgBox "请输入数字", vbExclamation, "打折设置"
                    ctxt打折.Text = ""
                    func打折信息保存 = False
                    Exit Function
                End If
            Else
                MsgBox "请录入打折信息", vbExclamation, "打折设置"
                func打折信息保存 = False
                Exit Function
            End If
            '功能：增加对空值的处理，时间：2002/01/29，徐冀川
            mcol打折信息添加.Add IIf(IsNull(cinb单位编号.Text), "", cinb单位编号.Text), "单位编号"
            mcol打折信息添加.Add IIf(IsNull(ctxt打折.Text), "", ctxt打折.Text), "打折比率"
            If pobj业务设置.func添加打折信息(mcol打折信息添加) Then
            Else
                MsgBox "不能添加打折信息，请检查信息录入是否正常，如还有问题，请与系统管理员联系", vbExclamation, "打折设置"
                func打折信息保存 = False
                Exit Function
            End If
            
            Set mcol打折信息添加 = Nothing
            cFlg打折.AddItem "" ', cFlg打折.Rows + 1
            cFlg打折.Cell(flexcpText, cFlg打折.Rows - 1, 1) = cinb单位编号.Text
            cFlg打折.Cell(flexcpText, cFlg打折.Rows - 1, 2) = cinp交费单位.Text
            cFlg打折.Cell(flexcpText, cFlg打折.Rows - 1, 3) = ctxt打折.Text
            cFlg打折.AutoSize 0, cFlg打折.Cols - 1
        Case "修改"
            If ctxt打折.Text <> "" And cinp交费单位.Text <> "" Then
                If IsNumeric(ctxt打折.Text) Then
                    If ctxt打折.Text <= 0 Or ctxt打折.Text > 1 Then
                        MsgBox "打折比率必须在0-1之间！", vbExclamation, "打折设置"
                        func打折信息保存 = False
                        Exit Function
                    Else
                         '增加对两位小数的限制 2002/01/29 徐冀川
                         If (Len(ctxt打折.Text) - InStr(1, ctxt打折.Text, ".", vbTextCompare)) > 2 Then
                            MsgBox "打折比率只能有两位小数，请修改！", vbExclamation, "打折设置"
                            func打折信息保存 = False
                            Exit Function
                        End If
                    End If
                Else
                        MsgBox "请输入数字", vbExclamation, "打折设置"
                        ctxt打折.Text = ""
                        func打折信息保存 = False
                        Exit Function
                End If
            Else
                MsgBox "请录入打折信息", vbExclamation, "打折设置"
                func打折信息保存 = False
                Exit Function
            End If
            
            Set mcol打折信息修改 = New Collection
            mcol打折信息修改.Add LTrim(RTrim(cinb单位编号.Text)), "单位编号"
            mcol打折信息修改.Add ctxt打折.Text, "打折比率"
            pobj业务设置.func修改打折信息 mcol打折信息修改
            Set mcol打折信息修改 = Nothing
            cFlg打折.Cell(flexcpText, cFlg打折.RowSel, 1) = cinb单位编号.Text
            cFlg打折.Cell(flexcpText, cFlg打折.RowSel, 2) = cinp交费单位.Text
            cFlg打折.Cell(flexcpText, cFlg打折.RowSel, 3) = ctxt打折.Text
            cFlg打折.AutoSize 0, cFlg打折.Cols - 1
        Case "删除"
            If cinb单位编号.Text <> "" Then
                If pobj业务设置.func删除打折信息(LTrim(RTrim(cinb单位编号.Text))) Then
                    cFlg打折.RemoveItem cFlg打折.RowSel
                Else
                    func打折信息保存 = False
                    Exit Function
                End If
            Else
                MsgBox "请选择具体的单位", vbExclamation, "打折设置"
                func打折信息保存 = False
                Exit Function
            End If
        Case Else
            func打折信息保存 = True
        
    End Select
    func打折信息保存 = True
    Exit Function
errHandler:
    ctxt打折.Text = ""
    cinb单位编号.Text = ""
    func打折信息保存 = False
    sfsub错误处理 "收费界面", "设置", "打折保存", Err.Number, Err.Description, True
End Function

'点击选中的项目，信息显示在相应的位置
Private Sub cFlg打折_Click()
    On Error GoTo errHandler
    If cstbShezhi.TabEnabled(3) = False Or mbln退出 Then
        Exit Sub
    End If
    If cFlg打折.RowSel > 0 Then
        '处理空值,时间2002/01/29,徐冀川
        cinb单位编号.Text = IIf(IsNull(cFlg打折.Cell(flexcpText, cFlg打折.RowSel, 1)), "", cFlg打折.Cell(flexcpText, cFlg打折.RowSel, 1))
        cinp交费单位.Text = IIf(IsNull(cFlg打折.Cell(flexcpText, cFlg打折.RowSel, 2)), "", cFlg打折.Cell(flexcpText, cFlg打折.RowSel, 2))
        ctxt打折.Text = IIf(IsNull(cFlg打折.Cell(flexcpText, cFlg打折.RowSel, 3)), "", cFlg打折.Cell(flexcpText, cFlg打折.RowSel, 3))
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "收费界面", "frm设置", "cFlg打折_Click", Err.Number, Err.Description, False
End Sub

Private Sub cFlg打折_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandler
    If cstbShezhi.TabEnabled(3) = False Or mbln退出 Then
        Exit Sub
    End If
    If (cFlg打折.RowSel > 0) And (KeyAscii = vbKeyReturn) Then
        '处理空值,时间2002/01/29,徐冀川
        cinb单位编号.Text = IIf(IsNull(cFlg打折.Cell(flexcpText, cFlg打折.RowSel, 1)), "", cFlg打折.Cell(flexcpText, cFlg打折.RowSel, 1))
        cinp交费单位.Text = IIf(IsNull(cFlg打折.Cell(flexcpText, cFlg打折.RowSel, 2)), "", cFlg打折.Cell(flexcpText, cFlg打折.RowSel, 2))
        ctxt打折.Text = IIf(IsNull(cFlg打折.Cell(flexcpText, cFlg打折.RowSel, 3)), "", cFlg打折.Cell(flexcpText, cFlg打折.RowSel, 3))
    End If
    If mstr打折信息操作标志 = "修改" Then
        ctxt打折.SetFocus
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "收费界面", "frm设置", "cFlg打折_Click", Err.Number, Err.Description, False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set mobjGUI = Nothing
    pblnInUse = False
'    Set pobj收费管理 = Nothing
'    Set pobj业务设置 = Nothing
'    Set pobj单位定位 = Nothing
End Sub




'*************************************************************************
'******************************以下为点击工具栏时的事件*********************
'*************************************************************************

'点击工具栏执行相应添加、删除、修改、清除、保存操作

Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
 On Error GoTo errHandler
    
    Dim lnodparent As Node
    Dim lnodbrother As Node
    Dim lnodtemp As Node
    Dim lstrparentkey As String
    Dim lstrbrotherkey As String
    Dim lstrtemp As String
    Dim lstrmekey As String
    Dim lintcountchild As Integer
    Dim lintpositionme As Integer
    Dim lintpositiontemp As Integer
    Dim lstrkey() As String
    Dim linttemp As Integer
    Select Case cstbShezhi.Caption
        Case "收费项目设置"
            Select Case Operate
                Case "添加"
                    sub收费项目添加
                    cinb收费项目名称.SetFocus
                Case "删除"
                    sub删除收费项目
                    cinb收费项目名称.SetFocus
                Case "修改"
                    sub修改收费项目
                    cinb收费项目名称.SetFocus
                Case "清空"
                    cinb收费项目编号.Enabled = False
                    cinb收费项目名称.Enabled = False
                    cinb收费项目单价.Enabled = False
                    cinb收费项目计量单位.Enabled = False
                    cinb收费项目助记符.Enabled = False
                    cinb收费项目最大单价.Enabled = False
                    cinb收费项目最小单价.Enabled = False
                    ccmb收费项目票据类型.Enabled = False
                    cTvw收费项目1.Enabled = True
                    '如下这些项置为空
                    'cinb收费项目编号.Text = ""
                    cinb收费项目名称.Text = ""
                    cinb收费项目单价.Text = ""
                    cinb收费项目计量单位.Text = ""
                    cinb收费项目助记符.Text = ""
                    cinb收费项目最大单价.Text = ""
                    cinb收费项目最小单价.Text = ""
                    
                    '工具栏按钮可用
                    ctlb设置.Buttons("修改(&U)").Enabled = True
                    ctlb设置.Buttons("删除(&D)").Enabled = True
                    ctlb设置.Buttons("添加(&A)").Enabled = True
                    mstr收费项目操作标志 = ""
                    
                Case "保存"
                    
                    '自动判断收费项目编号，原则是如果没有同级，则为上级编号 & "001"，
                    '否则为同级中最大编号加1
                     If mstr收费项目操作标志 <> "" Then
                        Select Case mstr收费项目操作标志
                        
                            Case "添加"
                            
                                 Set lnodparent = cTvw收费项目1.SelectedItem
                                 If lnodparent Is Nothing Then
                                    MsgBox "请选择项目", vbExclamation, "收费项目设置"
                                    Cancel = True
                                    Exit Sub
                                 End If
                                 lintcountchild = lnodparent.Children
                                 
                                 If lintcountchild > 0 Then
                                     Set lnodbrother = lnodparent.Child
                                     ReDim lstrkey(lintcountchild)
                                     lstrkey(1) = lnodbrother.Key
                                     Set lnodtemp = lnodbrother
                                     For linttemp = 1 To lintcountchild - 1
                                         Set lnodtemp = lnodtemp.Next
                                         lstrkey(linttemp + 1) = lnodtemp.Key
                                     Next
                                     lstrtemp = lstrkey(1)
                                     For linttemp = 1 To lintcountchild
                                         If lstrtemp < lstrkey(linttemp) Then
                                             lstrtemp = lstrkey(linttemp)
                                         End If
                                     Next
                                     lstrmekey = Left(lstrtemp, Len(lstrtemp) - 3) & Format(CInt(Right(lstrtemp, 3)) + 1, "000")
                                     cinb收费项目编号.Text = Right(LTrim(RTrim(lstrmekey)), Len(LTrim(RTrim(lstrmekey))) - 1)
                                 Else
                                     'Format(CStr(Right(nodtt.Key, 3) + 1), "000")
                                     lstrmekey = lnodparent.Key & "001"
                                     cinb收费项目编号.Text = LTrim(RTrim(Right(lstrmekey, Len(lstrmekey) - 1)))
                                 End If
                                 
                                 '判断用户录入是否合理
                                
                                             If Trim(cinb收费项目名称.Text) = "" Then
                                                 MsgBox "收费项目名称不能为空，请录入", vbExclamation, "收费项目设置"
                                                 Cancel = True
                                                 Exit Sub
                                             End If
                                             If Len(LTrim(RTrim(cinb收费项目编号.Text))) / 3 = pint科目级数 Then
                                                 If cinb收费项目单价.Text = "" Then
                                                     MsgBox "最后一级单价不能为空，请修改", vbExclamation, "收费项目设置"
                                                     Cancel = True
                                                     Exit Sub
                                                 Else
                                                      If IsNumeric(cinb收费项目单价.Text) And IsNumeric(cinb收费项目最大单价.Text) And IsNumeric(cinb收费项目最小单价.Text) Then
                                                        If CDbl(cinb收费项目单价.Text) < 0 Or CDbl(cinb收费项目最大单价.Text) < 0 Or CDbl(cinb收费项目最小单价.Text) < 0 Then
                                                            MsgBox "单价、最小单价、最大单价大于零，请修改", vbExclamation, "收费项目设置"
                                                            Cancel = True
                                                            Exit Sub
                                                        End If
                                                      Else
                                                          MsgBox "单价、最小单价、最大单价必须为数值，请修改", vbExclamation, "收费项目设置"
                                                          Cancel = True
                                                          Exit Sub
                                                      End If
                                                      If cinb收费项目单价.Text = 0 Then
                                                        MsgBox "最后一级单价不能为零", vbExclamation, "收费项目设置"
                                                        Cancel = True
                                                        Exit Sub
                                                      End If
                                                      If CDbl(cinb收费项目单价.Text) > CDbl(cinb收费项目最大单价.Text) Or CDbl(cinb收费项目单价.Text) < CDbl(cinb收费项目最小单价.Text) Then
                                                          MsgBox "单价必须在最小单价和最大单价之间，请修改", vbExclamation, "收费项目设置"
                                                          Cancel = True
                                                          Exit Sub
                                                      End If
                                                      
                                                      If cinb收费项目最大单价.Text = "" And cinb收费项目最小单价.Text <> "" Then
                                                          If CDbl(cinb收费项目最小单价.Text) < CDbl(cinb收费项目单价.Text) Then
                                                              cinb收费项目最大单价.Text = cinb收费项目单价.Text
                                                          End If
                                                      End If
                                                      
                                                      '功能：判断计量单位不能不输入，因为在划价界面，必须要有计量单位 时间：2002/02/02 徐冀川
                                                      If cinb收费项目计量单位.Text = "" Or IsNull(cinb收费项目计量单位.Text) Then
                                                         sffuncMsg "请输入计量单位!"
                                                         Cancel = True
                                                         Exit Sub
                                                      End If
                                                      
                                                      If cinb收费项目最大单价.Text <> "" And cinb收费项目最小单价.Text = "" Then
                                                          If CDbl(cinb收费项目最大单价.Text) > CDbl(cinb收费项目单价.Text) Then
                                                              cinb收费项目最小单价.Text = cinb收费项目单价.Text
                                                          End If
                                                      End If
                                                      
                                                      If cinb收费项目最大单价.Text = "" And cinb收费项目最小单价.Text = "" Then
                                                          cinb收费项目最大单价.Text = cinb收费项目单价.Text
                                                          cinb收费项目最小单价.Text = cinb收费项目单价.Text
                                                      End If
                                                 End If
                                             Else
                                                 If Len(LTrim(RTrim(cinb收费项目编号.Text))) / 3 > pint科目级数 Then
                                                     MsgBox "收费项目级数超过限制，该项目设置无效", vbExclamation, "收费项目设置"
                                                     Cancel = True
                                                     Exit Sub
                                                 Else
                                                     cinb收费项目单价.Text = "0"
                                                     cinb收费项目最大单价.Text = "0"
                                                     cinb收费项目最小单价.Text = "0"
                                                 End If
                                             End If
                                             
                                             '向数据库中添加记录
                                             Set mcol收费项目添加 = New Collection
                                             mcol收费项目添加.Add cinb收费项目编号.Text, "收费项目编号"
                                             mcol收费项目添加.Add cinb收费项目名称.Text, "收费项目名称"
                                             mcol收费项目添加.Add cinb收费项目单价.Text, "单价"
                                             mcol收费项目添加.Add cinb收费项目计量单位.Text, "计量单位"
                                             If Not ccmb收费项目票据类型.Text = "" Then
                                                mcol收费项目添加.Add ccmb收费项目票据类型.ItemData(ccmb收费项目票据类型.ListIndex), "票据类型编号"
                                             Else
                                                MsgBox "请选择票据类型", vbExclamation, "收费项目设置"
                                                Cancel = True
                                                Exit Sub
                                                Set mcol收费项目修改 = Nothing
                                             End If
                                             mcol收费项目添加.Add cinb收费项目最小单价.Text, "最小单价"
                                             mcol收费项目添加.Add cinb收费项目最大单价.Text, "最大单价"
                                             mcol收费项目添加.Add cinb收费项目助记符.Text, "助记符"
                                             If pobj收费管理.func添加收费项目(mcol收费项目添加) Then
                                                 '添加成功则增加节点.
                                                 Dim lstr添加节点 As String
                                                 lstr添加节点 = lnodparent.Key
                                                 Set lnodtemp = cTvw收费项目1.Nodes.Add(lstr添加节点, tvwChild, lstrmekey, cinb收费项目名称.Text)
                                                 Set lnodtemp = ctvw收费项目2.Nodes.Add(lstr添加节点, tvwChild, lstrmekey, cinb收费项目名称.Text)
                                                 cTvw收费项目1.Refresh
                                                 ctvw收费项目2.Refresh
                                             Else
                                                Cancel = True
                                             End If
                                             Set mcol收费项目添加 = Nothing
                                             
                            Case "修改"                                     '修改
                            
                                '判断用户录入是否合理。
                                    If Trim(cinb收费项目名称.Text) = "" Or cinb收费项目编号.Text = "" Then
                                        MsgBox "收费项目名称或收费编号不能为空，请录入", vbExclamation, "收费项目设置"
                                        Cancel = True
                                        Exit Sub
                                    End If
                                    If Len(LTrim(RTrim(cinb收费项目编号.Text))) / 3 = pint科目级数 Then
                                        If cinb收费项目单价.Text = "" Then
                                            MsgBox "最后一级单价不能为空，请修改", vbExclamation, "收费项目设置"
                                            Cancel = True
                                            Exit Sub
                                        Else
                                             If IsNumeric(cinb收费项目单价.Text) And IsNumeric(cinb收费项目最大单价.Text) And IsNumeric(cinb收费项目最小单价.Text) Then
                                                If CDbl(cinb收费项目单价.Text) < 0 Or CDbl(cinb收费项目最大单价.Text) < 0 Or CDbl(cinb收费项目最小单价.Text) < 0 Then
                                                    MsgBox "单价、最小单价、最大单价大于零，请修改", vbExclamation, "收费项目设置"
                                                    Cancel = True
                                                    Exit Sub
                                                End If
                                             Else
                                                 MsgBox "单价、最小单价、最大单价必须为数值，请修改", vbExclamation, "收费项目设置"
                                                 Cancel = True
                                                 Exit Sub
                                             End If
                                             If cinb收费项目单价.Text = "0" Then
                                               MsgBox "最后一级单价不能为零", vbExclamation, "收费项目设置"
                                               Cancel = True
                                               Exit Sub
                                             End If
                                            '功能：判断计量单位不能不输入，因为在划价界面，必须要有计量单位 时间：2002/02/02 徐冀川
                                            If cinb收费项目计量单位.Text = "" Or IsNull(cinb收费项目计量单位.Text) Then
                                               sffuncMsg "请输入计量单位!"
                                               Cancel = True
                                               Exit Sub
                                            End If
                                             If CDbl(cinb收费项目单价.Text) > CDbl(cinb收费项目最大单价.Text) Or CDbl(cinb收费项目单价.Text) < CDbl(cinb收费项目最小单价.Text) Then
                                                 MsgBox "单价必须在最小单价和最大单价之间，请修改", vbExclamation, "收费项目设置"
                                                 Cancel = True
                                                 Exit Sub
                                             End If
                                             
                                             If cinb收费项目最大单价.Text = "" And cinb收费项目最小单价.Text <> "" Then
                                                 If CDbl(cinb收费项目最小单价.Text) < CDbl(cinb收费项目单价.Text) Then
                                                     cinb收费项目最大单价.Text = cinb收费项目单价.Text
                                                 End If
                                             End If
                                        
                                             If cinb收费项目最大单价.Text <> "" And cinb收费项目最小单价.Text = "" Then
                                                 If CDbl(cinb收费项目最大单价.Text) > CDbl(cinb收费项目单价.Text) Then
                                                     cinb收费项目最小单价.Text = cinb收费项目单价.Text
                                                 End If
                                             End If
                                             
                                             If cinb收费项目最大单价.Text = "" And cinb收费项目最小单价.Text = "" Then
                                                 cinb收费项目最大单价.Text = cinb收费项目单价.Text
                                                 cinb收费项目最小单价.Text = cinb收费项目单价.Text
                                             End If
                                        End If
                                    Else
                                        If Len(LTrim(RTrim(cinb收费项目编号.Text))) / 3 > pint科目级数 Then
                                            MsgBox "收费项目级数超过限制，该项目设置无效", vbExclamation, "收费项目设置"
                                            Cancel = True
                                            Exit Sub
                                        Else
                                            cinb收费项目单价.Text = "0"
                                            cinb收费项目最大单价.Text = "0"
                                            cinb收费项目最小单价.Text = "0"
                                        End If
                                    End If
                                
                                '用户录入合理则保存信息
                                Set mcol收费项目修改 = New Collection
                                mcol收费项目修改.Add cinb收费项目编号.Text, "收费项目编号"
                                mcol收费项目修改.Add cinb收费项目名称.Text, "收费项目名称"
                                mcol收费项目修改.Add LTrim(RTrim(cinb收费项目单价.Text)), "单价"
                                mcol收费项目修改.Add cinb收费项目计量单位.Text, "计量单位"
                                If Not ccmb收费项目票据类型.Text = "" Then
                                    mcol收费项目修改.Add ccmb收费项目票据类型.ItemData(ccmb收费项目票据类型.ListIndex), "票据类型编号"
                                Else
                                    MsgBox "请选择票据类型", vbExclamation, "收费项目设置"
                                    Cancel = True
                                    Exit Sub
                                    Set mcol收费项目修改 = Nothing
                                End If
                                mcol收费项目修改.Add cinb收费项目最小单价.Text, "最小单价"
                                mcol收费项目修改.Add cinb收费项目最大单价.Text, "最大单价"
                                mcol收费项目修改.Add cinb收费项目助记符.Text, "助记符"
                                If pobj收费管理.func修改收费项目(mcol收费项目修改) Then
                                    cTvw收费项目1.SelectedItem.Text = cinb收费项目名称.Text
                                    ctvw收费项目2.Nodes(cTvw收费项目1.SelectedItem.Key).Text = cinb收费项目名称.Text
                                Else
                                    Cancel = True
                                End If
                                Set mcol收费项目添加 = Nothing
                          '删除
                            Case "删除"
                                If cTvw收费项目1.SelectedItem Is Nothing Then Exit Sub
                                lintcountchild = cTvw收费项目1.SelectedItem.Children
                                If (Len(cTvw收费项目1.SelectedItem.Key) - 1) / 3 < pint科目级数 Then
                                    If cTvw收费项目1.SelectedItem.Key = "s" Then
                                        MsgBox "该节点不允许删除，操作无效", vbExclamation, "收费项目设置"
                                        '设树形结构为可用状态
                                        cTvw收费项目1.Enabled = True
                                        Cancel = True
                                        Exit Sub
                                    Else
'                                        If MsgBox("你要删除该节点以及其下所包含的所有项目吗", vbExclamation + vbOKCancel, "收费项目设置") = "vbok" Then
'                                            For linttemp = 1 To lintcountchild - 1
'                                                Set lnodtemp = cTvw收费项目1.SelectedItem.Child
'                                                pobj收费管理.func删除收费项目 LTrim(RTrim(Right(lnodtemp.Key, Len(lnodtemp.Key) - 1)))
'                                                cTvw收费项目1.Nodes.Remove (lnodtemp.Key)
'                                                ctvw收费项目2.Nodes.Remove (lnodtemp.Key)
'                                            Next
'                                            pobj收费管理.func删除收费项目 cinb收费项目编号.Text
'                                            ctvw收费项目2.Nodes.Remove (cTvw收费项目1.SelectedItem.Key)
'                                            cTvw收费项目1.Nodes.Remove (cTvw收费项目1.SelectedItem.Key)
'                                        Else
'                                            Exit Sub
'                                        End If
                                        If cTvw收费项目1.SelectedItem.Children > 0 Then
                                            MsgBox "请从最下级开始删除", vbExclamation, "收费项目设置"
                                            
                                            '设树形结构为可用状态
                                            cTvw收费项目1.Enabled = True
                                            Cancel = True
                                            Exit Sub
                                        Else
                                            If MsgBox("确定删除" & cTvw收费项目1.SelectedItem.Text, vbOKCancel, "收费项目设置") = vbCancel Then
                                                Cancel = True
                                                Exit Sub
                                            End If
                                            If pobj收费管理.func删除收费项目(cinb收费项目编号.Text) Then
                                                ctvw收费项目2.Nodes.Remove (cTvw收费项目1.SelectedItem.Key)
                                                cTvw收费项目1.Nodes.Remove (cTvw收费项目1.SelectedItem.Key)
                                            Else
                                                Cancel = True
                                            End If
                                        End If
                                    End If
                                Else
                                    If cTvw收费项目1.SelectedItem.Children = 0 Or (Len(cTvw收费项目1.SelectedItem.Key) - 1) / 3 = pint科目级数 Then
                                        If MsgBox("确定删除" & cTvw收费项目1.SelectedItem.Text, vbOKCancel, "收费项目设置") = vbCancel Then
                                                Cancel = True
                                                Exit Sub
                                        End If
                                        If pobj收费管理.func删除收费项目(LTrim(RTrim(Right(cTvw收费项目1.SelectedItem.Key, Len(cTvw收费项目1.SelectedItem.Key) - 1)))) Then
                                            Dim lstr删除节点 As String
                                            lstr删除节点 = cTvw收费项目1.SelectedItem.Key
                                            ctvw收费项目2.Nodes.Remove (lstr删除节点)
                                            cTvw收费项目1.Nodes.Remove (lstr删除节点)
                                        Else
                                            Cancel = True
                                        End If
                                    End If
                                End If
                                cTvw收费项目1.Enabled = True
                        End Select
                        cinb收费项目编号.Enabled = False
                        cinb收费项目名称.Enabled = False
                        cinb收费项目单价.Enabled = False
                        cinb收费项目计量单位.Enabled = False
                        cinb收费项目助记符.Enabled = False
                        cinb收费项目最大单价.Enabled = False
                        cinb收费项目最小单价.Enabled = False
                        ccmb收费项目票据类型.Enabled = False
                        
                        '如下这些项置为空
                        'cinb收费项目编号.Text = ""
                        cinb收费项目名称.Text = ""
                        cinb收费项目单价.Text = ""
                        cinb收费项目计量单位.Text = ""
                        cinb收费项目助记符.Text = ""
                        cinb收费项目最大单价.Text = ""
                        cinb收费项目最小单价.Text = ""
                        ccmb收费项目票据类型.ListIndex = -1
                        
                        '工具栏按钮可用
                        ctlb设置.Buttons("修改(&U)").Enabled = True
                        ctlb设置.Buttons("删除(&D)").Enabled = True
                        ctlb设置.Buttons("添加(&A)").Enabled = True
                        mstr收费项目操作标志 = ""
                    Else
                        'msgbox "请选择操作方式",vbInformation,"收费项目设置"
                    End If
                
                Case "退出"
                    Unload Me
            End Select
            Cancel = True
        Case "票据设置"
            Select Case Operate
                Case "添加"
                    sub票据信息添加
                    cinb票据名称.SetFocus
                Case "修改"
                    sub票据信息修改
                    cFlg票据信息.SetFocus
                Case "清空"
                    sub票据信息清除
                    
                Case "删除"
                    sub票据信息删除
                    cFlg票据信息.SetFocus
                Case "保存"
                    If Not func票据信息保存 Then
                        Cancel = True
                        Exit Sub
                    End If
                Case "退出"
                    Unload Me
                    
                Case Else
            End Select
            Cancel = True
        Case "收费标准设置"
            Select Case Operate
                Case "添加"
                    sub收费标准添加
                    cflg收费标准.SetFocus
                Case "删除"
                    sub收费标准删除
                    Cchk收费标准列表.SetFocus
                Case "修改"
                    sub收费标准修改
                    Cchk收费标准列表.SetFocus
                Case "清空"
                    sub收费标准清除
                    
                Case "保存"
                    If Not func收费标准保存 Then
                        Cancel = True
                        Exit Sub
                    End If
                Case "退出"
                    Unload Me
            End Select
            Cancel = True
        Case "打折设置"
            Select Case Operate
                Case "添加"
                    sub打折信息添加
                    ccmd单位定位.SetFocus
                Case "删除"
                    sub打折信息删除
                    If cFlg打折.Enabled And cFlg打折.Visible Then cFlg打折.SetFocus
                Case "修改"
                    sub打折信息修改
                    cFlg打折.SetFocus
                Case "清空"
                    sub打折信息清除
                Case "保存"
                    
                    If func打折信息保存 Then
                        If mstr打折控制操作标志 = "添加" Then
                            Set mcol业务配置 = New Collection
                            mcol业务配置.Add mint打折控制, "打折控制"
                            mcol业务配置.Add pint科目级数, "科目级数"
                            If pobj业务设置.func修改业务配置信息(mcol业务配置) Then
                                Cancel = True
                                pint打折控制 = mint打折控制
                            Else
                                MsgBox "打折控制情况修改有问题，请检查机器或网络连接是否正常，如还有问题，请与系统管理员联系", vbExclamation, "打折控制设置"
                                mint打折控制 = pint打折控制
                                Cancel = True
                                Exit Sub
                            End If
                            Set mcol业务配置 = Nothing
                            mstr打折控制操作标志 = ""
                        End If
                        Cancel = True
                        sub打折信息清除
                        Exit Sub
                    Else
                        MsgBox "打折信息不能保存", vbExclamation, "打折设置"
                        Cancel = True
                        Exit Sub
                    End If
                Case "退出"
                    Unload Me
            End Select
            Cancel = True
        '功能：对其它设置进行工栏上的操作控制(保存、修改...)
        '作者：徐冀川
        '时间：2002/07/17
        Case "其它设置"
            Select Case Operate
                Case "添加"
                     ctlb设置.Buttons(1).Enabled = False
                     ctlb设置.Buttons(3).Enabled = False
                     ctlb设置.Buttons(4).Enabled = False
                Case "删除"
                     ctlb设置.Buttons(1).Enabled = False
                     ctlb设置.Buttons(3).Enabled = False
                     ctlb设置.Buttons(4).Enabled = False
                Case "修改"
                    Ctxt收据号值.Enabled = True
                    Ctxt收据号值.SetFocus
                Case "清空"
                    ctlb设置.Buttons(1).Enabled = False
                    ctlb设置.Buttons(3).Enabled = False
                    ctlb设置.Buttons(4).Enabled = False
                Case "保存"
                    sub保存设置收据号值
                Case "退出"
                    Unload Me
            End Select
            Cancel = True
    End Select
    Exit Sub
errHandler:
    If Err.Number = -2147217873 Then
        MsgBox "该项目正在使用，不能删除", vbExclamation, "设置"
        Cancel = True
        sub收费项目清除
    Else
        sffuncMsg Err.Description, sf警告
    End If
    
End Sub


'功能：保存设置值
'作者：徐冀川
'时间：2002/07/17
Private Sub sub保存设置收据号值()
On Error GoTo errHandler
    Dim lstrSql As String           '定义变量记录SQL语句
    Dim lint当前值 As Long          '记录当前值
    
    lint当前值 = Ctxt收据号值.Text
    
    lstrSql = "update 系统管理_系统编号生成记录表 set 当前值= " & _
              lint当前值 & " where 业务名称= '收费管理' and 编号名称='收据号'"
    dafuncGetData (lstrSql)
Exit Sub
errHandler:
    sfsub错误处理 "收费界面", "frm设置", "sub保存设置收据号值", Err.Number, Err.Description, False
End Sub




Private Function funcgetfilename(filename As String) As String
    Dim lintOffset As Integer
    Dim lstrResult As String
    
    lstrResult = filename
    lintOffset = InStr(lstrResult, "\")
    If lintOffset = 0 Then
        funcgetfilename = filename
        Exit Function
    End If
    
    Do While lintOffset <> 0
        lstrResult = Mid(lstrResult, lintOffset + 1)
        lintOffset = InStr(lstrResult, "\")
    Loop
    funcgetfilename = lstrResult
End Function
Private Sub sub收费项目清除()
    cinb收费项目编号.Enabled = False
    cinb收费项目名称.Enabled = False
    cinb收费项目单价.Enabled = False
    cinb收费项目计量单位.Enabled = False
    cinb收费项目助记符.Enabled = False
    cinb收费项目最大单价.Enabled = False
    cinb收费项目最小单价.Enabled = False
    ccmb收费项目票据类型.Enabled = False
    cTvw收费项目1.Enabled = True
    '如下这些项置为空
    'cinb收费项目编号.Text = ""
    cinb收费项目名称.Text = ""
    cinb收费项目单价.Text = ""
    cinb收费项目计量单位.Text = ""
    cinb收费项目助记符.Text = ""
    cinb收费项目最大单价.Text = ""
    cinb收费项目最小单价.Text = ""
    
    '工具栏按钮可用
    ctlb设置.Buttons("修改(&U)").Enabled = True
    ctlb设置.Buttons("删除(&D)").Enabled = True
    ctlb设置.Buttons("添加(&A)").Enabled = True
    mstr收费项目操作标志 = ""
End Sub
