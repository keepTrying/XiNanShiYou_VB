VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm֢״ѯ�� 
   Caption         =   "֢״ѯ��"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13845
   LinkTopic       =   "Form1"
   ScaleHeight     =   7080
   ScaleWidth      =   13845
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Caption         =   "�˳�"
      Height          =   375
      Left            =   2160
      TabIndex        =   155
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   375
      Left            =   480
      TabIndex        =   154
      Top             =   240
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5655
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   9975
      _Version        =   393216
      Tabs            =   9
      TabsPerRow      =   9
      TabHeight       =   520
      BackColor       =   0
      ForeColor       =   16711680
      TabCaption(0)   =   "��ϵͳ"
      TabPicture(0)   =   "frm֢״ѯ��.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "freRadiation"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "freOrdinary"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "freNuclear"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "��Ѫ���ڷ���ϵͳ"
      TabPicture(1)   =   "frm֢״ѯ��.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame5"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "����ϵͳ"
      TabPicture(2)   =   "frm֢״ѯ��.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame6"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "����ϵͳ"
      TabPicture(3)   =   "frm֢״ѯ��.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame7"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "���⼰��֫�ؽ�"
      TabPicture(4)   =   "frm֢״ѯ��.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame8"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "�۶��ǿ�ǻ���ʺ�"
      TabPicture(5)   =   "frm֢״ѯ��.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame9"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "��Ѫ��ϵͳ"
      TabPicture(6)   =   "frm֢״ѯ��.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame10"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "����ϵͳ"
      TabPicture(7)   =   "frm֢״ѯ��.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Frame11"
      Tab(7).ControlCount=   1
      TabCaption(8)   =   "����"
      TabPicture(8)   =   "frm֢״ѯ��.frx":00E0
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Frame12"
      Tab(8).Control(0).Enabled=   0   'False
      Tab(8).ControlCount=   1
      Begin VB.Frame Frame11 
         Height          =   4095
         Left            =   -74640
         TabIndex        =   316
         Top             =   600
         Width           =   7335
         Begin VB.TextBox Text144 
            Height          =   270
            Index           =   6
            Left            =   3480
            TabIndex        =   333
            Top             =   2760
            Width           =   1695
         End
         Begin VB.TextBox Text144 
            Height          =   270
            Index           =   5
            Left            =   3480
            TabIndex        =   332
            Top             =   2400
            Width           =   1695
         End
         Begin VB.TextBox Text144 
            Height          =   270
            Index           =   4
            Left            =   3480
            TabIndex        =   331
            Top             =   2040
            Width           =   1695
         End
         Begin VB.TextBox Text144 
            Height          =   270
            Index           =   3
            Left            =   3480
            TabIndex        =   330
            Top             =   1680
            Width           =   1695
         End
         Begin VB.TextBox Text144 
            Height          =   270
            Index           =   2
            Left            =   3480
            TabIndex        =   329
            Top             =   1320
            Width           =   1695
         End
         Begin VB.TextBox Text144 
            Height          =   270
            Index           =   1
            Left            =   3480
            TabIndex        =   328
            Top             =   960
            Width           =   1695
         End
         Begin VB.ComboBox Combo128 
            Height          =   300
            Index           =   6
            ItemData        =   "frm֢״ѯ��.frx":00FC
            Left            =   1920
            List            =   "frm֢״ѯ��.frx":010F
            TabIndex        =   327
            Text            =   "-"
            Top             =   2760
            Width           =   975
         End
         Begin VB.ComboBox Combo128 
            Height          =   300
            Index           =   5
            ItemData        =   "frm֢״ѯ��.frx":0126
            Left            =   1920
            List            =   "frm֢״ѯ��.frx":0139
            TabIndex        =   326
            Text            =   "-"
            Top             =   2400
            Width           =   975
         End
         Begin VB.ComboBox Combo128 
            Height          =   300
            Index           =   4
            ItemData        =   "frm֢״ѯ��.frx":0150
            Left            =   1920
            List            =   "frm֢״ѯ��.frx":0163
            TabIndex        =   325
            Text            =   "-"
            Top             =   2040
            Width           =   975
         End
         Begin VB.ComboBox Combo128 
            Height          =   300
            Index           =   3
            ItemData        =   "frm֢״ѯ��.frx":017A
            Left            =   1920
            List            =   "frm֢״ѯ��.frx":018D
            TabIndex        =   324
            Text            =   "-"
            Top             =   1680
            Width           =   975
         End
         Begin VB.ComboBox Combo128 
            Height          =   300
            Index           =   2
            ItemData        =   "frm֢״ѯ��.frx":01A4
            Left            =   1920
            List            =   "frm֢״ѯ��.frx":01B7
            TabIndex        =   323
            Text            =   "-"
            Top             =   1320
            Width           =   975
         End
         Begin VB.ComboBox Combo128 
            Height          =   300
            Index           =   1
            ItemData        =   "frm֢״ѯ��.frx":01CE
            Left            =   1920
            List            =   "frm֢״ѯ��.frx":01E1
            TabIndex        =   322
            Text            =   "-"
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox Text144 
            Height          =   270
            Index           =   0
            Left            =   3480
            TabIndex        =   318
            Top             =   600
            Width           =   1695
         End
         Begin VB.ComboBox Combo128 
            Height          =   300
            Index           =   0
            ItemData        =   "frm֢״ѯ��.frx":01F8
            Left            =   1920
            List            =   "frm֢״ѯ��.frx":020B
            TabIndex        =   317
            Text            =   "-"
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "��Ѫ"
            Height          =   255
            Index           =   6
            Left            =   600
            TabIndex        =   340
            Top             =   2760
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "����"
            Height          =   255
            Index           =   5
            Left            =   600
            TabIndex        =   339
            Top             =   2400
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "��к"
            Height          =   255
            Index           =   3
            Left            =   600
            TabIndex        =   338
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "���͡���ʹ"
            Height          =   255
            Index           =   2
            Left            =   600
            TabIndex        =   337
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "������ʹ"
            Height          =   255
            Index           =   1
            Left            =   600
            TabIndex        =   336
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "���ġ�Ż��"
            Height          =   255
            Index           =   4
            Left            =   600
            TabIndex        =   335
            Top             =   2040
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "ʳ������"
            Height          =   255
            Index           =   0
            Left            =   600
            TabIndex        =   334
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label40 
            Caption         =   "����ʱ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3480
            TabIndex        =   321
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label41 
            Caption         =   "��   ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1920
            TabIndex        =   320
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label42 
            Caption         =   "��   Ŀ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   319
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame12 
         Height          =   4095
         Left            =   -74640
         TabIndex        =   124
         Top             =   600
         Width           =   7335
         Begin VB.CommandButton Command4 
            Caption         =   "ɾ��"
            Height          =   375
            Left            =   2280
            TabIndex        =   158
            Top             =   1080
            Width           =   1215
         End
         Begin VB.CommandButton Command3 
            Caption         =   "����"
            Height          =   375
            Left            =   600
            TabIndex        =   157
            Top             =   1080
            Width           =   1335
         End
         Begin VB.TextBox Text247 
            Height          =   270
            Left            =   600
            TabIndex        =   153
            Top             =   600
            Width           =   1215
         End
         Begin VB.ComboBox Combo228 
            Height          =   300
            Index           =   0
            ItemData        =   "frm֢״ѯ��.frx":0222
            Left            =   1920
            List            =   "frm֢״ѯ��.frx":0235
            TabIndex        =   152
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox Text244 
            Height          =   270
            Index           =   0
            Left            =   3480
            TabIndex        =   151
            Top             =   600
            Width           =   1695
         End
         Begin VSFlex8Ctl.VSFlexGrid cgrd����֢״ 
            Height          =   2175
            Left            =   600
            TabIndex        =   156
            Top             =   1680
            Width           =   4815
            _cx             =   8493
            _cy             =   3836
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
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
            Cols            =   3
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frm֢״ѯ��.frx":024C
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
         Begin VB.Label Label45 
            Caption         =   "��   Ŀ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   127
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label44 
            Caption         =   "��   ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1920
            TabIndex        =   126
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label43 
            Caption         =   "����ʱ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3480
            TabIndex        =   125
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame10 
         Height          =   4095
         Left            =   -74640
         TabIndex        =   120
         Top             =   600
         Width           =   7335
         Begin VB.TextBox Text44 
            Height          =   270
            Index           =   1
            Left            =   3480
            TabIndex        =   315
            Top             =   960
            Width           =   1695
         End
         Begin VB.TextBox Text44 
            Height          =   270
            Index           =   0
            Left            =   3480
            TabIndex        =   314
            Top             =   600
            Width           =   1695
         End
         Begin VB.ComboBox Combo28 
            Height          =   300
            Index           =   1
            ItemData        =   "frm֢״ѯ��.frx":02B1
            Left            =   1920
            List            =   "frm֢״ѯ��.frx":02C4
            TabIndex        =   313
            Text            =   "-"
            Top             =   960
            Width           =   975
         End
         Begin VB.ComboBox Combo28 
            Height          =   300
            Index           =   0
            ItemData        =   "frm֢״ѯ��.frx":02DB
            Left            =   1920
            List            =   "frm֢״ѯ��.frx":02EE
            TabIndex        =   312
            Text            =   "-"
            Top             =   600
            Width           =   975
         End
         Begin VB.ComboBox Combo28 
            Height          =   300
            Index           =   2
            ItemData        =   "frm֢״ѯ��.frx":0305
            Left            =   1920
            List            =   "frm֢״ѯ��.frx":0318
            TabIndex        =   149
            Text            =   "-"
            Top             =   1320
            Width           =   975
         End
         Begin VB.TextBox Text44 
            Height          =   270
            Index           =   2
            Left            =   3480
            TabIndex        =   148
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label Label76 
            Caption         =   "��ǰ����ʹ"
            Height          =   255
            Index           =   2
            Left            =   600
            TabIndex        =   311
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label Label76 
            Caption         =   "��ǰ������"
            Height          =   255
            Index           =   1
            Left            =   600
            TabIndex        =   310
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label76 
            Caption         =   "�ļ�"
            Height          =   255
            Index           =   0
            Left            =   600
            TabIndex        =   150
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label39 
            Caption         =   "��   Ŀ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   123
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label38 
            Caption         =   "��   ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1920
            TabIndex        =   122
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label37 
            Caption         =   "����ʱ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3480
            TabIndex        =   121
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame9 
         Height          =   4335
         Left            =   -74640
         TabIndex        =   116
         Top             =   600
         Width           =   12375
         Begin VB.TextBox Text37 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "yyyy/MM/dd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   3
            EndProperty
            Height          =   270
            Index           =   19
            Left            =   9000
            TabIndex        =   309
            Top             =   3840
            Width           =   1695
         End
         Begin VB.TextBox Text37 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "yyyy/MM/dd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   3
            EndProperty
            Height          =   270
            Index           =   18
            Left            =   9000
            TabIndex        =   308
            Top             =   3480
            Width           =   1695
         End
         Begin VB.TextBox Text37 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "yyyy/MM/dd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   3
            EndProperty
            Height          =   270
            Index           =   17
            Left            =   9000
            TabIndex        =   307
            Top             =   3120
            Width           =   1695
         End
         Begin VB.TextBox Text37 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "yyyy/MM/dd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   3
            EndProperty
            Height          =   270
            Index           =   16
            Left            =   9000
            TabIndex        =   306
            Top             =   2760
            Width           =   1695
         End
         Begin VB.TextBox Text37 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "yyyy/MM/dd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   3
            EndProperty
            Height          =   270
            Index           =   15
            Left            =   9000
            TabIndex        =   305
            Top             =   2400
            Width           =   1695
         End
         Begin VB.TextBox Text37 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "yyyy/MM/dd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   3
            EndProperty
            Height          =   270
            Index           =   14
            Left            =   9000
            TabIndex        =   304
            Top             =   2040
            Width           =   1695
         End
         Begin VB.TextBox Text37 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "yyyy/MM/dd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   3
            EndProperty
            Height          =   270
            Index           =   13
            Left            =   9000
            TabIndex        =   303
            Top             =   1680
            Width           =   1695
         End
         Begin VB.TextBox Text37 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "yyyy/MM/dd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   3
            EndProperty
            Height          =   270
            Index           =   12
            Left            =   9000
            TabIndex        =   302
            Top             =   1320
            Width           =   1695
         End
         Begin VB.TextBox Text37 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "yyyy/MM/dd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   3
            EndProperty
            Height          =   270
            Index           =   11
            Left            =   9000
            TabIndex        =   301
            Top             =   960
            Width           =   1695
         End
         Begin VB.TextBox Text37 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "yyyy/MM/dd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   3
            EndProperty
            Height          =   270
            Index           =   10
            Left            =   9000
            TabIndex        =   300
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox Text37 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "yyyy/MM/dd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   3
            EndProperty
            Height          =   270
            Index           =   9
            Left            =   3480
            TabIndex        =   299
            Top             =   3840
            Width           =   1695
         End
         Begin VB.TextBox Text37 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "yyyy/MM/dd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   3
            EndProperty
            Height          =   270
            Index           =   8
            Left            =   3480
            TabIndex        =   298
            Top             =   3480
            Width           =   1695
         End
         Begin VB.TextBox Text37 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "yyyy/MM/dd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   3
            EndProperty
            Height          =   270
            Index           =   7
            Left            =   3480
            TabIndex        =   297
            Top             =   3120
            Width           =   1695
         End
         Begin VB.TextBox Text37 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "yyyy/MM/dd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   3
            EndProperty
            Height          =   270
            Index           =   6
            Left            =   3480
            TabIndex        =   296
            Top             =   2760
            Width           =   1695
         End
         Begin VB.TextBox Text37 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "yyyy/MM/dd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   3
            EndProperty
            Height          =   270
            Index           =   5
            Left            =   3480
            TabIndex        =   295
            Top             =   2400
            Width           =   1695
         End
         Begin VB.TextBox Text37 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "yyyy/MM/dd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   3
            EndProperty
            Height          =   270
            Index           =   4
            Left            =   3480
            TabIndex        =   294
            Top             =   2040
            Width           =   1695
         End
         Begin VB.TextBox Text37 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "yyyy/MM/dd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   3
            EndProperty
            Height          =   270
            Index           =   3
            Left            =   3480
            TabIndex        =   293
            Top             =   1680
            Width           =   1695
         End
         Begin VB.TextBox Text37 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "yyyy/MM/dd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   3
            EndProperty
            Height          =   270
            Index           =   2
            Left            =   3480
            TabIndex        =   292
            Top             =   1320
            Width           =   1695
         End
         Begin VB.TextBox Text37 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "yyyy/MM/dd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   3
            EndProperty
            Height          =   270
            Index           =   1
            Left            =   3480
            TabIndex        =   291
            Top             =   960
            Width           =   1695
         End
         Begin VB.ComboBox Combo44 
            Height          =   300
            Index           =   19
            ItemData        =   "frm֢״ѯ��.frx":032F
            Left            =   7440
            List            =   "frm֢״ѯ��.frx":0342
            TabIndex        =   290
            Text            =   "-"
            Top             =   3840
            Width           =   975
         End
         Begin VB.ComboBox Combo44 
            Height          =   300
            Index           =   18
            ItemData        =   "frm֢״ѯ��.frx":0359
            Left            =   7440
            List            =   "frm֢״ѯ��.frx":036C
            TabIndex        =   289
            Text            =   "-"
            Top             =   3480
            Width           =   975
         End
         Begin VB.ComboBox Combo44 
            Height          =   300
            Index           =   17
            ItemData        =   "frm֢״ѯ��.frx":0383
            Left            =   7440
            List            =   "frm֢״ѯ��.frx":0396
            TabIndex        =   288
            Text            =   "-"
            Top             =   3120
            Width           =   975
         End
         Begin VB.ComboBox Combo44 
            Height          =   300
            Index           =   16
            ItemData        =   "frm֢״ѯ��.frx":03AD
            Left            =   7440
            List            =   "frm֢״ѯ��.frx":03C0
            TabIndex        =   287
            Text            =   "-"
            Top             =   2760
            Width           =   975
         End
         Begin VB.ComboBox Combo44 
            Height          =   300
            Index           =   15
            ItemData        =   "frm֢״ѯ��.frx":03D7
            Left            =   7440
            List            =   "frm֢״ѯ��.frx":03EA
            TabIndex        =   286
            Text            =   "-"
            Top             =   2400
            Width           =   975
         End
         Begin VB.ComboBox Combo44 
            Height          =   300
            Index           =   14
            ItemData        =   "frm֢״ѯ��.frx":0401
            Left            =   7440
            List            =   "frm֢״ѯ��.frx":0414
            TabIndex        =   285
            Text            =   "-"
            Top             =   2040
            Width           =   975
         End
         Begin VB.ComboBox Combo44 
            Height          =   300
            Index           =   13
            ItemData        =   "frm֢״ѯ��.frx":042B
            Left            =   7440
            List            =   "frm֢״ѯ��.frx":043E
            TabIndex        =   284
            Text            =   "-"
            Top             =   1680
            Width           =   975
         End
         Begin VB.ComboBox Combo44 
            Height          =   300
            Index           =   12
            ItemData        =   "frm֢״ѯ��.frx":0455
            Left            =   7440
            List            =   "frm֢״ѯ��.frx":0468
            TabIndex        =   283
            Text            =   "-"
            Top             =   1320
            Width           =   975
         End
         Begin VB.ComboBox Combo44 
            Height          =   300
            Index           =   11
            ItemData        =   "frm֢״ѯ��.frx":047F
            Left            =   7440
            List            =   "frm֢״ѯ��.frx":0492
            TabIndex        =   282
            Text            =   "-"
            Top             =   960
            Width           =   975
         End
         Begin VB.ComboBox Combo44 
            Height          =   300
            Index           =   10
            ItemData        =   "frm֢״ѯ��.frx":04A9
            Left            =   7440
            List            =   "frm֢״ѯ��.frx":04BC
            TabIndex        =   281
            Text            =   "-"
            Top             =   600
            Width           =   975
         End
         Begin VB.ComboBox Combo44 
            Height          =   300
            Index           =   9
            ItemData        =   "frm֢״ѯ��.frx":04D3
            Left            =   1920
            List            =   "frm֢״ѯ��.frx":04E6
            TabIndex        =   280
            Text            =   "-"
            Top             =   3840
            Width           =   975
         End
         Begin VB.ComboBox Combo44 
            Height          =   300
            Index           =   8
            ItemData        =   "frm֢״ѯ��.frx":04FD
            Left            =   1920
            List            =   "frm֢״ѯ��.frx":0510
            TabIndex        =   279
            Text            =   "-"
            Top             =   3480
            Width           =   975
         End
         Begin VB.ComboBox Combo44 
            Height          =   300
            Index           =   7
            ItemData        =   "frm֢״ѯ��.frx":0527
            Left            =   1920
            List            =   "frm֢״ѯ��.frx":053A
            TabIndex        =   278
            Text            =   "-"
            Top             =   3120
            Width           =   975
         End
         Begin VB.ComboBox Combo44 
            Height          =   300
            Index           =   6
            ItemData        =   "frm֢״ѯ��.frx":0551
            Left            =   1920
            List            =   "frm֢״ѯ��.frx":0564
            TabIndex        =   277
            Text            =   "-"
            Top             =   2760
            Width           =   975
         End
         Begin VB.ComboBox Combo44 
            Height          =   300
            Index           =   5
            ItemData        =   "frm֢״ѯ��.frx":057B
            Left            =   1920
            List            =   "frm֢״ѯ��.frx":058E
            TabIndex        =   276
            Text            =   "-"
            Top             =   2400
            Width           =   975
         End
         Begin VB.ComboBox Combo44 
            Height          =   300
            Index           =   4
            ItemData        =   "frm֢״ѯ��.frx":05A5
            Left            =   1920
            List            =   "frm֢״ѯ��.frx":05B8
            TabIndex        =   275
            Text            =   "-"
            Top             =   2040
            Width           =   975
         End
         Begin VB.ComboBox Combo44 
            Height          =   300
            Index           =   3
            ItemData        =   "frm֢״ѯ��.frx":05CF
            Left            =   1920
            List            =   "frm֢״ѯ��.frx":05E2
            TabIndex        =   274
            Text            =   "-"
            Top             =   1680
            Width           =   975
         End
         Begin VB.ComboBox Combo44 
            Height          =   300
            Index           =   2
            ItemData        =   "frm֢״ѯ��.frx":05F9
            Left            =   1920
            List            =   "frm֢״ѯ��.frx":060C
            TabIndex        =   273
            Text            =   "-"
            Top             =   1320
            Width           =   975
         End
         Begin VB.ComboBox Combo44 
            Height          =   300
            Index           =   1
            ItemData        =   "frm֢״ѯ��.frx":0623
            Left            =   1920
            List            =   "frm֢״ѯ��.frx":0636
            TabIndex        =   272
            Text            =   "-"
            Top             =   960
            Width           =   975
         End
         Begin VB.ComboBox Combo44 
            Height          =   300
            Index           =   0
            ItemData        =   "frm֢״ѯ��.frx":064D
            Left            =   1920
            List            =   "frm֢״ѯ��.frx":0660
            TabIndex        =   143
            Text            =   "-"
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox Text37 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "yyyy/MM/dd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   3
            EndProperty
            Height          =   270
            Index           =   0
            Left            =   3480
            TabIndex        =   142
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label Label92 
            Caption         =   "��˻"
            Height          =   255
            Index           =   19
            Left            =   6120
            TabIndex        =   271
            Top             =   3840
            Width           =   735
         End
         Begin VB.Label Label92 
            Caption         =   "�ʲ���ʹ"
            Height          =   255
            Index           =   18
            Left            =   6120
            TabIndex        =   270
            Top             =   3480
            Width           =   735
         End
         Begin VB.Label Label92 
            Caption         =   "��ǻ����"
            Height          =   255
            Index           =   17
            Left            =   6120
            TabIndex        =   269
            Top             =   3120
            Width           =   735
         End
         Begin VB.Label Label92 
            Caption         =   "��ǻ��ζ"
            Height          =   255
            Index           =   16
            Left            =   6120
            TabIndex        =   268
            Top             =   2760
            Width           =   735
         End
         Begin VB.Label Label92 
            Caption         =   "ˢ����Ѫ"
            Height          =   255
            Index           =   15
            Left            =   6120
            TabIndex        =   267
            Top             =   2400
            Width           =   735
         End
         Begin VB.Label Label92 
            Caption         =   "�����ɶ�"
            Height          =   255
            Index           =   14
            Left            =   6120
            TabIndex        =   266
            Top             =   2040
            Width           =   735
         End
         Begin VB.Label Label92 
            Caption         =   "��ʹ"
            Height          =   255
            Index           =   13
            Left            =   6120
            TabIndex        =   265
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label92 
            Caption         =   "����"
            Height          =   255
            Index           =   12
            Left            =   6120
            TabIndex        =   264
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label92 
            Caption         =   "����"
            Height          =   255
            Index           =   11
            Left            =   6120
            TabIndex        =   263
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label92 
            Caption         =   "����"
            Height          =   255
            Index           =   10
            Left            =   6120
            TabIndex        =   262
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label92 
            Caption         =   "����"
            Height          =   255
            Index           =   9
            Left            =   600
            TabIndex        =   261
            Top             =   3840
            Width           =   735
         End
         Begin VB.Label Label92 
            Caption         =   "����Ѫ"
            Height          =   255
            Index           =   8
            Left            =   600
            TabIndex        =   260
            Top             =   3480
            Width           =   735
         End
         Begin VB.Label Label92 
            Caption         =   "�Ƕ�"
            Height          =   255
            Index           =   7
            Left            =   600
            TabIndex        =   259
            Top             =   3120
            Width           =   735
         End
         Begin VB.Label Label92 
            Caption         =   "�Ǹ�"
            Height          =   255
            Index           =   6
            Left            =   600
            TabIndex        =   258
            Top             =   2760
            Width           =   735
         End
         Begin VB.Label Label92 
            Caption         =   "�������"
            Height          =   255
            Index           =   5
            Left            =   600
            TabIndex        =   257
            Top             =   2400
            Width           =   735
         End
         Begin VB.Label Label92 
            Caption         =   "����"
            Height          =   255
            Index           =   4
            Left            =   600
            TabIndex        =   256
            Top             =   2040
            Width           =   735
         End
         Begin VB.Label Label92 
            Caption         =   "����"
            Height          =   255
            Index           =   3
            Left            =   600
            TabIndex        =   255
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label92 
            Caption         =   "��ʹ"
            Height          =   255
            Index           =   2
            Left            =   600
            TabIndex        =   254
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label92 
            Caption         =   "�����½�"
            Height          =   255
            Index           =   1
            Left            =   600
            TabIndex        =   253
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label34 
            Caption         =   "����ʱ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   9000
            TabIndex        =   147
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label35 
            Caption         =   "��   ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   7440
            TabIndex        =   146
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label36 
            Caption         =   "��   Ŀ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   6120
            TabIndex        =   145
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label92 
            Caption         =   "����ģ��"
            Height          =   255
            Index           =   0
            Left            =   600
            TabIndex        =   144
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label36 
            Caption         =   "��   Ŀ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   600
            TabIndex        =   119
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label35 
            Caption         =   "��   ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   1920
            TabIndex        =   118
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label34 
            Caption         =   "����ʱ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   3480
            TabIndex        =   117
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame8 
         Height          =   4095
         Left            =   -74640
         TabIndex        =   112
         Top             =   600
         Width           =   7335
         Begin VB.TextBox Text33 
            Height          =   270
            Index           =   3
            Left            =   3480
            TabIndex        =   252
            Top             =   1680
            Width           =   1695
         End
         Begin VB.TextBox Text33 
            Height          =   270
            Index           =   2
            Left            =   3480
            TabIndex        =   251
            Top             =   1320
            Width           =   1695
         End
         Begin VB.TextBox Text33 
            Height          =   270
            Index           =   1
            Left            =   3480
            TabIndex        =   250
            Top             =   960
            Width           =   1695
         End
         Begin VB.ComboBox Combo40 
            Height          =   300
            Index           =   3
            ItemData        =   "frm֢״ѯ��.frx":0677
            Left            =   1920
            List            =   "frm֢״ѯ��.frx":068A
            TabIndex        =   249
            Text            =   "-"
            Top             =   1680
            Width           =   975
         End
         Begin VB.ComboBox Combo40 
            Height          =   300
            Index           =   2
            ItemData        =   "frm֢״ѯ��.frx":06A1
            Left            =   1920
            List            =   "frm֢״ѯ��.frx":06B4
            TabIndex        =   248
            Text            =   "-"
            Top             =   1320
            Width           =   975
         End
         Begin VB.ComboBox Combo40 
            Height          =   300
            Index           =   1
            ItemData        =   "frm֢״ѯ��.frx":06CB
            Left            =   1920
            List            =   "frm֢״ѯ��.frx":06DE
            TabIndex        =   247
            Text            =   "-"
            Top             =   960
            Width           =   975
         End
         Begin VB.ComboBox Combo40 
            Height          =   300
            Index           =   0
            ItemData        =   "frm֢״ѯ��.frx":06F5
            Left            =   1920
            List            =   "frm֢״ѯ��.frx":0708
            TabIndex        =   140
            Text            =   "-"
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox Text33 
            Height          =   270
            Index           =   0
            Left            =   3480
            TabIndex        =   139
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label Label88 
            Caption         =   "�ؽ���ʹ"
            Height          =   255
            Index           =   3
            Left            =   600
            TabIndex        =   246
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label88 
            Caption         =   "������"
            Height          =   255
            Index           =   2
            Left            =   600
            TabIndex        =   245
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label88 
            Caption         =   "������ʹ"
            Height          =   255
            Index           =   1
            Left            =   600
            TabIndex        =   244
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label88 
            Caption         =   "ȫ����ʹ"
            Height          =   255
            Index           =   0
            Left            =   600
            TabIndex        =   141
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label33 
            Caption         =   "��   Ŀ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   115
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label32 
            Caption         =   "��   ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1920
            TabIndex        =   114
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label31 
            Caption         =   "����ʱ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3480
            TabIndex        =   113
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame7 
         Height          =   4095
         Left            =   -74640
         TabIndex        =   108
         Top             =   600
         Width           =   7335
         Begin VB.TextBox Text28 
            Height          =   270
            Index           =   4
            Left            =   3480
            TabIndex        =   243
            Top             =   2040
            Width           =   1695
         End
         Begin VB.TextBox Text28 
            Height          =   270
            Index           =   3
            Left            =   3480
            TabIndex        =   242
            Top             =   1680
            Width           =   1695
         End
         Begin VB.TextBox Text28 
            Height          =   270
            Index           =   2
            Left            =   3480
            TabIndex        =   241
            Top             =   1320
            Width           =   1695
         End
         Begin VB.TextBox Text28 
            Height          =   270
            Index           =   1
            Left            =   3480
            TabIndex        =   240
            Top             =   960
            Width           =   1695
         End
         Begin VB.ComboBox Combo35 
            Height          =   300
            Index           =   4
            ItemData        =   "frm֢״ѯ��.frx":071F
            Left            =   1920
            List            =   "frm֢״ѯ��.frx":0732
            TabIndex        =   239
            Text            =   "-"
            Top             =   2040
            Width           =   975
         End
         Begin VB.ComboBox Combo35 
            Height          =   300
            Index           =   3
            ItemData        =   "frm֢״ѯ��.frx":0749
            Left            =   1920
            List            =   "frm֢״ѯ��.frx":075C
            TabIndex        =   238
            Text            =   "-"
            Top             =   1680
            Width           =   975
         End
         Begin VB.ComboBox Combo35 
            Height          =   300
            Index           =   2
            ItemData        =   "frm֢״ѯ��.frx":0773
            Left            =   1920
            List            =   "frm֢״ѯ��.frx":0786
            TabIndex        =   237
            Text            =   "-"
            Top             =   1320
            Width           =   975
         End
         Begin VB.ComboBox Combo35 
            Height          =   300
            Index           =   1
            ItemData        =   "frm֢״ѯ��.frx":079D
            Left            =   1920
            List            =   "frm֢״ѯ��.frx":07B0
            TabIndex        =   236
            Text            =   "-"
            Top             =   960
            Width           =   975
         End
         Begin VB.ComboBox Combo35 
            Height          =   300
            Index           =   0
            ItemData        =   "frm֢״ѯ��.frx":07C7
            Left            =   1920
            List            =   "frm֢״ѯ��.frx":07DA
            TabIndex        =   137
            Text            =   "-"
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox Text28 
            Height          =   270
            Index           =   0
            Left            =   3480
            TabIndex        =   136
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label Label83 
            Caption         =   "��������"
            Height          =   255
            Index           =   4
            Left            =   600
            TabIndex        =   235
            Top             =   2040
            Width           =   975
         End
         Begin VB.Label Label83 
            Caption         =   "ˮ��"
            Height          =   255
            Index           =   3
            Left            =   600
            TabIndex        =   234
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label Label83 
            Caption         =   "��ʹ"
            Height          =   255
            Index           =   2
            Left            =   600
            TabIndex        =   233
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label Label83 
            Caption         =   "Ѫ��"
            Height          =   255
            Index           =   1
            Left            =   600
            TabIndex        =   232
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label83 
            Caption         =   "��Ƶ����"
            Height          =   255
            Index           =   0
            Left            =   600
            TabIndex        =   138
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label30 
            Caption         =   "��   Ŀ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   111
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label29 
            Caption         =   "��   ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1920
            TabIndex        =   110
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label28 
            Caption         =   "����ʱ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3480
            TabIndex        =   109
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame6 
         Height          =   4095
         Left            =   -74640
         TabIndex        =   104
         Top             =   600
         Width           =   7335
         Begin VB.TextBox Text20 
            Height          =   270
            Index           =   7
            Left            =   3480
            TabIndex        =   231
            Top             =   3120
            Width           =   1695
         End
         Begin VB.TextBox Text20 
            Height          =   270
            Index           =   6
            Left            =   3480
            TabIndex        =   230
            Top             =   2760
            Width           =   1695
         End
         Begin VB.TextBox Text20 
            Height          =   270
            Index           =   5
            Left            =   3480
            TabIndex        =   229
            Top             =   2400
            Width           =   1695
         End
         Begin VB.TextBox Text20 
            Height          =   270
            Index           =   4
            Left            =   3480
            TabIndex        =   228
            Top             =   2040
            Width           =   1695
         End
         Begin VB.TextBox Text20 
            Height          =   270
            Index           =   3
            Left            =   3480
            TabIndex        =   227
            Top             =   1680
            Width           =   1695
         End
         Begin VB.TextBox Text20 
            Height          =   270
            Index           =   2
            Left            =   3480
            TabIndex        =   226
            Top             =   1320
            Width           =   1695
         End
         Begin VB.TextBox Text20 
            Height          =   270
            Index           =   1
            Left            =   3480
            TabIndex        =   225
            Top             =   960
            Width           =   1695
         End
         Begin VB.ComboBox Combo27 
            Height          =   300
            Index           =   7
            ItemData        =   "frm֢״ѯ��.frx":07F1
            Left            =   1920
            List            =   "frm֢״ѯ��.frx":0804
            TabIndex        =   224
            Text            =   "-"
            Top             =   3120
            Width           =   975
         End
         Begin VB.ComboBox Combo27 
            Height          =   300
            Index           =   6
            ItemData        =   "frm֢״ѯ��.frx":081B
            Left            =   1920
            List            =   "frm֢״ѯ��.frx":082E
            TabIndex        =   223
            Text            =   "-"
            Top             =   2760
            Width           =   975
         End
         Begin VB.ComboBox Combo27 
            Height          =   300
            Index           =   5
            ItemData        =   "frm֢״ѯ��.frx":0845
            Left            =   1920
            List            =   "frm֢״ѯ��.frx":0858
            TabIndex        =   222
            Text            =   "-"
            Top             =   2400
            Width           =   975
         End
         Begin VB.ComboBox Combo27 
            Height          =   300
            Index           =   4
            ItemData        =   "frm֢״ѯ��.frx":086F
            Left            =   1920
            List            =   "frm֢״ѯ��.frx":0882
            TabIndex        =   221
            Text            =   "-"
            Top             =   2040
            Width           =   975
         End
         Begin VB.ComboBox Combo27 
            Height          =   300
            Index           =   3
            ItemData        =   "frm֢״ѯ��.frx":0899
            Left            =   1920
            List            =   "frm֢״ѯ��.frx":08AC
            TabIndex        =   220
            Text            =   "-"
            Top             =   1680
            Width           =   975
         End
         Begin VB.ComboBox Combo27 
            Height          =   300
            Index           =   2
            ItemData        =   "frm֢״ѯ��.frx":08C3
            Left            =   1920
            List            =   "frm֢״ѯ��.frx":08D6
            TabIndex        =   219
            Text            =   "-"
            Top             =   1320
            Width           =   975
         End
         Begin VB.ComboBox Combo27 
            Height          =   300
            Index           =   1
            ItemData        =   "frm֢״ѯ��.frx":08ED
            Left            =   1920
            List            =   "frm֢״ѯ��.frx":0900
            TabIndex        =   218
            Text            =   "-"
            Top             =   960
            Width           =   975
         End
         Begin VB.ComboBox Combo27 
            Height          =   300
            Index           =   0
            ItemData        =   "frm֢״ѯ��.frx":0917
            Left            =   1920
            List            =   "frm֢״ѯ��.frx":092A
            TabIndex        =   134
            Text            =   "-"
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox Text20 
            Height          =   270
            Index           =   0
            Left            =   3480
            TabIndex        =   133
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label Label75 
            Caption         =   "����"
            Height          =   255
            Index           =   7
            Left            =   600
            TabIndex        =   217
            Top             =   3120
            Width           =   735
         End
         Begin VB.Label Label75 
            Caption         =   "��Ѫ"
            Height          =   255
            Index           =   6
            Left            =   600
            TabIndex        =   216
            Top             =   2760
            Width           =   735
         End
         Begin VB.Label Label75 
            Caption         =   "��̵"
            Height          =   255
            Index           =   5
            Left            =   600
            TabIndex        =   215
            Top             =   2400
            Width           =   735
         End
         Begin VB.Label Label75 
            Caption         =   "����"
            Height          =   255
            Index           =   4
            Left            =   600
            TabIndex        =   214
            Top             =   2040
            Width           =   735
         End
         Begin VB.Label Label75 
            Caption         =   "��ʹ"
            Height          =   255
            Index           =   3
            Left            =   600
            TabIndex        =   213
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label75 
            Caption         =   "����"
            Height          =   255
            Index           =   2
            Left            =   600
            TabIndex        =   212
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label75 
            Caption         =   "����"
            Height          =   255
            Index           =   1
            Left            =   600
            TabIndex        =   211
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label75 
            Caption         =   "����"
            Height          =   255
            Index           =   0
            Left            =   600
            TabIndex        =   135
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label27 
            Caption         =   "��   Ŀ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   107
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label26 
            Caption         =   "��   ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1920
            TabIndex        =   106
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label25 
            Caption         =   "����ʱ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3480
            TabIndex        =   105
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame5 
         Height          =   4575
         Left            =   -74640
         TabIndex        =   100
         Top             =   600
         Width           =   7335
         Begin VB.TextBox Text12 
            Height          =   270
            Index           =   7
            Left            =   3480
            TabIndex        =   210
            Top             =   3120
            Width           =   1695
         End
         Begin VB.TextBox Text12 
            Height          =   270
            Index           =   6
            Left            =   3480
            TabIndex        =   209
            Top             =   2760
            Width           =   1695
         End
         Begin VB.TextBox Text12 
            Height          =   270
            Index           =   5
            Left            =   3480
            TabIndex        =   208
            Top             =   2400
            Width           =   1695
         End
         Begin VB.TextBox Text12 
            Height          =   270
            Index           =   4
            Left            =   3480
            TabIndex        =   207
            Top             =   2040
            Width           =   1695
         End
         Begin VB.TextBox Text12 
            Height          =   270
            Index           =   3
            Left            =   3480
            TabIndex        =   206
            Top             =   1680
            Width           =   1695
         End
         Begin VB.TextBox Text12 
            Height          =   270
            Index           =   2
            Left            =   3480
            TabIndex        =   205
            Top             =   1320
            Width           =   1695
         End
         Begin VB.TextBox Text12 
            Height          =   270
            Index           =   1
            Left            =   3480
            TabIndex        =   204
            Top             =   960
            Width           =   1695
         End
         Begin VB.ComboBox Combo20 
            Height          =   300
            Index           =   7
            ItemData        =   "frm֢״ѯ��.frx":0941
            Left            =   1920
            List            =   "frm֢״ѯ��.frx":0954
            TabIndex        =   203
            Text            =   "-"
            Top             =   3120
            Width           =   975
         End
         Begin VB.ComboBox Combo20 
            Height          =   300
            Index           =   6
            ItemData        =   "frm֢״ѯ��.frx":096B
            Left            =   1920
            List            =   "frm֢״ѯ��.frx":097E
            TabIndex        =   202
            Text            =   "-"
            Top             =   2760
            Width           =   975
         End
         Begin VB.ComboBox Combo20 
            Height          =   300
            Index           =   5
            ItemData        =   "frm֢״ѯ��.frx":0995
            Left            =   1920
            List            =   "frm֢״ѯ��.frx":09A8
            TabIndex        =   201
            Text            =   "-"
            Top             =   2400
            Width           =   975
         End
         Begin VB.ComboBox Combo20 
            Height          =   300
            Index           =   4
            ItemData        =   "frm֢״ѯ��.frx":09BF
            Left            =   1920
            List            =   "frm֢״ѯ��.frx":09D2
            TabIndex        =   200
            Text            =   "-"
            Top             =   2040
            Width           =   975
         End
         Begin VB.ComboBox Combo20 
            Height          =   300
            Index           =   3
            ItemData        =   "frm֢״ѯ��.frx":09E9
            Left            =   1920
            List            =   "frm֢״ѯ��.frx":09FC
            TabIndex        =   199
            Text            =   "-"
            Top             =   1680
            Width           =   975
         End
         Begin VB.ComboBox Combo20 
            Height          =   300
            Index           =   2
            ItemData        =   "frm֢״ѯ��.frx":0A13
            Left            =   1920
            List            =   "frm֢״ѯ��.frx":0A26
            TabIndex        =   198
            Text            =   "-"
            Top             =   1320
            Width           =   975
         End
         Begin VB.ComboBox Combo20 
            Height          =   300
            Index           =   1
            ItemData        =   "frm֢״ѯ��.frx":0A3D
            Left            =   1920
            List            =   "frm֢״ѯ��.frx":0A50
            TabIndex        =   197
            Text            =   "-"
            Top             =   960
            Width           =   975
         End
         Begin VB.ComboBox Combo20 
            Height          =   300
            Index           =   0
            ItemData        =   "frm֢״ѯ��.frx":0A67
            Left            =   1920
            List            =   "frm֢״ѯ��.frx":0A7A
            TabIndex        =   131
            Text            =   "-"
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox Text12 
            Height          =   270
            Index           =   0
            Left            =   3480
            TabIndex        =   130
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label Label70 
            Caption         =   "����"
            Height          =   255
            Index           =   7
            Left            =   600
            TabIndex        =   196
            Top             =   3120
            Width           =   735
         End
         Begin VB.Label Label70 
            Caption         =   "�ѷ�"
            Height          =   255
            Index           =   6
            Left            =   600
            TabIndex        =   195
            Top             =   2760
            Width           =   735
         End
         Begin VB.Label Label70 
            Caption         =   "����"
            Height          =   255
            Index           =   5
            Left            =   600
            TabIndex        =   194
            Top             =   2400
            Width           =   735
         End
         Begin VB.Label Label70 
            Caption         =   "Ƥ��"
            Height          =   255
            Index           =   4
            Left            =   600
            TabIndex        =   193
            Top             =   2040
            Width           =   735
         End
         Begin VB.Label Label70 
            Caption         =   "Ƥ������"
            Height          =   255
            Index           =   3
            Left            =   600
            TabIndex        =   192
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label70 
            Caption         =   "�ڿ�"
            Height          =   255
            Index           =   2
            Left            =   600
            TabIndex        =   191
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label70 
            Caption         =   "����"
            Height          =   255
            Index           =   1
            Left            =   600
            TabIndex        =   190
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label70 
            Caption         =   "�¾��쳣"
            Height          =   255
            Index           =   0
            Left            =   600
            TabIndex        =   132
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label23 
            Caption         =   "��   Ŀ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   103
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label6 
            Caption         =   "��   ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1920
            TabIndex        =   102
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "����ʱ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3480
            TabIndex        =   101
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame1 
         Height          =   4575
         Left            =   360
         TabIndex        =   96
         Top             =   600
         Width           =   7335
         Begin VB.ComboBox Combo1 
            Height          =   300
            Index           =   10
            ItemData        =   "frm֢״ѯ��.frx":0A91
            Left            =   1920
            List            =   "frm֢״ѯ��.frx":0AA4
            TabIndex        =   178
            Text            =   "-"
            Top             =   4200
            Width           =   975
         End
         Begin VB.ComboBox Combo1 
            Height          =   300
            Index           =   9
            ItemData        =   "frm֢״ѯ��.frx":0ABB
            Left            =   1920
            List            =   "frm֢״ѯ��.frx":0ACE
            TabIndex        =   177
            Text            =   "-"
            Top             =   3840
            Width           =   975
         End
         Begin VB.ComboBox Combo1 
            Height          =   300
            Index           =   8
            ItemData        =   "frm֢״ѯ��.frx":0AE5
            Left            =   1920
            List            =   "frm֢״ѯ��.frx":0AF8
            TabIndex        =   176
            Text            =   "-"
            Top             =   3480
            Width           =   975
         End
         Begin VB.ComboBox Combo1 
            Height          =   300
            Index           =   7
            ItemData        =   "frm֢״ѯ��.frx":0B0F
            Left            =   1920
            List            =   "frm֢״ѯ��.frx":0B22
            TabIndex        =   175
            Text            =   "-"
            Top             =   3120
            Width           =   975
         End
         Begin VB.ComboBox Combo1 
            Height          =   300
            Index           =   6
            ItemData        =   "frm֢״ѯ��.frx":0B39
            Left            =   1920
            List            =   "frm֢״ѯ��.frx":0B4C
            TabIndex        =   174
            Text            =   "-"
            Top             =   2760
            Width           =   975
         End
         Begin VB.ComboBox Combo1 
            Height          =   300
            Index           =   5
            ItemData        =   "frm֢״ѯ��.frx":0B63
            Left            =   1920
            List            =   "frm֢״ѯ��.frx":0B76
            TabIndex        =   173
            Text            =   "-"
            Top             =   2400
            Width           =   975
         End
         Begin VB.ComboBox Combo1 
            Height          =   300
            Index           =   4
            ItemData        =   "frm֢״ѯ��.frx":0B8D
            Left            =   1920
            List            =   "frm֢״ѯ��.frx":0BA0
            TabIndex        =   172
            Text            =   "-"
            Top             =   2040
            Width           =   975
         End
         Begin VB.ComboBox Combo1 
            Height          =   300
            Index           =   3
            ItemData        =   "frm֢״ѯ��.frx":0BB7
            Left            =   1920
            List            =   "frm֢״ѯ��.frx":0BCA
            TabIndex        =   171
            Text            =   "-"
            Top             =   1680
            Width           =   975
         End
         Begin VB.ComboBox Combo1 
            Height          =   300
            Index           =   2
            ItemData        =   "frm֢״ѯ��.frx":0BE1
            Left            =   1920
            List            =   "frm֢״ѯ��.frx":0BF4
            TabIndex        =   170
            Text            =   "-"
            Top             =   1320
            Width           =   975
         End
         Begin VB.ComboBox Combo1 
            Height          =   300
            Index           =   1
            ItemData        =   "frm֢״ѯ��.frx":0C0B
            Left            =   1920
            List            =   "frm֢״ѯ��.frx":0C1E
            TabIndex        =   169
            Text            =   "-"
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   10
            Left            =   3480
            TabIndex        =   168
            Top             =   4200
            Width           =   1695
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   9
            Left            =   3480
            TabIndex        =   167
            Top             =   3840
            Width           =   1695
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   8
            Left            =   3480
            TabIndex        =   166
            Top             =   3480
            Width           =   1695
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   7
            Left            =   3480
            TabIndex        =   165
            Top             =   3120
            Width           =   1695
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   6
            Left            =   3480
            TabIndex        =   164
            Top             =   2760
            Width           =   1695
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   5
            Left            =   3480
            TabIndex        =   163
            Top             =   2400
            Width           =   1695
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   4
            Left            =   3480
            TabIndex        =   162
            Top             =   2040
            Width           =   1695
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   3
            Left            =   3480
            TabIndex        =   161
            Top             =   1680
            Width           =   1695
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   2
            Left            =   3480
            TabIndex        =   160
            Top             =   1320
            Width           =   1695
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   1
            Left            =   3480
            TabIndex        =   159
            Top             =   960
            Width           =   1695
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   0
            Left            =   3480
            TabIndex        =   129
            Top             =   600
            Width           =   1695
         End
         Begin VB.ComboBox Combo1 
            Height          =   300
            Index           =   0
            ItemData        =   "frm֢״ѯ��.frx":0C35
            Left            =   1920
            List            =   "frm֢״ѯ��.frx":0C48
            TabIndex        =   128
            Text            =   "-"
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label46 
            Caption         =   "���������"
            Height          =   255
            Index           =   10
            Left            =   600
            TabIndex        =   189
            Top             =   4200
            Width           =   975
         End
         Begin VB.Label Label46 
            Caption         =   "��֫��ľ"
            Height          =   255
            Index           =   9
            Left            =   600
            TabIndex        =   188
            Top             =   3840
            Width           =   735
         End
         Begin VB.Label Label46 
            Caption         =   "ƣ������"
            Height          =   255
            Index           =   8
            Left            =   600
            TabIndex        =   187
            Top             =   3480
            Width           =   735
         End
         Begin VB.Label Label46 
            Caption         =   "�׼���"
            Height          =   255
            Index           =   7
            Left            =   600
            TabIndex        =   186
            Top             =   3120
            Width           =   735
         End
         Begin VB.Label Label46 
            Caption         =   "����������"
            Height          =   255
            Index           =   6
            Left            =   600
            TabIndex        =   185
            Top             =   2760
            Width           =   1095
         End
         Begin VB.Label Label46 
            Caption         =   "����"
            Height          =   255
            Index           =   5
            Left            =   600
            TabIndex        =   184
            Top             =   2400
            Width           =   735
         End
         Begin VB.Label Label46 
            Caption         =   "��˯"
            Height          =   255
            Index           =   4
            Left            =   600
            TabIndex        =   183
            Top             =   2040
            Width           =   735
         End
         Begin VB.Label Label46 
            Caption         =   "ʧ��"
            Height          =   255
            Index           =   3
            Left            =   600
            TabIndex        =   182
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label46 
            Caption         =   "ѣ��"
            Height          =   255
            Index           =   2
            Left            =   600
            TabIndex        =   181
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label46 
            Caption         =   "ͷ(��)��"
            Height          =   255
            Index           =   1
            Left            =   600
            TabIndex        =   180
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label46 
            Caption         =   "ͷʹ"
            Height          =   255
            Index           =   0
            Left            =   600
            TabIndex        =   179
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label4 
            Caption         =   "����ʱ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3480
            TabIndex        =   99
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "��   ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1920
            TabIndex        =   98
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "��   Ŀ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   97
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame freNuclear 
         BorderStyle     =   0  'None
         Height          =   3735
         Left            =   120
         TabIndex        =   95
         Top             =   960
         Width           =   11175
      End
      Begin VB.Frame freOrdinary 
         BorderStyle     =   0  'None
         Height          =   3735
         Left            =   120
         TabIndex        =   58
         Top             =   960
         Width           =   11175
         Begin VB.Frame Frame4 
            Caption         =   "����ʷ"
            ForeColor       =   &H000080FF&
            Height          =   1815
            Index           =   1
            Left            =   6000
            TabIndex        =   93
            Top             =   1800
            Width           =   5055
            Begin VB.TextBox ctxt����ʷ 
               Height          =   1455
               Index           =   2
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   94
               Top             =   240
               Width           =   4695
            End
         End
         Begin VB.Frame Frame19 
            Caption         =   "�̾�ʷ"
            ForeColor       =   &H000080FF&
            Height          =   1455
            Index           =   1
            Left            =   6000
            TabIndex        =   73
            Top             =   240
            Width           =   5055
            Begin VB.TextBox ctxt���� 
               Height          =   270
               Index           =   2
               Left            =   960
               TabIndex        =   80
               Top             =   800
               Width           =   975
            End
            Begin VB.TextBox ctxt���� 
               Height          =   270
               Index           =   2
               Left            =   960
               TabIndex        =   79
               Top             =   480
               Width           =   975
            End
            Begin VB.TextBox ctxt���� 
               Height          =   270
               Index           =   2
               Left            =   3360
               TabIndex        =   78
               Top             =   480
               Width           =   975
            End
            Begin VB.TextBox ctxt������ 
               Height          =   270
               Index           =   2
               Left            =   3360
               TabIndex        =   77
               Top             =   800
               Width           =   975
            End
            Begin VB.TextBox ctxt������ 
               Height          =   270
               Index           =   2
               Left            =   960
               TabIndex        =   76
               Top             =   1120
               Width           =   975
            End
            Begin VB.ComboBox ccmb���� 
               Height          =   300
               Index           =   2
               Left            =   3360
               TabIndex        =   75
               Top             =   120
               Width           =   1335
            End
            Begin VB.ComboBox ccmb���� 
               Height          =   300
               Index           =   2
               Left            =   960
               TabIndex        =   74
               Top             =   120
               Width           =   1335
            End
            Begin VB.Label Label110 
               AutoSize        =   -1  'True
               Caption         =   "��"
               Height          =   180
               Index           =   1
               Left            =   2040
               TabIndex        =   92
               Top             =   840
               Width           =   180
            End
            Begin VB.Label Label109 
               AutoSize        =   -1  'True
               Caption         =   "��"
               Height          =   180
               Index           =   1
               Left            =   4440
               TabIndex        =   91
               Top             =   480
               Width           =   180
            End
            Begin VB.Label Label108 
               AutoSize        =   -1  'True
               Caption         =   "��"
               Height          =   180
               Index           =   1
               Left            =   2040
               TabIndex        =   90
               Top             =   480
               Width           =   180
            End
            Begin VB.Label Label107 
               AutoSize        =   -1  'True
               Caption         =   "����ʱ����"
               Height          =   180
               Index           =   1
               Left            =   120
               TabIndex        =   89
               Top             =   860
               Width           =   900
            End
            Begin VB.Label Label106 
               AutoSize        =   -1  'True
               Caption         =   "���䣺"
               Height          =   180
               Index           =   1
               Left            =   360
               TabIndex        =   88
               Top             =   540
               Width           =   540
            End
            Begin VB.Label Label105 
               AutoSize        =   -1  'True
               Caption         =   "���䣺"
               Height          =   180
               Index           =   1
               Left            =   2640
               TabIndex        =   87
               Top             =   540
               Width           =   540
            End
            Begin VB.Label Label104 
               AutoSize        =   -1  'True
               Caption         =   "��������"
               Height          =   180
               Index           =   1
               Left            =   2640
               TabIndex        =   86
               Top             =   860
               Width           =   720
            End
            Begin VB.Label Label103 
               AutoSize        =   -1  'True
               Caption         =   "��������"
               Height          =   180
               Index           =   1
               Left            =   240
               TabIndex        =   85
               Top             =   1155
               Width           =   720
            End
            Begin VB.Label Label102 
               AutoSize        =   -1  'True
               Caption         =   "֧/��"
               Height          =   180
               Index           =   1
               Left            =   4440
               TabIndex        =   84
               Top             =   840
               Width           =   450
            End
            Begin VB.Label Label101 
               AutoSize        =   -1  'True
               Caption         =   "ML/��"
               Height          =   180
               Index           =   1
               Left            =   2040
               TabIndex        =   83
               Top             =   1120
               Width           =   450
            End
            Begin VB.Label Label100 
               AutoSize        =   -1  'True
               Caption         =   "���Ƴ̶ȣ�"
               Height          =   180
               Index           =   1
               Left            =   2520
               TabIndex        =   82
               Top             =   195
               Width           =   900
            End
            Begin VB.Label Label99 
               AutoSize        =   -1  'True
               Caption         =   "���̶̳ȣ�"
               Height          =   180
               Index           =   1
               Left            =   120
               TabIndex        =   81
               Top             =   195
               Width           =   900
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "����ʷ"
            ForeColor       =   &H000080FF&
            Height          =   615
            Index           =   2
            Left            =   120
            TabIndex        =   70
            Top             =   240
            Width           =   5775
            Begin VB.ComboBox Ccmb��� 
               Height          =   300
               Index           =   2
               Left            =   1680
               TabIndex        =   71
               Top             =   240
               Width           =   1695
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "�Ƿ��飺"
               Height          =   180
               Index           =   2
               Left            =   480
               TabIndex        =   72
               Top             =   300
               Width           =   900
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "����ʷ(����ż����ʷ)"
            ForeColor       =   &H000080FF&
            Height          =   2655
            Index           =   2
            Left            =   120
            TabIndex        =   59
            Top             =   960
            Width           =   5775
            Begin VB.TextBox ctxt���� 
               Height          =   270
               Index           =   2
               Left            =   4200
               TabIndex        =   64
               Text            =   "0"
               Top             =   720
               Width           =   975
            End
            Begin VB.TextBox ctxt��� 
               Height          =   270
               Index           =   2
               Left            =   4200
               TabIndex        =   63
               Text            =   "0"
               Top             =   360
               Width           =   975
            End
            Begin VB.TextBox ctxt���� 
               Height          =   270
               Index           =   2
               Left            =   1680
               TabIndex        =   62
               Text            =   "0"
               Top             =   720
               Width           =   1215
            End
            Begin VB.TextBox ctxt������Ů 
               Height          =   270
               Index           =   2
               Left            =   1680
               TabIndex        =   61
               Text            =   "0"
               Top             =   360
               Width           =   1215
            End
            Begin VB.TextBox ctxt�쳣̥ 
               Height          =   270
               Left            =   1680
               TabIndex        =   60
               Text            =   "0"
               Top             =   1080
               Width           =   1215
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "������"
               Height          =   180
               Index           =   1
               Left            =   3480
               TabIndex        =   69
               Top             =   720
               Width           =   540
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "�����"
               Height          =   180
               Index           =   1
               Left            =   3480
               TabIndex        =   68
               Top             =   360
               Width           =   540
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               Caption         =   "��Ȼ������"
               Height          =   180
               Index           =   1
               Left            =   480
               TabIndex        =   67
               Top             =   720
               Width           =   900
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               Caption         =   "������Ů��Ŀ��"
               Height          =   180
               Index           =   1
               Left            =   480
               TabIndex        =   66
               Top             =   360
               Width           =   1260
            End
            Begin VB.Label Label�쳣̥ 
               AutoSize        =   -1  'True
               Caption         =   "�쳣̥��"
               Height          =   180
               Left            =   480
               TabIndex        =   65
               Top             =   1080
               Width           =   720
            End
         End
      End
      Begin VB.Frame freRadiation 
         BorderStyle     =   0  'None
         Height          =   3735
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   11175
         Begin VB.Frame Frame3 
            Caption         =   "����ʷ(����ż����ʷ)"
            ForeColor       =   &H000080FF&
            Height          =   2175
            Index           =   0
            Left            =   120
            TabIndex        =   35
            Top             =   1440
            Width           =   5775
            Begin VB.TextBox ctxt��Ů���� 
               Height          =   270
               Index           =   0
               Left            =   1560
               TabIndex        =   46
               Top             =   1080
               Width           =   1335
            End
            Begin VB.TextBox ctxt������Ů 
               Height          =   270
               Index           =   0
               Left            =   240
               TabIndex        =   45
               Text            =   "0"
               Top             =   1080
               Width           =   1215
            End
            Begin VB.TextBox ctxt���� 
               Height          =   270
               Index           =   0
               Left            =   4680
               TabIndex        =   44
               Text            =   "0"
               Top             =   480
               Width           =   855
            End
            Begin VB.TextBox ctxt��̥ 
               Height          =   270
               Index           =   0
               Left            =   240
               TabIndex        =   43
               Text            =   "0"
               Top             =   1680
               Width           =   1215
            End
            Begin VB.TextBox ctxt��̥ 
               Height          =   270
               Index           =   0
               Left            =   2760
               TabIndex        =   42
               Text            =   "0"
               Top             =   480
               Width           =   975
            End
            Begin VB.TextBox ctxt�д� 
               Height          =   270
               Index           =   0
               Left            =   1920
               TabIndex        =   41
               Text            =   "0"
               Top             =   480
               Width           =   735
            End
            Begin VB.TextBox ctxt��λ���� 
               Height          =   270
               Index           =   0
               Left            =   1560
               TabIndex        =   40
               Text            =   "0"
               Top             =   1680
               Width           =   855
            End
            Begin VB.TextBox ctxt��� 
               Height          =   270
               Index           =   0
               Left            =   3840
               TabIndex        =   39
               Text            =   "0"
               Top             =   480
               Width           =   735
            End
            Begin VB.TextBox ctxt��� 
               Height          =   270
               Index           =   0
               Left            =   240
               TabIndex        =   38
               Text            =   "0"
               Top             =   480
               Width           =   735
            End
            Begin VB.TextBox ctxt���� 
               Height          =   270
               Index           =   0
               Left            =   1080
               TabIndex        =   37
               Text            =   "0"
               Top             =   480
               Width           =   735
            End
            Begin VB.TextBox ctxt���в��� 
               Height          =   855
               Index           =   0
               Left            =   3000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   36
               Top             =   1080
               Width           =   2535
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               Caption         =   "��Ů����״����"
               Height          =   180
               Index           =   0
               Left            =   1560
               TabIndex        =   57
               Top             =   840
               Width           =   1260
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               Caption         =   "������Ů��Ŀ��"
               Height          =   180
               Index           =   0
               Left            =   240
               TabIndex        =   56
               Top             =   840
               Width           =   1260
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               Caption         =   "��Ȼ������"
               Height          =   180
               Index           =   0
               Left            =   4680
               TabIndex        =   55
               Top             =   240
               Width           =   900
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               Caption         =   "��̥��"
               Height          =   180
               Index           =   0
               Left            =   240
               TabIndex        =   54
               Top             =   1440
               Width           =   540
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               Caption         =   "��̥��"
               Height          =   180
               Index           =   0
               Left            =   2760
               TabIndex        =   53
               Top             =   240
               Width           =   540
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "�дΣ�"
               Height          =   180
               Index           =   0
               Left            =   1920
               TabIndex        =   52
               Top             =   240
               Width           =   540
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               Caption         =   "��λ���"
               Height          =   180
               Index           =   0
               Left            =   1560
               TabIndex        =   51
               Top             =   1440
               Width           =   900
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "�����"
               Height          =   180
               Index           =   0
               Left            =   3840
               TabIndex        =   50
               Top             =   240
               Width           =   540
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "�����"
               Height          =   180
               Index           =   0
               Left            =   240
               TabIndex        =   49
               Top             =   240
               Width           =   540
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "������"
               Height          =   180
               Index           =   0
               Left            =   1080
               TabIndex        =   48
               Top             =   240
               Width           =   540
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               Caption         =   "���в���ԭ��"
               Height          =   180
               Index           =   0
               Left            =   3000
               TabIndex        =   47
               Top             =   840
               Width           =   1260
            End
         End
         Begin VB.Frame Frame19 
            Caption         =   "�̾�ʷ"
            ForeColor       =   &H000080FF&
            Height          =   2175
            Index           =   0
            Left            =   6000
            TabIndex        =   13
            Top             =   1440
            Width           =   5055
            Begin VB.ComboBox ccmb���� 
               Height          =   300
               Index           =   0
               Left            =   960
               TabIndex        =   21
               Top             =   120
               Width           =   1335
            End
            Begin VB.ComboBox ccmb���� 
               Height          =   300
               Index           =   0
               Left            =   3360
               TabIndex        =   20
               Top             =   120
               Width           =   1335
            End
            Begin VB.TextBox ctxt������ 
               Height          =   270
               Index           =   0
               Left            =   3360
               TabIndex        =   19
               Top             =   840
               Width           =   975
            End
            Begin VB.TextBox ctxt������ 
               Height          =   270
               Index           =   0
               Left            =   960
               TabIndex        =   18
               Top             =   1155
               Width           =   975
            End
            Begin VB.TextBox ctxt���� 
               Height          =   270
               Index           =   0
               Left            =   3360
               TabIndex        =   17
               Top             =   480
               Width           =   975
            End
            Begin VB.TextBox ctxt���� 
               Height          =   270
               Index           =   0
               Left            =   960
               TabIndex        =   16
               Top             =   480
               Width           =   975
            End
            Begin VB.TextBox ctxt���� 
               Height          =   270
               Index           =   0
               Left            =   960
               TabIndex        =   15
               Top             =   795
               Width           =   975
            End
            Begin VB.TextBox ctxtMore 
               Height          =   375
               Index           =   0
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   14
               Top             =   1680
               Width           =   4695
            End
            Begin VB.Label Label99 
               AutoSize        =   -1  'True
               Caption         =   "���̶̳ȣ�"
               Height          =   180
               Index           =   0
               Left            =   120
               TabIndex        =   34
               Top             =   195
               Width           =   900
            End
            Begin VB.Label Label100 
               AutoSize        =   -1  'True
               Caption         =   "���Ƴ̶ȣ�"
               Height          =   180
               Index           =   0
               Left            =   2520
               TabIndex        =   33
               Top             =   195
               Width           =   900
            End
            Begin VB.Label Label101 
               AutoSize        =   -1  'True
               Caption         =   "ML/��"
               Height          =   180
               Index           =   0
               Left            =   4440
               TabIndex        =   32
               Top             =   880
               Width           =   450
            End
            Begin VB.Label Label102 
               AutoSize        =   -1  'True
               Caption         =   "֧/��"
               Height          =   180
               Index           =   0
               Left            =   2040
               TabIndex        =   31
               Top             =   1200
               Width           =   450
            End
            Begin VB.Label Label103 
               AutoSize        =   -1  'True
               Caption         =   "��������"
               Height          =   180
               Index           =   0
               Left            =   2640
               TabIndex        =   30
               Top             =   915
               Width           =   720
            End
            Begin VB.Label Label104 
               AutoSize        =   -1  'True
               Caption         =   "��������"
               Height          =   180
               Index           =   0
               Left            =   240
               TabIndex        =   29
               Top             =   1215
               Width           =   720
            End
            Begin VB.Label Label105 
               AutoSize        =   -1  'True
               Caption         =   "���䣺"
               Height          =   180
               Index           =   0
               Left            =   2640
               TabIndex        =   28
               Top             =   540
               Width           =   540
            End
            Begin VB.Label Label106 
               AutoSize        =   -1  'True
               Caption         =   "���䣺"
               Height          =   180
               Index           =   0
               Left            =   360
               TabIndex        =   27
               Top             =   540
               Width           =   540
            End
            Begin VB.Label Label107 
               AutoSize        =   -1  'True
               Caption         =   "����ʱ����"
               Height          =   180
               Index           =   0
               Left            =   120
               TabIndex        =   26
               Top             =   855
               Width           =   900
            End
            Begin VB.Label Label108 
               AutoSize        =   -1  'True
               Caption         =   "��"
               Height          =   180
               Index           =   0
               Left            =   2040
               TabIndex        =   25
               Top             =   480
               Width           =   180
            End
            Begin VB.Label Label109 
               AutoSize        =   -1  'True
               Caption         =   "��"
               Height          =   180
               Index           =   0
               Left            =   4440
               TabIndex        =   24
               Top             =   480
               Width           =   180
            End
            Begin VB.Label Label110 
               AutoSize        =   -1  'True
               Caption         =   "��"
               Height          =   180
               Index           =   0
               Left            =   2040
               TabIndex        =   23
               Top             =   840
               Width           =   180
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               Caption         =   "�����ס��������ʳϰ�ߡ��̾��Ⱥ�������"
               Height          =   180
               Index           =   1
               Left            =   120
               TabIndex        =   22
               Top             =   1440
               Width           =   3420
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "����ʷ"
            ForeColor       =   &H000080FF&
            Height          =   1095
            Index           =   0
            Left            =   120
            TabIndex        =   2
            Top             =   120
            Width           =   10935
            Begin VB.ComboBox Ccmb��� 
               Height          =   300
               Index           =   0
               Left            =   960
               TabIndex        =   6
               Top             =   240
               Width           =   1695
            End
            Begin VB.TextBox ctxtmatejob 
               Height          =   270
               Index           =   0
               Left            =   3960
               TabIndex        =   5
               Top             =   720
               Width           =   1695
            End
            Begin VB.TextBox ctxtmateradioac 
               Height          =   495
               Index           =   0
               Left            =   5880
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   4
               Top             =   480
               Width           =   4815
            End
            Begin VB.TextBox ctxtmatehelh 
               Height          =   270
               Index           =   0
               Left            =   3960
               TabIndex        =   3
               Top             =   240
               Width           =   1695
            End
            Begin MSComCtl2.DTPicker ctxtmarrydate 
               Height          =   300
               Index           =   0
               Left            =   960
               TabIndex        =   7
               Top             =   720
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   529
               _Version        =   393216
               CustomFormat    =   "yyyy/MM/dd"
               Format          =   60096512
               CurrentDate     =   41013
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "�Ƿ��飺"
               Height          =   180
               Index           =   0
               Left            =   120
               TabIndex        =   12
               Top             =   300
               Width           =   900
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "������ڣ�"
               Height          =   180
               Index           =   0
               Left            =   120
               TabIndex        =   11
               Top             =   760
               Width           =   900
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "��żְҵ��"
               Height          =   180
               Index           =   0
               Left            =   2760
               TabIndex        =   10
               Top             =   760
               Width           =   900
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "��ż�Ӵ������������"
               Height          =   180
               Index           =   0
               Left            =   5880
               TabIndex        =   9
               Top             =   240
               Width           =   1800
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "��ż����״����"
               Height          =   180
               Index           =   0
               Left            =   2760
               TabIndex        =   8
               Top             =   300
               Width           =   1260
            End
         End
      End
   End
End
Attribute VB_Name = "frm֢״ѯ��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click(Index As Integer)
If Combo1(Index).Text <> "-" Then
 Text1(Index).Text = "��"
 Else
 Text1(Index).Text = ""
 End If
 
End Sub

Private Sub Combo128_Click(Index As Integer)
If Combo128(Index).Text <> "-" Then
 Text144(Index).Text = "��"
 Else
 Text144(Index).Text = ""
 End If
 
End Sub

Private Sub Combo20_Click(Index As Integer)
If Combo20(Index).Text <> "-" Then
 Text12(Index).Text = "��"
 Else
 Text12(Index).Text = ""
 End If
 
End Sub


Private Sub Combo228_Click(Index As Integer)
If Combo228(Index).Text <> "-" Then
 Text244(Index).Text = "��"
 Else
 Text244(Index).Text = ""
 End If
End Sub

Private Sub Combo27_Click(Index As Integer)
If Combo27(Index).Text <> "-" Then
 Text20(Index).Text = "��"
 Else
 Text20(Index).Text = ""
 End If
 
End Sub

Private Sub Combo28_Click(Index As Integer)
If Combo28(Index).Text <> "-" Then
 Text44(Index).Text = "��"
 Else
 Text44(Index).Text = ""
 End If
 
End Sub

Private Sub Combo35_Click(Index As Integer)
If Combo35(Index).Text <> "-" Then
 Text28(Index).Text = "��"
 Else
 Text28(Index).Text = ""
 End If
 
End Sub



Private Sub Combo40_Click(Index As Integer)
If Combo40(Index).Text <> "-" Then
 Text33(Index).Text = "��"
 Else
 Text33(Index).Text = ""
 End If
 
End Sub
Private Sub Combo44_Click(Index As Integer)
If Combo44(Index).Text <> "-" Then
 Text37(Index).Text = "��"
 Else
 Text37(Index).Text = ""
 End If
 
End Sub

Private Sub Command1_Click()

Dim tsysno As String
tsysno = frmCareerHstRegt.ctxtsysno
Dim lobject As Object
Set lobject = dafuncGetData("select * from ְҵ�����_�Ծ�֢״�� where ϵͳ���='" & tsysno & "'")
If lobject.RecordCount > 66 Then
    frmCareerHstRegt.sub��ѯ�����
    Unload Me
Else
'��ϵͳ��11��
For i = 0 To 10
  dafuncGetData ("insert into ְҵ�����_�Ծ�֢״��(ϵͳ���,���,֢״,�̶�,����ʱ��,���ҽʦ) values ('" & tsysno & "','1','" & Label46(i).Caption & "','" & Combo1(i).Text & " ','" & Text1(i).Text & " ','" & um�û���� & "' )")
Next i
'��Ѫ�ڷ��ڹ�8��
For j = 0 To 7
  dafuncGetData ("insert into ְҵ�����_�Ծ�֢״��(ϵͳ���,���,֢״,�̶�,����ʱ��,���ҽʦ) values ('" & tsysno & "','1','" & Label70(j).Caption & "','" & Combo20(j).Text & " ','" & Text12(j).Text & " ','" & um�û���� & "' )")
Next j
'����ϵͳ��8��
For k = 0 To 7
  dafuncGetData ("insert into ְҵ�����_�Ծ�֢״��(ϵͳ���,���,֢״,�̶�,����ʱ��,���ҽʦ) values ('" & tsysno & "','1','" & Label75(k).Caption & "','" & Combo27(k).Text & " ','" & Text20(k).Text & " ','" & um�û���� & "' )")
Next k
'����ϵͳ��5��
For l = 0 To 4
  dafuncGetData ("insert into ְҵ�����_�Ծ�֢״��(ϵͳ���,���,֢״,�̶�,����ʱ��,���ҽʦ) values ('" & tsysno & "','1','" & Label83(l).Caption & "','" & Combo35(l).Text & " ','" & Text28(l).Text & " ','" & um�û���� & "' )")
Next l
'���⼰�ؽڹ�4��
For m = 0 To 3
  dafuncGetData ("insert into ְҵ�����_�Ծ�֢״��(ϵͳ���,���,֢״,�̶�,����ʱ��,���ҽʦ) values ('" & tsysno & "','1','" & Label88(m).Caption & "','" & Combo40(m).Text & " ','" & Text33(m).Text & " ','" & um�û���� & "' )")
Next m
'�۶��ǿ�ǻ�ʺ�20��
For n = 0 To 19
  dafuncGetData ("insert into ְҵ�����_�Ծ�֢״��(ϵͳ���,���,֢״,�̶�,����ʱ��,���ҽʦ) values ('" & tsysno & "','1','" & Label92(n).Caption & "','" & Combo44(n).Text & " ','" & Text37(n).Text & " ','" & um�û���� & "' )")
Next n
'��Ѫ��ϵͳ��3��
For o = 0 To 2
  dafuncGetData ("insert into ְҵ�����_�Ծ�֢״��(ϵͳ���,���,֢״,�̶�,����ʱ��,���ҽʦ) values ('" & tsysno & "','1','" & Label76(o).Caption & "','" & Combo28(o).Text & " ','" & Text44(o).Text & " ','" & um�û���� & "' )")
Next o
'����ϵͳ��7��
For p = 0 To 6
  dafuncGetData ("insert into ְҵ�����_�Ծ�֢״��(ϵͳ���,���,֢״,�̶�,����ʱ��,���ҽʦ) values ('" & tsysno & "','1','" & Label1(p).Caption & "','" & Combo128(p).Text & " ','" & Text144(p).Text & " ','" & um�û���� & "' )")
Next p

       Dim msgtip
       msgtip = MsgBox("������ɡ�", vbOKOnly + vbInformation, "��ʾ")
       frmCareerHstRegt.sub��ѯ�����
       
       Unload Me
End If
End Sub

Private Sub Command2_Click()
Unload frm֢״ѯ��

End Sub

Private Sub Command3_Click()
Dim tsysno As String
tsysno = frmCareerHstRegt.ctxtsysno
   dafuncGetData ("insert into ְҵ�����_�Ծ�֢״��(ϵͳ���,���,֢״,�̶�,����ʱ��) values ('" & tsysno & "','2','" & Text247.Text & "','" & Combo228(0).Text & " ','" & Text244(0).Text & " ')")
   Dim mtip
   mtip = MsgBox("�����ɡ�", vbOKOnly + vbInformation, "��ʾ")
   Text247.Text = ""
   Combo228(0).Text = ""
   Text244(0).Text = ""
   sub��ѯ�������֢״���
End Sub


Public Sub sub��ѯ�������֢״���()
Dim tsysno As String
tsysno = frmCareerHstRegt.ctxtsysno
Dim lobjRec As Object
        dasubSetQueryTimeout 600
        Dim lstrsql As String
        lstrsql = "select ֢״,�̶�,����ʱ�� from ְҵ�����_�Ծ�֢״�� where ϵͳ���='" & tsysno & "' and ���='2' "
        
        Set lobjRec = dafuncGetData(lstrsql)
        cgrd����֢״.Rows = 1
        
        If Not lobjRec.EOF Then
            With cgrd����֢״
                Set .DataSource = lobjRec
                If cgrd����֢״.Rows > 1 Then
                    Set mcolIndex = New Collection
                    For i = 0 To cgrd����֢״.Cols - 1
                        mcolIndex.Add i, cgrd����֢״.TextMatrix(0, i)
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
            cgrd����֢״.Rows = 1
          '  clblInfo = cgrdzzxw.Rows - 1
            Exit Sub
        End If

End Sub




Private Sub Command4_Click()
Dim selectsysno As String
Dim selectzz As String
Dim selectcd As String
Dim selectcxrq As String

selectzz = cgrd����֢״.TextMatrix(cgrd����֢״.RowSel, 0)
selectcd = cgrd����֢״.TextMatrix(cgrd����֢״.RowSel, 1)
selectcxrq = cgrd����֢״.TextMatrix(cgrd����֢״.RowSel, 2)

 dafuncGetData ("delete from ְҵ�����_�Ծ�֢״�� where �̶�='" & selectcd & "' and ����ʱ��='" & selectcxrq & "' and ϵͳ���='" & frmCareerHstRegt.ctxtsysno & "' and ֢״='" & selectzz & "'")
  
  sub��ѯ�������֢״���

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Caption = "����" Then
 sub��ѯ�������֢״���
Else
End If

End Sub

