VERSION 5.00
Object = "{8099FCC2-0A81-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frm�ڲ��շ� 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�ڲ��շ�"
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
   StartUpPosition =   3  '����ȱʡ
   Begin VSFlex6Ctl.vsFlexGrid cind�ֵ� 
      Height          =   3615
      Index           =   1
      Left            =   3240
      TabIndex        =   41
      Tag             =   "�շѱ�׼"
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
         Name            =   "����"
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
   Begin VSFlex6Ctl.vsFlexGrid cind�ֵ� 
      Height          =   3615
      Index           =   2
      Left            =   3000
      TabIndex        =   73
      Tag             =   "�շ���Ŀ"
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
         Name            =   "����"
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
   Begin VSFlex6Ctl.vsFlexGrid cind�ֵ� 
      Height          =   3615
      Index           =   0
      Left            =   2880
      TabIndex        =   40
      Tag             =   "�շ���Ŀ"
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
         Name            =   "����"
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
   Begin MSComctlLib.ImageList cimg��ťͼ�� 
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
      Caption         =   "���ü���"
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   2040
      TabIndex        =   25
      Top             =   6840
      Width           =   8970
      Begin VB.ComboBox cmb���ѷ�ʽ 
         BeginProperty Font 
            Name            =   "����"
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
      Begin VB.TextBox cinb�շ����� 
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
            Name            =   "����"
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
      Begin VB.TextBox cinb�շ����� 
         BackColor       =   &H00F0F0F0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """��""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "����"
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
      Begin VB.TextBox cinb�շ����� 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """��""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "����"
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
      Begin VB.TextBox cinb�շ����� 
         BackColor       =   &H00F0F0F0&
         BeginProperty Font 
            Name            =   "����"
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
      Begin VB.TextBox cinb�շ����� 
         BackColor       =   &H00F0F0F0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """��""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "��������"
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
         Caption         =   "���ѷ�ʽ"
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
         Caption         =   "�Ҳ����"
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
         Caption         =   "ʵ�ս��"
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
         Caption         =   "Ӧ�ս���д"
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
         Caption         =   "Ӧ�ս��"
         Height          =   180
         Index           =   20
         Left            =   150
         TabIndex        =   48
         Top             =   360
         Width           =   720
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "�������"
      Height          =   930
      Left            =   0
      TabIndex        =   24
      Top             =   6840
      Width           =   1920
      Begin VB.TextBox cinb�շ����� 
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
      Begin VB.CheckBox cchk��ӡ���۱��� 
         Caption         =   "��ӡ���۱���"
         Height          =   195
         Left            =   165
         TabIndex        =   36
         Top             =   600
         Width           =   1455
      End
      Begin MSComCtl2.UpDown cupd�޸Ĵ��۱��� 
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
         Caption         =   "���۱���"
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
            Key             =   "������ʾ"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar ctlb������ 
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
      ImageList       =   "cimg��ťͼ��"
      _Version        =   393216
      Begin VB.CheckBox Cchk��λ 
         Caption         =   "��λ��λ"
         Height          =   180
         Left            =   9240
         TabIndex        =   87
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox cchkԤ�� 
         Caption         =   "��ӡǰԤ��"
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
      TabCaption(0)   =   "�շ�"
      TabPicture(0)   =   "frm�ڲ��շ�.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Clabҵ�����"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cchkͬ���շ�"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cing�շѻ�����Ϣ��"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "ccmdѡ��"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Copt����"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Copt����"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Copt����"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Cchk��������Ϣ��ѯ"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cchk�ڲ��շ�"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Ccboҵ�����"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "�����շ�"
      TabPicture(1)   =   "frm�ڲ��շ�.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(1)=   "Frame12"
      Tab(1).ControlCount=   2
      Begin VB.ComboBox Ccboҵ����� 
         Height          =   300
         Left            =   6600
         TabIndex        =   77
         Top             =   5625
         Width           =   2295
      End
      Begin VB.CheckBox cchk�ڲ��շ� 
         Caption         =   "�ڲ��շ�(&I)"
         Height          =   300
         Left            =   165
         TabIndex        =   81
         Top             =   5625
         Width           =   1290
      End
      Begin VB.CheckBox Cchk��������Ϣ��ѯ 
         Caption         =   "��������Ϣ��ѯ"
         Height          =   300
         Left            =   9000
         TabIndex        =   80
         Top             =   5640
         Width           =   1575
      End
      Begin VB.OptionButton Copt���� 
         Caption         =   "����"
         Height          =   255
         Left            =   4920
         TabIndex        =   76
         Top             =   5625
         Width           =   735
      End
      Begin VB.OptionButton Copt���� 
         Caption         =   "����"
         Height          =   255
         Left            =   3840
         TabIndex        =   75
         Top             =   5625
         Width           =   735
      End
      Begin VB.OptionButton Copt���� 
         Caption         =   "����"
         Height          =   255
         Left            =   2760
         TabIndex        =   74
         Top             =   5625
         Width           =   735
      End
      Begin VB.CommandButton ccmdѡ�� 
         Caption         =   "ȫѡ"
         Enabled         =   0   'False
         Height          =   270
         Left            =   1680
         TabIndex        =   72
         Top             =   5625
         Width           =   780
      End
      Begin VB.Frame Frame5 
         Appearance      =   0  'Flat
         Caption         =   "�������޸ĵ� "
         ForeColor       =   &H80000008&
         Height          =   4440
         Left            =   1560
         TabIndex        =   34
         Top             =   1080
         Width           =   9300
         Begin VB.ComboBox Ccbo�շ���Ŀ���� 
            Height          =   300
            Left            =   1320
            TabIndex        =   4
            Top             =   240
            Width           =   2295
         End
         Begin VSFlex6Ctl.vsFlexGrid cing�����嵥 
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
         Begin VB.TextBox cinb�շ����� 
            Height          =   300
            Index           =   6
            Left            =   7440
            TabIndex        =   6
            Top             =   240
            Width           =   1740
         End
         Begin VB.TextBox cinb�շ����� 
            Height          =   300
            Index           =   5
            Left            =   4680
            TabIndex        =   5
            Top             =   240
            Width           =   1620
         End
         Begin VB.Label Clab�շ���Ŀ���� 
            AutoSize        =   -1  'True
            Caption         =   "�շ���Ŀ����"
            Height          =   180
            Left            =   120
            TabIndex        =   82
            Top             =   360
            Width           =   960
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�շѱ�׼"
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
            Caption         =   "�շ���Ŀ"
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
            Caption         =   "����Del��������ɾ����ǰѡ�е���Ŀ"
            Height          =   180
            Left            =   120
            TabIndex        =   35
            Top             =   3960
            Width           =   2970
         End
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         Caption         =   "������Ϣ"
         ForeColor       =   &H80000008&
         Height          =   550
         Left            =   120
         TabIndex        =   32
         Top             =   450
         Width           =   10725
         Begin VB.TextBox cinb�շ����� 
            Height          =   300
            Index           =   4
            Left            =   9120
            TabIndex        =   3
            Top             =   200
            Width           =   1485
         End
         Begin VB.TextBox cinb�շ����� 
            Height          =   300
            Index           =   3
            Left            =   4680
            TabIndex        =   2
            Top             =   200
            Width           =   1995
         End
         Begin VB.TextBox cinb�շ����� 
            Height          =   300
            Index           =   2
            Left            =   3000
            TabIndex        =   1
            Top             =   200
            Width           =   840
         End
         Begin VB.TextBox cinb�շ����� 
            Height          =   300
            Index           =   1
            Left            =   810
            TabIndex        =   0
            Top             =   200
            Width           =   1425
         End
         Begin VB.Label ClabƬ�� 
            BackStyle       =   0  'Transparent
            Caption         =   "Ƭ����(����)"
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
            Caption         =   "���ܿ���"
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
            Caption         =   "���ѵ�λ"
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
            Caption         =   "������"
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
            Caption         =   "�շѱ��"
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
         Caption         =   "������Ϣ"
         ForeColor       =   &H80000008&
         Height          =   5595
         Left            =   -74880
         TabIndex        =   29
         Top             =   360
         Width           =   3885
         Begin VB.TextBox cinb�շ����� 
            Height          =   300
            Index           =   8
            Left            =   1095
            TabIndex        =   9
            Top             =   255
            Width           =   2640
         End
         Begin VB.TextBox cinb�շ����� 
            Height          =   300
            Index           =   9
            Left            =   1095
            TabIndex        =   10
            Top             =   795
            Width           =   2640
         End
         Begin VB.TextBox cinb�շ����� 
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
         Begin VB.TextBox cinb�շ����� 
            Height          =   300
            Index           =   12
            Left            =   1095
            TabIndex        =   13
            Top             =   1890
            Width           =   2640
         End
         Begin VB.TextBox cinb�շ����� 
            Height          =   300
            Index           =   13
            Left            =   1095
            TabIndex        =   14
            Top             =   2430
            Width           =   2640
         End
         Begin VB.TextBox cinb�շ����� 
            Height          =   300
            Index           =   14
            Left            =   1095
            TabIndex        =   16
            Top             =   4185
            Width           =   720
         End
         Begin VB.TextBox cinb�շ����� 
            Height          =   300
            Index           =   15
            Left            =   2760
            TabIndex        =   17
            Top             =   4200
            Width           =   960
         End
         Begin VB.TextBox cinb�շ����� 
            Height          =   300
            Index           =   16
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   4680
            Width           =   2640
         End
         Begin VB.TextBox cinb�շ����� 
            Height          =   300
            Index           =   11
            Left            =   2070
            TabIndex        =   12
            Top             =   1350
            Width           =   420
         End
         Begin MSComCtl2.DTPicker cdtp���� 
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
         Begin MSComCtl2.DTPicker cdtp���� 
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
            Caption         =   "��Ժ����Ա"
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
            Caption         =   "����ҽ��"
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
            Caption         =   "0:��,����:Ů"
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
            Caption         =   "���ܿ���"
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
            Caption         =   "����"
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
            Caption         =   "סԺ��"
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
            Caption         =   "�Ա�"
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
            Caption         =   "����"
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
            Caption         =   "������"
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
            Caption         =   "���ѵ�λ"
            Height          =   180
            Index           =   9
            Left            =   120
            TabIndex        =   60
            Top             =   855
            Width           =   765
         End
         Begin VB.Label Label26 
            BackStyle       =   0  'Transparent
            Caption         =   "��Ժ����"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   3030
            Width           =   735
         End
         Begin VB.Label Label33 
            BackStyle       =   0  'Transparent
            Caption         =   "��Ժ����"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   3645
            Width           =   735
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         Caption         =   "�����嵥 "
         ForeColor       =   &H80000008&
         Height          =   5595
         Left            =   -70845
         TabIndex        =   27
         Top             =   360
         Width           =   6800
         Begin VB.ComboBox Ccbo���շѴ��� 
            Height          =   300
            Left            =   1320
            TabIndex        =   18
            Top             =   240
            Width           =   2295
         End
         Begin VB.TextBox cinb�շ����� 
            Height          =   300
            Index           =   18
            Left            =   1320
            TabIndex        =   21
            Top             =   720
            Width           =   2895
         End
         Begin VB.TextBox cinb�շ����� 
            Height          =   300
            Index           =   17
            Left            =   4560
            TabIndex        =   20
            Top             =   240
            Width           =   2085
         End
         Begin VSFlex6Ctl.vsFlexGrid cing�����嵥 
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
         Begin VB.Label Clab���շѴ��� 
            AutoSize        =   -1  'True
            Caption         =   "�շ���Ŀ����"
            Height          =   180
            Left            =   120
            TabIndex        =   84
            Top             =   360
            Width           =   1080
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            Caption         =   "�շѱ�׼"
            Height          =   180
            Index           =   18
            Left            =   120
            TabIndex        =   68
            Top             =   810
            Width           =   720
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            Caption         =   "�շ���Ŀ"
            Height          =   180
            Index           =   17
            Left            =   3750
            TabIndex        =   67
            Top             =   330
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "����Del��������ɾ����ǰѡ�е���Ŀ"
            Height          =   225
            Left            =   120
            TabIndex        =   28
            Top             =   3960
            Width           =   2970
         End
      End
      Begin VSFlex6Ctl.vsFlexGrid cing�շѻ�����Ϣ�� 
         Height          =   4440
         Left            =   120
         TabIndex        =   70
         Tag             =   "cing�շѻ�����Ϣ��"
         Top             =   1080
         Width           =   1365
         _cx             =   4196712
         _cy             =   4202136
         _ConvInfo       =   1
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
      Begin VB.CheckBox cchkͬ���շ� 
         Caption         =   "ͬ���շ�"
         Height          =   195
         Left            =   8880
         TabIndex        =   33
         Top             =   5280
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label Clabҵ����� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҵ�����"
         Height          =   180
         Left            =   5760
         TabIndex        =   83
         Top             =   5625
         Width           =   720
      End
   End
End
Attribute VB_Name = "frm�ڲ��շ�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Const WS_THICKFRAME = &H40000
Private Const GWL_STYLE = (-16)

Private Const ������ = "�շѽ������"
Private Const ģ���� = "frm�շѹ���"

Private Const �շ�_�շѱ�� = 1
Private Const �շ�_������ = 2
Private Const �շ�_���ѵ�λ = 3
Private Const �շ�_���ܿ��� = 4
Private Const �շ�_�շ���Ŀ = 5
Private Const �շ�_�շѱ�׼ = 6
Private Const �����շ�_�շѱ�� = 7
Private Const �����շ�_���ѵ�λ = 9
Private Const �����շ�_������ = 8
Private Const �����շ�_���� = 10
Private Const �����շ�_�Ա� = 11
Private Const �����շ�_סԺ�� = 12
Private Const �����շ�_���� = 13
Private Const �����շ�_��Ժ����Ա = 14
Private Const �����շ�_����ҽ�� = 15
Private Const �����շ�_���ܿ��� = 16
Private Const �����շ�_�շ���Ŀ = 17
Private Const �����շ�_�շѱ�׼ = 18
Private Const ���۱��� = 19
Private Const Ӧ�ս�� = 20
Private Const Ӧ�ս���д = 21
Private Const ʵ�ս�� = 22
Private Const �Ҳ���� = 23
Private Const �������� = 24

Private Const �շ� = 0
Private Const �����շ� = 1
Private Const ��Ժ = 0
Private Const ��Ժ = 1
''''''''''''''''''''''''''''''''''''''''''''
'�޸��ˣ� �켽���޸�
'���ܣ��������ѻ�����Ϣ������ֶ�����
'ʱ�䣺2001-12-20
'''''''''''''''''''''''''''''''''''''''''''
Private Const ������Ϣ_ѡ�� = 0
Private Const ������Ϣ_�շ����� = 1
Private Const ������Ϣ_�շѱ�� = 2
Private Const ������Ϣ_������ = 3
Private Const ������Ϣ_���ѵ�λ���� = 4
Private Const ������Ϣ_��� = 5

Private Const �����嵥_�շ���Ŀ��� = 0
Private Const �����嵥_�շ���Ŀ���� = 1
Private Const �����嵥_���� = 2
Private Const �����嵥_���� = 3
'Private Const �����嵥_������λ = 4
Private Const �����嵥_��� = 4

Private Const �ֵ�_�շ���Ŀ = 0
Private Const �ֵ�_�շѱ�׼ = 1
Private Const �ֵ�_���ܿ��� = 2

Private Const �ֵ�_�շ���Ŀ��� = 0
Private Const �ֵ�_�շѱ�׼��� = 0
Private Const �ֵ�_�շѱ�׼���� = 1
Private Const �ֵ�_�շ���Ŀ���� = 1
Private Const �ֵ�_���� = 1
Private Const �ֵ�_���Ƿ� = 2
Private Const �ֵ�_���� = 3
Private Const �ֵ�_��С���� = 4
Private Const �ֵ�_��󵥼� = 5
Private Const �ֵ�_������λ = 6
Private Const �ֵ�_Ʊ�����ͱ�� = 7

Dim mrds�շ���Ŀ As Recordset   '�����ȡ���շ���Ŀ�����ڳ�ʼ���ֵ�
Dim mrds�շѱ�׼ As Recordset   '�����ȡ���շѱ�׼�����ڳ�ʼ���ֵ�
Dim mrds���ܿ��� As Recordset   '�����ȡ�����ܿ��ң����ڳ�ʼ���ֵ�

Dim mrds���ѷ�ʽ As Recordset   '�����ȡ�Ľ��ѷ�ʽ,���ڳ�ʼ�����б�

Dim WithEvents mobj����ͨ�ö��� As cls����ͨ�ö���
Attribute mobj����ͨ�ö���.VB_VarHelpID = -1

Public pblnInUse As Boolean


Dim mstrUndoCount As String          '���ڱ�������ԭ�����ַ���,�Ա������벻�Ϸ�ʱ�ܹ���ԭ
Dim mstrUndoMoney As String          '���ڱ�������ԭ�����ַ���,�Ա������벻�Ϸ�ʱ�ܹ���ԭ
Dim mcur��С���� As Currency
Dim mcur��󵥼� As Currency
Dim mintCurInput As Integer     '��ǰ����������
Dim mlngX As Long               '�����"cind�ֵ�"�а��µ�Xλ��
Dim mlngY As Long               '�����"cind�ֵ�"�а��µ�Xλ��
Dim mstr���ѵ�λ��� As String  '�ӵ�λ��λ�ӿڵõ��Ľ��ѵ�λ�ı��
Dim mstr���ܿ��ұ�� As String  '�շ���Ϣ�е����ܿ��ұ��
Dim mint���ѷ�ʽ��� As Integer '���ѷ�ʽ�ı��
Dim mstr�շ����� As String
Dim mcur�ܽ�� As Currency
Dim mcur�����շ��ܽ�� As Currency
Dim mcur������Ϣ�ܽ�� As Currency


Dim mobj�շѹ��� As Object
Dim mobjҵ������ As Object
Dim mobj��λ���� As Object
Dim mint���ۿ��� As Integer
Dim mint��Ŀ���� As Integer
Dim msng���۱��� As Single
'�μ�����
Dim mblnʹ�� As Boolean         '�Ƿ���ʹ��ϵͳ
Dim mint�Ƿ��Ҽ� As Integer     '�ڽ��ѵ�λ�ı������Ƿ�ʹ�����Ҽ�
Dim mstr�շѱ�� As String      '���������¼�շѱ��
Dim mblntemp As Boolean         '�ж�������Ŀ�����Ƿ�ִ�й�

'�޸ģ�2002-10-17����������ӡǰԤ����
Private mobj����  As cls�û���������

Private Sub Ccbo���շѴ���_Click()
On Error Resume Next
   Me.MousePointer = 11
   sub�շ���Ŀ��� (Ccbo���շѴ���.Text)
   Me.MousePointer = 0
   cinb�շ�����(17).SetFocus
End Sub

Private Sub Ccbo���շѴ���_GotFocus()
On Error Resume Next
    cind�ֵ�(�ֵ�_�շѱ�׼).Visible = False
    cind�ֵ�(�ֵ�_�շ���Ŀ).Visible = False
    cinb�շ�����(17).Text = ""
End Sub

Private Sub Ccbo���շѴ���_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii <> 0 Then KeyAscii = 0
End Sub

Private Sub Ccbo�շ���Ŀ����_Click()
On Error Resume Next
   Me.MousePointer = 11
   sub�շ���Ŀ��� (Ccbo�շ���Ŀ����.Text)
   Me.MousePointer = 0
   cinb�շ�����(5).SetFocus
End Sub

Private Sub Ccbo�շ���Ŀ����_GotFocus()
On Error Resume Next
    cind�ֵ�(�ֵ�_�շѱ�׼).Visible = False
    cind�ֵ�(�ֵ�_�շ���Ŀ).Visible = False
    cinb�շ�����(�շ�_�շ���Ŀ).Text = ""
End Sub

Private Sub Ccbo�շ���Ŀ����_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii <> 0 Then KeyAscii = 0
End Sub

'����:�����շ���Ŀ����
'����:�켽��
'ʱ��:2002/07/01
Private Sub Sub�����շ���Ŀ����()
On Error GoTo errhandler
    Dim lstrSql As String           '���������¼SQL���
    Dim lobjRec As Object           '���������¼���ݼ�
    
    If Ccboҵ�����.Text = "����ҵ��" Then
    
        lstrSql = "select ҵ����� from �շѹ���_������Ϣ�� where ҵ����� is not null" & _
           " group by ҵ�����  order by ҵ����� "
    
        Set lobjRec = dafuncGetData(lstrSql)
        
        Ccboҵ�����.Clear
        
        Ccboҵ�����.AddItem "����ҵ��"
        
        Do While Not lobjRec.EOF
            If lobjRec("ҵ�����") = "" Then
                Ccboҵ�����.AddItem "����ҵ��"
            Else
                Ccboҵ�����.AddItem lobjRec("ҵ�����")
            End If
            lobjRec.MoveNext
        Loop
        
        If Ccboҵ�����.ListCount > 0 Then
            Ccboҵ�����.ListIndex = 0
            Ccboҵ�����.Refresh
        End If
    End If
Exit Sub
errhandler:
    sfsub������ "�շѹ���������", "frm�շ�", " Sub�����շ���Ŀ����", Err.Number, Err.Description
End Sub

Private Sub Ccboҵ�����_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii <> 0 Then KeyAscii = 0
End Sub

''''''''''''''''''''''''''''''''''''''
'�����ˣ��켽��
'���ܣ����ƽ����ѯ����������
'ʱ�䣺2001/12/20
'
'''''''''''''''''''''''''''''''''''
Private Sub Cchk��������Ϣ��ѯ_Click()
On Error GoTo errhandle
    If Cchk��������Ϣ��ѯ.Value = vbChecked Then
        Frame4.Enabled = True
        cinb�շ�����(1).SetFocus
        Copt����.Enabled = False
        Copt����.Enabled = False
        Copt����.Enabled = False
        
        cinb�շ�����(1).BackColor = &H80000005
        cinb�շ�����(�շ�_���ѵ�λ).BackColor = &H80000005
        cinb�շ�����(�շ�_������).BackColor = &H80000005
        cinb�շ�����(�շ�_���ܿ���).BackColor = &H80000005
        
    Else
        cinb�շ�����(�շ�_�շѱ��).Text = ""
        cinb�շ�����(�շ�_������).Text = ""
        cinb�շ�����(�շ�_���ѵ�λ).Text = ""
        cinb�շ�����(�շ�_���ܿ���).Text = ""
        Frame4.Enabled = False
        Copt����.Enabled = True
        Copt����.Enabled = True
        Copt����.Enabled = True
        Copt����.Value = True
        Copt����.Value = False
        Copt����.Value = False
        
        cinb�շ�����(1).BackColor = &H80000000
        cinb�շ�����(�շ�_���ѵ�λ).BackColor = &H80000000
        cinb�շ�����(�շ�_������).BackColor = &H80000000
        cinb�շ�����(�շ�_���ܿ���).BackColor = &H80000000
    End If
Exit Sub
errhandle:
    sfsub������ "�շѹ���������", "frm�շ�", " Cchk��������Ϣ��ѯ_Click", Err.Number, Err.Description
End Sub

Private Sub cchk�ڲ��շ�_Click()
On Error GoTo errhandle
    Dim lstrSql As String           '���������¼SQL���
    Dim lobjRec As Object           '���������¼���ݼ�
    
    Frame4.Enabled = True
    sub�������
    cing�����嵥(ctabShoufei.Tab).Rows = 1
    
    If ctabShoufei.Tab = �շ� Then
        cing�շѻ�����Ϣ��.Rows = 1
        cinb�շ�����(�շ�_�շ���Ŀ).Text = ""
        cinb�շ�����(�շ�_�շѱ�׼).Text = ""
    End If
    ''''''''''''''''''''''''''''''''''''''''
    '�޸��ˣ��켽��
    '���ܣ������ڲ��շѲ�ѯ�����ؼ��Ľ�������
    'ʱ�䣺2001-12-20
    '
    ''''''''''''''''''''''''''''''''''''''''''
    If cchk�ڲ��շ�.Value = vbChecked Then
        '�޸���:�켽��
        '����:�������Ԫ��.
        'ʱ��:2002/06/21
        
        '����Ȩ���ж��Ƿ񱨷Ϸ�����Ϣ��Ȩ��.
        '�޸���:�켽��
        'ʱ��:2002/06/22
        '��û��Ȩ���޸ĺ�ɾ���ڲ��շ���Ϣ
        
        If umfuncУ���û�Ȩ��("�շѹ���_�ڲ��շ���Ϣ�޸�") Then
            Frame5.ForeColor = &H80000012
            Frame5.Caption = "������޸ĺ�ɾ���ڲ��շ���Ϣ"
            Label2.Enabled = True
            
            '����:��������շ���Ŀ,�Խ���Ԫ�صĿ���
            'ʱ��:2002/07/19
            Clab�շ���Ŀ����.Visible = True
            Ccbo�շ���Ŀ����.Visible = True
            lblCaption(5).Visible = True
            cinb�շ�����(5).Visible = True
            
            Clab�շ���Ŀ����.Left = 120
            Ccbo�շ���Ŀ����.Left = 1250
            Ccbo�շ���Ŀ����.Width = 1500
            lblCaption(5).Left = 2900
            cinb�շ�����(5).Left = 3600
            cinb�շ�����(5).Width = 750
            cinb�շ�����(�շ�_�շ���Ŀ).Enabled = True
            Frame5.Left = 5880
            Frame5.Width = 4500
            cing�����嵥(0).Top = 600
'            cing�����嵥(0).Height = 2200
            
            Clab�շ���Ŀ����.Enabled = False
            Ccbo�շ���Ŀ����.Enabled = False
            lblCaption(5).Enabled = False
            cinb�շ�����(5).Enabled = False
            
            
        Else
            Frame5.ForeColor = &HFF&
            Frame5.Caption = "��û��Ȩ�����ӡ��޸ĺ�ɾ���ڲ��շ���Ϣ"
            Label2.Enabled = False
            cing�����嵥(0).Top = 300
'            cing�����嵥(0).Height = 2500
            Clab�շ���Ŀ����.Visible = False
            Ccbo�շ���Ŀ����.Visible = False
            lblCaption(5).Visible = False
            cinb�շ�����(5).Visible = False
            cinb�շ�����(�շ�_�շ���Ŀ).Enabled = False
        End If
        
        If umfuncУ���û�Ȩ��("�շѹ���_�ڲ��շ���Ϣ����") Then
            ctlb������.Buttons(7).Enabled = True
        End If
        cing�շѻ�����Ϣ��.Visible = True
        cing�շѻ�����Ϣ��.Width = 5685
        cing�����嵥(0).Width = 4845
        Frame5.Left = 5880
        Frame5.Width = 5100
        lblCaption(6).Visible = False
        cinb�շ�����(6).Visible = False
        
        ccmdѡ��.Enabled = True
        Copt����.Enabled = True
        Copt����.Value = True
        Copt����.Enabled = True
        Copt����.Enabled = True
        Cchk��������Ϣ��ѯ.Enabled = True
        
        'cing�����嵥(�շ�).Enabled = False
        cinb�շ�����(�շ�_�շѱ��).Enabled = True
        cinb�շ�����(�շ�_�շѱ��).SetFocus
        cchkͬ���շ�.Enabled = True
        
        If cchkͬ���շ�.Value = vbChecked Then
            cing�շѻ�����Ϣ��.Enabled = True
        Else
            cing�շѻ�����Ϣ��.Enabled = False
        End If
        ctlb������.Buttons(1).Enabled = True
        cinb�շ�����(�շ�_�շѱ�׼).Enabled = False
        cing�շѻ�����Ϣ��.Enabled = True
        'cing�����嵥(�շ�).Editable = False
        
        '����:��ȡϵͳ�ж��շѵ�ҵ�����
        'ʱ��:2002/06/25
        '����:�켽��
        Clabҵ�����.Enabled = True
        Ccboҵ�����.Enabled = True
        Ccboҵ�����.BackColor = &H80000005
        lstrSql = "select ҵ����� from �շѹ���_������Ϣ�� where ҵ����� is not null" & _
           " group by ҵ�����  order by ҵ����� "
    
        Set lobjRec = dafuncGetData(lstrSql)
        
        Ccboҵ�����.Clear
        
        Ccboҵ�����.AddItem "����ҵ��"
        
        'lobjRec.MoveFirst
        Do While Not lobjRec.EOF
            If lobjRec("ҵ�����") = "" Then
                Ccboҵ�����.AddItem "����ҵ��"
            Else
                Ccboҵ�����.AddItem lobjRec("ҵ�����")
            End If
            lobjRec.MoveNext
        Loop
        
        Ccboҵ�����.ListIndex = 0
        Ccboҵ�����.Refresh
        
        Frame4.Enabled = False
        
        '�޸ģ�2001-11-22���������ѵ�λ�����ܿ��ң���ѯ��
        cinb�շ�����(1).BackColor = &H80000000
        cinb�շ�����(�շ�_���ѵ�λ).BackColor = &H80000000
        cinb�շ�����(�շ�_������).BackColor = &H80000000
        cinb�շ�����(�շ�_���ܿ���).BackColor = &H80000000
    Else

        '�޸���:�켽��
        '����:�������Ԫ��.
        'ʱ��:2002/06/21
        
        
        Clab�շ���Ŀ����.Left = 120
        Ccbo�շ���Ŀ����.Left = 1320
        Ccbo�շ���Ŀ����.Width = 2295
        lblCaption(5).Left = 3960
        cinb�շ�����(5).Left = 4800
        cinb�շ�����(5).Width = 1380
        Frame5.ForeColor = &H80000012
        Frame5.Caption = "�����嵥�޸�"
        Frame5.Left = 120
        Frame5.Width = 10845
        
        ctlb������.Buttons(7).Enabled = False
        cing�շѻ�����Ϣ��.Visible = False

        cing�����嵥(0).Width = 10600
        Clab�շ���Ŀ����.Visible = True
        Ccbo�շ���Ŀ����.Visible = True
        lblCaption(5).Visible = True
        lblCaption(6).Visible = True
        cinb�շ�����(5).Visible = True
        cinb�շ�����(6).Visible = True
        cing�����嵥(0).Top = 600
'        cing�����嵥(0).Height = 2100
         
        ClabƬ��.Caption = "Ƭ����(����)"
    
        '�켽�����޸ģ���ȡ���ڲ��շѺ�Cchk��������Ϣ��ѯ����Ϊû��ѡ��״̬
        Cchk��������Ϣ��ѯ.Value = Unchecked
        
        Frame4.Enabled = True
        ccmdѡ��.Enabled = False
        Copt����.Enabled = False
        Copt����.Enabled = False
        Copt����.Enabled = False
        Cchk��������Ϣ��ѯ.Enabled = False
         
        cing�����嵥(�շ�).Enabled = True
        cinb�շ�����(�շ�_�շѱ��).Enabled = False
        cchkͬ���շ�.Enabled = False
        cing�շѻ�����Ϣ��.Enabled = False
        cing�����嵥(�շ�).Editable = True
        ctlb������.Buttons(1).Enabled = False
        cinb�շ�����(�շ�_�շ���Ŀ).Enabled = True
        cinb�շ�����(�շ�_�շѱ�׼).Enabled = True
        cinb�շ�����(�շ�_������).Enabled = True
        If cinb�շ�����(�շ�_������).Enabled Then
            cinb�շ�����(�շ�_������).SetFocus
        ElseIf cinb�շ�����(�շ�_�շѱ��).Enabled Then
            cinb�շ�����(�շ�_�շѱ��).SetFocus
        End If
        cinb�շ�����(��������).Text = Date
        cinb�շ�����(�շ�_���ܿ���).Text = um�û���������
        cinb�շ�����(�շ�_���ѵ�λ).Enabled = True
        cinb�շ�����(�շ�_������).Enabled = True
        cinb�շ�����(�շ�_���ܿ���).Enabled = True
        
        'ҵ�����������
        'ʱ��:2002/05/06
        '����: �켽��
    
        Clabҵ�����.Enabled = False
        Ccboҵ�����.Enabled = False
        Ccboҵ�����.BackColor = &H80000000
        
        cinb�շ�����(1).BackColor = &H80000005
        cinb�շ�����(�շ�_���ѵ�λ).BackColor = &H80000005
        cinb�շ�����(�շ�_������).BackColor = &H80000005
        cinb�շ�����(�շ�_���ܿ���).BackColor = &H80000005
            
        Clab�շ���Ŀ����.Enabled = True
        Ccbo�շ���Ŀ����.Enabled = True
        lblCaption(5).Enabled = True
            
    End If
Exit Sub
errhandle:
    sfsub������ "�շѹ���������", "frm�շ�", " cchk�ڲ��շ�_Click", Err.Number, Err.Description
End Sub
Private Sub cchk�ڲ��շ�_GotFocus()
On Error GoTo errhandle
    cind�ֵ�(�ֵ�_�շѱ�׼).Visible = False
    cind�ֵ�(�ֵ�_�շ���Ŀ).Visible = False
    cind�ֵ�(�ֵ�_���ܿ���).Visible = False
Exit Sub
errhandle:
    sfsub������ "�շѹ���������", "frm�շ�", " cchk�ڲ��շ�_GotFocus", Err.Number, Err.Description
End Sub

Private Sub ccmdѡ��_Click()
On Error GoTo errhandle
    Dim i As Long
    Dim lcur�ܽ�� As Currency
    If ccmdѡ��.Caption = "ȫѡ" Then
        
        '*******����Ϊ�μ�����-0823********
        mcur������Ϣ�ܽ�� = 0
        '*******����Ϊ�μ�����-0823********
        
        For i = 1 To cing�շѻ�����Ϣ��.Rows - 1
            cing�շѻ�����Ϣ��.Cell(flexcpChecked, i, 0) = 1
            mcur������Ϣ�ܽ�� = (mcur������Ϣ�ܽ�� + cing�շѻ�����Ϣ��.TextMatrix(i, ������Ϣ_���))
        Next
        
        '*******����Ϊ�μ�����-0823********
        lcur�ܽ�� = mcur������Ϣ�ܽ�� * CDbl(cinb�շ�����(19).Text)
        '*******����Ϊ�μ�����-0823********
        
        ccmdѡ��.Caption = "���"
    Else
        For i = 1 To cing�շѻ�����Ϣ��.Rows - 1
            cing�շѻ�����Ϣ��.Cell(flexcpChecked, i, 0) = 2
        Next
        mcur������Ϣ�ܽ�� = 0
        
        '********����Ϊ�μ�����-0823*********
        lcur�ܽ�� = 0
        '********����Ϊ�μ�����-0823*********
        
        ccmdѡ��.Caption = "ȫѡ"
    End If
    
    '********����Ϊ�μ�����-0823*********
    'cinb�շ�����(Ӧ�ս��).Text = mcur������Ϣ�ܽ��
    cinb�շ�����(Ӧ�ս��).Text = lcur�ܽ��
    '********����Ϊ�μ�����-0823*********
Exit Sub
errhandle:
    sfsub������ "�շѹ���������", "frm�շ�", " ccmdѡ��_Click", Err.Number, Err.Description
End Sub

Private Sub ccmdѡ��_GotFocus()
On Error GoTo errhandle
    cind�ֵ�(�ֵ�_�շѱ�׼).Visible = False
    cind�ֵ�(�ֵ�_�շ���Ŀ).Visible = False
    cind�ֵ�(�ֵ�_���ܿ���).Visible = False
Exit Sub
errhandle:
    sfsub������ "�շѹ���������", "frm�շ�", "ccmdѡ��_GotFocus", Err.Number, Err.Description
End Sub

Private Sub cdtp����_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo errhandle
    If KeyCode = vbKeyReturn Then
        Select Case Index
            Case ��Ժ
                If cdtp����(��Ժ).Enabled Then cdtp����(��Ժ).SetFocus
            Case ��Ժ
                If cinb�շ�����(�����շ�_��Ժ����Ա).Enabled Then cinb�շ�����(�����շ�_��Ժ����Ա).SetFocus
        End Select
    End If
Exit Sub
errhandle:
    sfsub������ "�շѹ���������", "frm�շ�", "cdtp����_KeyDown", Err.Number, Err.Description
End Sub

Private Sub cinb�շ�����_Change(Index As Integer)
    On Error GoTo errhandle
    Static lcurMoney As Currency
    Static lintAge As Integer
    Static lsngR As Single

    Select Case Index
        Case Ӧ�ս��
            cinb�շ�����(Ӧ�ս���д).Text = FuncConvertToCapsStr(Val(cinb�շ�����(Ӧ�ս��).Text))
            cinb�շ�����(�Ҳ����).Text = Format(Val(cinb�շ�����(ʵ�ս��).Text) - Val(cinb�շ�����(Ӧ�ս��).Text), "0.00")
            
        Case ���۱���
            If cinb�շ�����(���۱���).Text = vbNullString Then cinb�շ�����(���۱���).Text = "1.00"
            If Val(cinb�շ�����(���۱���).Text) > 1 Then cinb�շ�����(���۱���).Text = "1.00"
            If Val(cinb�շ�����(���۱���).Text) < 0 Then cinb�շ�����(���۱���).Text = "0.00"
            If Not IsNumeric(cinb�շ�����(���۱���).Text) Then cinb�շ�����(���۱���).Text = "1.00"
            
            
            If ctabShoufei.Tab = �շ� And cchk�ڲ��շ�.Value = 1 Then
                cinb�շ�����(Ӧ�ս��).Text = mcur������Ϣ�ܽ�� * Val(cinb�շ�����(���۱���).Text)
            Else
                cinb�շ�����(Ӧ�ս��).Text = mcur�ܽ�� * Val(cinb�շ�����(���۱���).Text)
            End If
            cinb�շ�����(Ӧ�ս���д).Text = FuncConvertToCapsStr(Val(cinb�շ�����(Ӧ�ս��).Text))
            cinb�շ�����(�Ҳ����).Text = Format(Val(cinb�շ�����(ʵ�ս��).Text) - Val(cinb�շ�����(Ӧ�ս��).Text), "0.00")
            
        Case ʵ�ս��
            If cinb�շ�����(ʵ�ս��).Text = vbNullString Then cinb�շ�����(ʵ�ս��).Text = 0
            If Not IsNumeric(cinb�շ�����(ʵ�ս��).Text) Then
                cinb�շ�����(ʵ�ս��).Text = CStr(lcurMoney)
            Else
                lcurMoney = Val(cinb�շ�����(ʵ�ս��).Text)
            End If
            cinb�շ�����(�Ҳ����).Text = Format(Val(cinb�շ�����(ʵ�ս��).Text) - Val(cinb�շ�����(Ӧ�ս��).Text), "0.00")
            
        Case �Ҳ����
            If Val(cinb�շ�����(�Ҳ����).Text) < 0 Then
                cinb�շ�����(�Ҳ����).ForeColor = &HFF
            Else
                cinb�շ�����(�Ҳ����).ForeColor = &HFF0000
            End If
            
        Case �շ�_���ܿ���, �����շ�_���ܿ���
        Case �����շ�_����
            If cinb�շ�����(�����շ�_����).Text = vbNullString Then cinb�շ�����(�����շ�_����).Text = "0"
            If IsNumeric(cinb�շ�����(�����շ�_����).Text) Then lintAge = Fix(cinb�շ�����(�����շ�_����).Text)
            If lintAge < 0 Then lintAge = 0
            cinb�շ�����(�����շ�_����).Text = CStr(lintAge)
            
        Case �����շ�_�Ա�
            If cinb�շ�����(Index).Text = "" Then Exit Sub
            If cinb�շ�����(Index).Text = "Ů" Then Exit Sub
            If Asc(cinb�շ�����(Index).Text) = Asc("0") Or cinb�շ�����(Index).Text = "��" Then
                cinb�շ�����(Index).Text = "��"
            Else
                cinb�շ�����(Index).Text = "Ů"
            End If
        Case �շ�_�շ���Ŀ, �����շ�_�շ���Ŀ
            Call funcƥ���շ���Ŀ(cinb�շ�����(Index))
        Case �շ�_�շѱ�׼, �����շ�_�շѱ�׼
            Call funcƥ���շѱ�׼(cinb�շ�����(Index))
    End Select
    
errhandle:
    If Err.Number = 0 Then Exit Sub
    sfsub������ "�շѽ���", "frm�շ�", "cinb�շ�����_Change", Err.Number, Err.Description
End Sub

Private Sub cinb�շ�����_GotFocus(Index As Integer)
On Error GoTo errhandle
    Dim i As Long
    Dim lrdsTemp As Recordset
    Dim j As Long
        
    cinb�շ�����(Index).SelStart = 0
    cinb�շ�����(Index).SelLength = Len(cinb�շ�����(Index).Text)
    Set lrdsTemp = Nothing
    '���浱ǰ����������
    mintCurInput = Index
    Select Case Index
'&  ===========================| �շ���Ŀ��ý��� |==============================
        Case �շ�_�շ���Ŀ, �����շ�_�շ���Ŀ
            '���ô���Ԥ�ȴ�������¼�
            Me.KeyPreview = False
            '����Ҫ��ʾ���ֵ�
            If Not cind�ֵ�(�ֵ�_�շ���Ŀ).Visible Then
                cind�ֵ�(�ֵ�_�շѱ�׼).Visible = False
                cind�ֵ�(�ֵ�_�շ���Ŀ).Visible = True
                cind�ֵ�(�ֵ�_���ܿ���).Visible = False
                'cinb�շ�����(Index).SetFocus
            Else
                cinb�շ�����(Index).SetFocus
            End If
'&  ===========================| �շѱ�׼��ý��� |==============================
        Case �շ�_�շѱ�׼, �����շ�_�շѱ�׼
            '���ô���Ԥ�ȴ�������¼�
            Me.KeyPreview = False
            '����Ҫ��ʾ���ֵ�
            If Not cind�ֵ�(�ֵ�_�շѱ�׼).Visible Then
                cind�ֵ�(�ֵ�_�շѱ�׼).Visible = True
                cind�ֵ�(�ֵ�_�շ���Ŀ).Visible = False
                cind�ֵ�(�ֵ�_���ܿ���).Visible = False
                cinb�շ�����(Index).SetFocus
            Else
                cinb�շ�����(Index).SetFocus
            End If
'&  ===========================| ���ѵ�λ��ý��� |=============================
        Case �շ�_���ѵ�λ, �����շ�_���ѵ�λ
            '***************���ºμ��޸�09-12******************
            If mint�Ƿ��Ҽ� = 1 Then Exit Sub
            '***************���Ϻμ��޸�09-12******************
            If (Index = �շ�_���ѵ�λ And ctabShoufei.Tab = �շ�) Or (Index = �����շ�_���ѵ�λ And ctabShoufei.Tab = �����շ�) Then
                Dim lrds������Ϣ As Recordset               '��λ�Ĵ�����Ϣ
                'Dim lobj������Ϣ As Object
                '�ر��ֵ�
                cind�ֵ�(�ֵ�_�շѱ�׼).Visible = False
                cind�ֵ�(�ֵ�_�շ���Ŀ).Visible = False
                cind�ֵ�(�ֵ�_���ܿ���).Visible = False
                
                '���õ�λ�����Ķ�λ�ӿڻ�ȡ��λ��Ϣ
                '���ܣ���λ��λ�Ĺ��ܣ�����Checkbox����ֵ���ж��Ƿ����� �켽�� 2002/09/30
                If Cchk��λ.Value = 1 Then
                    Set lrdsTemp = mobj��λ����.func��λ�򵥶�λ(100, 100)
                    If Not (lrdsTemp Is Nothing) Then
                        If lrdsTemp.RecordCount > 0 Then
                            '��ʾ��λ����`
                            cinb�շ�����(Index).Text = lrdsTemp("��λ����")
                            If lrdsTemp("Ƭ��") = "" Then
                                ClabƬ��.Caption = "Ƭ����(����)"
                            Else
                                ClabƬ��.Caption = "(" + lrdsTemp("Ƭ��") + ")"
                            End If
                            '���浥λ��������
                            mstr���ѵ�λ��� = lrdsTemp("������")
                            '���ý���
                            If cinb�շ�����(Index).Enabled Then
                                cinb�շ�����(Index).SetFocus
                            End If
                        End If
                    End If
                End If
                
                If Not (mobjҵ������ Is Nothing) Then
                    '��ѯ������Ϣ
                    Set lrds������Ϣ = mobjҵ������.func��ѯ������Ϣ("��λ���='" & mstr���ѵ�λ��� & "'")
                End If
                If Not (lrds������Ϣ Is Nothing) Then
                    If Not (lrds������Ϣ.BOF And lrds������Ϣ.EOF) Then
                        '��ʾ������Ϣ
                        'cinb�շ�����(���۱���).Text = lrds������Ϣ("���۱���")
                        'msng���۱��� = lrds������Ϣ

                        If mint���ۿ��� = 0 Then
                            cinb�շ�����(���۱���).Text = "1.00"
                        Else
                            cinb�շ�����(���۱���).Text = lrds������Ϣ("���۱���")
                        End If
                    Else
                        cinb�շ�����(���۱���).Text = "1.00"
                    End If
                Else
                    cinb�շ�����(���۱���).Text = "1.00"
                End If
                Set lrds������Ϣ = Nothing
                Set lrdsTemp = Nothing
            End If
        Case �շ�_���ܿ���, �����շ�_���ܿ���
            Me.KeyPreview = False
            If Not cind�ֵ�(�ֵ�_�շѱ�׼).Visible Then
                cind�ֵ�(�ֵ�_�շѱ�׼).Visible = False
                cind�ֵ�(�ֵ�_�շ���Ŀ).Visible = False
                cind�ֵ�(�ֵ�_���ܿ���).Visible = True
                cinb�շ�����(Index).SetFocus
            Else
                cinb�շ�����(Index).SetFocus
            End If
'&  ===========================| ������ý��� |==============================
        Case Else
            cind�ֵ�(�ֵ�_�շѱ�׼).Visible = False
            cind�ֵ�(�ֵ�_�շ���Ŀ).Visible = False
            cind�ֵ�(�ֵ�_���ܿ���).Visible = False
    End Select
Exit Sub
errhandle:
    sfsub������ "�շѽ���", "frm�շ�", "cinb�շ�����_GotFocus", Err.Number, Err.Description
End Sub

'�ڴ˴����� "UP","DOWN"
Private Sub cinb�շ�����_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo errhandle

     
     
    '���ܣ���mstr���ѵ�λ����ѱ����е�λ���ʱ������û��ֲ����ֹ����뷽ʽ����Ҫ��ձ����б���ı��
    'ʱ�䣺2002/09/30 �켽��
    If (Index = �շ�_���ѵ�λ Or Index = �����շ�_���ѵ�λ) And KeyCode <> 13 Then
       mstr���ѵ�λ��� = ""
    End If

    '�жϰ���
    Select Case KeyCode
        '���� "UP"
        Case vbKeyUp
            Select Case Index
                'Ӱ���ֵ� cind�ֵ�(�ֵ�_�շ���Ŀ)
                Case �շ�_�շ���Ŀ, �����շ�_�շ���Ŀ
                    With cind�ֵ�(�ֵ�_�շ���Ŀ)
                        If .RowSel > 1 Then
                            .RowSel = .RowSel - 1
                            .Select .RowSel, 0, .RowSel
                            .TopRow = .RowSel
                        End If
                    End With
                'Ӱ���ֵ� cind�ֵ�(�ֵ�_�շѱ�׼)
                Case �շ�_�շѱ�׼, �����շ�_�շѱ�׼
                    With cind�ֵ�(�ֵ�_�շѱ�׼)
                        If .RowSel > 1 Then
                            .RowSel = .RowSel - 1
                            .Select .RowSel, 0, .RowSel
                            .TopRow = .RowSel
                        End If
                    End With
                Case �շ�_���ܿ���, �����շ�_���ܿ���
                    With cind�ֵ�(�ֵ�_���ܿ���)
                        If .RowSel > 1 Then
                            .RowSel = .RowSel - 1
                            .Select .RowSel, 0, .RowSel
                            .TopRow = .RowSel
                        End If
                    End With
            End Select
        '���� "DOWN"
        Case vbKeyDown
            Select Case Index
                Case �շ�_�շ���Ŀ, �����շ�_�շ���Ŀ
                    With cind�ֵ�(�ֵ�_�շ���Ŀ)
                        If .RowSel < .Rows - 1 Then
                            .RowSel = .RowSel + 1
                            .Select .RowSel, 0, .RowSel
                            .TopRow = .RowSel
                        End If
                    End With
                Case �շ�_�շѱ�׼, �����շ�_�շѱ�׼
                    With cind�ֵ�(�ֵ�_�շѱ�׼)
                        If .RowSel < .Rows - 1 Then
                            .RowSel = .RowSel + 1
                            .Select .RowSel, 0, .RowSel
                            .TopRow = .RowSel
                        End If
                    End With
                Case �շ�_���ܿ���, �����շ�_���ܿ���
                    With cind�ֵ�(�ֵ�_���ܿ���)
                        If .RowSel < .Rows - 1 Then
                            .RowSel = .RowSel + 1
                            .Select .RowSel, 0, .RowSel
                            .TopRow = .RowSel
                        End If
                    End With
            End Select
        Case vbKeyEscape
            cind�ֵ�(�ֵ�_�շѱ�׼).Visible = False
            cind�ֵ�(�ֵ�_�շ���Ŀ).Visible = False
            cind�ֵ�(�ֵ�_���ܿ���).Visible = False
            Me.KeyPreview = True
        Case Else
    End Select
Exit Sub
errhandle:
    sfsub������ "�շѽ���", "frm�շ�", "cinb�շ�����_KeyDown", Err.Number, Err.Description
End Sub

Private Sub cinb�շ�����_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo errhandle
    Dim lrds�շѱ�׼ As Recordset
    Dim i As Long
    Dim j As Long
    Dim lcurMoney As Currency
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        '����س���
        Case vbKeyReturn
            KeyAscii = 0
            '*******����Ϊ�μ����� 09-12 **************
            mint�Ƿ��Ҽ� = 0
            '*******����Ϊ�μ����� 09-12 **************
            Select Case Index
'&  ===========================| �շ�_�շ���Ŀ, �����շ�_�շ���Ŀ |==============================
                Case �շ�_�շѱ��
                    mobj����ͨ�ö���_BeforeOperate "��ѯ", False
                    
                Case �շ�_�շ���Ŀ, �����շ�_�շ���Ŀ
                    
                    cinb�շ�����(�շ�_�շ���Ŀ).Text = cind�ֵ�(�ֵ�_�շ���Ŀ).TextMatrix(cind�ֵ�(�ֵ�_�շ���Ŀ).RowSel, �ֵ�_�շ���Ŀ����)
                    If Not func�����Ŀ�Ƿ���ѡ(cind�ֵ�(�ֵ�_�շ���Ŀ).TextMatrix(cind�ֵ�(�ֵ�_�շ���Ŀ).RowSel, �ֵ�_�շ���Ŀ���)) Then
                        
                        'ֻ�����ڲ��շѣ�������Ȩ�޵������ִ�У��켽����2002/07/22
                        If cchk�ڲ��շ�.Value = vbChecked And ctabShoufei.Caption = "�շ�" Then
                            '�����ݿ��������շ���Ŀ
                            Dim lblntemp As Boolean         '��¼��������ֵ
                            lblntemp = False
                            lblntemp = func�����շ���Ŀ(mstr�շѱ��, cind�ֵ�(�ֵ�_�շ���Ŀ).TextMatrix(cind�ֵ�(�ֵ�_�շ���Ŀ).RowSel, �ֵ�_�շ���Ŀ���), cind�ֵ�(�ֵ�_�շ���Ŀ).TextMatrix(cind�ֵ�(�ֵ�_�շ���Ŀ).RowSel, �ֵ�_����))
                            
                            
                            '���ܣ����޸Ĵ�����Ϣ��ʽ���� ʱ�䣺2002/08/05 ���ߣ��켽��
                            '�޸�:��������ϸ���շ���Ŀ��Ϣ
                            
                            Dim lstrItemName As String      '���������¼�շ���Ŀ����
                            Dim lstrQuerySql As String      '���������¼��¼SQL���
                            Dim lstrPrice As String         '���������¼����
                            Dim lstrIntro As String         '������ϸ˵��
                            
                            lstrIntro = ""
                            lstrItemName = cind�ֵ�(�ֵ�_�շ���Ŀ).TextMatrix(cind�ֵ�(�ֵ�_�շ���Ŀ).RowSel, �ֵ�_�շ���Ŀ����)
                            lstrPrice = cind�ֵ�(�ֵ�_�շ���Ŀ).TextMatrix(cind�ֵ�(�ֵ�_�շ���Ŀ).RowSel, �ֵ�_����)
                            lstrIntro = "�շ���Ŀ���ƣ�" & lstrItemName & "�����ۣ� " & lstrPrice & "Ԫ��"
                            
                            sub��Ϣ���� mstr�շѱ��, "�շѱ��Ϊ��" & mstr�շѱ�� & "�ķ�����Ϣ���б��������շ���Ŀ��" & lstrIntro
                        End If
                        
                        '�ж������ˢ��,�켽����2002/07/22
                        If (mblntemp = False) Or (mblntemp = True And lblntemp = True) Then
                            cing�����嵥(ctabShoufei.Tab).AddItem cind�ֵ�(�ֵ�_�շ���Ŀ).TextMatrix(cind�ֵ�(�ֵ�_�շ���Ŀ).RowSel, �ֵ�_�շ���Ŀ���) & vbTab & _
                                                       cind�ֵ�(�ֵ�_�շ���Ŀ).TextMatrix(cind�ֵ�(�ֵ�_�շ���Ŀ).RowSel, �ֵ�_�շ���Ŀ����) & vbTab & _
                                                       cind�ֵ�(�ֵ�_�շ���Ŀ).TextMatrix(cind�ֵ�(�ֵ�_�շ���Ŀ).RowSel, �ֵ�_����) & vbTab & _
                                                       "1" & vbTab & _
                                                       cind�ֵ�(�ֵ�_�շ���Ŀ).TextMatrix(cind�ֵ�(�ֵ�_�շ���Ŀ).RowSel, �ֵ�_����)
                            For i = 1 To cing�����嵥(ctabShoufei.Tab).Rows - 1
                                lcurMoney = lcurMoney + Val(cing�����嵥(ctabShoufei.Tab).TextMatrix(i, �����嵥_���))
                            Next
                            mcur�ܽ�� = lcurMoney
                            If cchk�ڲ��շ�.Value = 0 Then
                                cinb�շ�����(Ӧ�ս��) = lcurMoney * Val(cinb�շ�����(���۱���).Text)
                                cinb�շ�����(Ӧ�ս���д) = FuncConvertToCapsStr(Val(cinb�շ�����(Ӧ�ս��)))
                            End If
                            cinb�շ�����(Index).SelStart = 0
                            cinb�շ�����(Index).SelLength = Len(cinb�շ�����(Index).Text)
                        End If
                        
                        '�ڳɹ������ݿ��������շ���Ŀ��Ҫ���³�ʼ���������켽����2002/07/22
                        lblntemp = False
                        mblntemp = False
                        
                        
                        'ˢ�½���
                        For i = 1 To cing�շѻ�����Ϣ��.Rows - 1
                            If cing�շѻ�����Ϣ��.Cell(flexcpText, i, 2) = mstr�շѱ�� Then
                            cing�շѻ�����Ϣ��.TextMatrix(i, 5) = mcur�ܽ�� * Val(cinb�շ�����(���۱���).Text)
                            End If
                        Next
                        
                        
                        '�ж��Ƿ���ѡ�еķ�����Ϣ
                        Dim lbln�Ƿ���ѡ���� As Boolean
                        lbln�Ƿ���ѡ���� = False
    
                        For i = 1 To cing�շѻ�����Ϣ��.Rows - 1
                            If cing�շѻ�����Ϣ��.Cell(flexcpChecked, i, 0) = 1 Then
                                lbln�Ƿ���ѡ���� = True
                                Exit For
                            End If
                        Next
    
                        If lbln�Ƿ���ѡ���� = True Then
                            sub��������ˢ��
                        End If
                        
                    Else
                        sffuncMsg "���շ���Ŀ��ѡ��" & vbCrLf & "�����޸�����,����������ֱ���޸�.", sf����
                        Exit Sub
                    End If
                
                Case �շ�_�շѱ�׼, �����շ�_�շѱ�׼
                    If mobjҵ������ Is Nothing Then
                        sffuncMsg "ҵ����� ""mobjҵ������"" ��δ������", sf����
                        Exit Sub
                    End If
                    
                    Set lrds�շѱ�׼ = mobj�շѹ���.funcExecute("select a.�շ���Ŀ���,b.�շ���Ŀ����,a.����,a.����,b.������λ,���=a.����*a.���� from �շѹ���_�շѱ�׼��Ϣ�� a,�շѹ���_�շ���Ŀ�ֵ�� b where b.�շ���Ŀ���=a.�շ���Ŀ��� and �շѱ�׼����='" & cind�ֵ�(�ֵ�_�շѱ�׼).TextMatrix(cind�ֵ�(�ֵ�_�շѱ�׼).RowSel, �ֵ�_�շѱ�׼����) & "'", "cls������Ϣ")
                    If lrds�շѱ�׼ Is Nothing Then
                        sffuncMsg "δ�ҵ�ָ�����շѱ�׼��", sf����
                        Exit Sub
                    End If
                    If lrds�շѱ�׼.BOF And lrds�շѱ�׼.EOF Then
                        sffuncMsg "�շѱ�׼�����շ���Ŀ��", sf����
                        Exit Sub
                    Else
                        lrds�շѱ�׼.MoveFirst
                        Dim llngItemCount As Long
                        For i = 0 To lrds�շѱ�׼.RecordCount - 1
                            If Not func�����Ŀ�Ƿ���ѡ(lrds�շѱ�׼("�շ���Ŀ���")) Then
                            
                            cing�����嵥(ctabShoufei.Tab).AddItem lrds�շѱ�׼("�շ���Ŀ���") & vbTab & _
                                                                  lrds�շѱ�׼("�շ���Ŀ����") & vbTab & _
                                                                  lrds�շѱ�׼("����") & vbTab & _
                                                                  lrds�շѱ�׼("����") & vbTab & _
                                                                  lrds�շѱ�׼("���")
                            llngItemCount = llngItemCount + 1
                            Else
                            End If
                            If Not lrds�շѱ�׼.EOF Then lrds�շѱ�׼.MoveNext
                        Next
                        For i = 1 To cing�����嵥(ctabShoufei.Tab).Rows - 1
                            lcurMoney = lcurMoney + Val(cing�����嵥(ctabShoufei.Tab).TextMatrix(i, �����嵥_���))
                        Next
                        mcur�ܽ�� = lcurMoney
                        cinb�շ�����(Ӧ�ս��) = lcurMoney * Val(cinb�շ�����(���۱���).Text)
                        cinb�շ�����(Ӧ�ս���д) = FuncConvertToCapsStr(Val(cinb�շ�����(Ӧ�ս��)))
                        cinb�շ�����(Index).SelStart = 0
                        cinb�շ�����(Index).SelLength = Len(cinb�շ�����(Index).Text)
                        If llngItemCount = lrds�շѱ�׼.RecordCount Then
                            MsgBox "�շѱ�׼�е������շ���Ŀ(" & llngItemCount & "��)����ӵ������嵥�У�" & vbCrLf & vbCrLf & "(���ι�������� " & lrds�շѱ�׼.RecordCount & " ���е� " & llngItemCount & " ���շ���Ŀ��)", vbInformation, "ϵͳ��ʾ"
                        ElseIf llngItemCount = 0 Then
                            MsgBox "�շѱ�׼�е������շ���Ŀ�ڷ����嵥������ӣ�" & vbCrLf & vbCrLf & "(���ι�������� " & lrds�շѱ�׼.RecordCount & " ���е� " & llngItemCount & " ���շ���Ŀ��)", vbInformation, "ϵͳ��ʾ"
                        Else
                            MsgBox "�շѱ�׼�в����շ���Ŀ�ڷ����嵥�������,����� " & llngItemCount & " ������ӵ������嵥��" & vbCrLf & vbCrLf & "(���ι�������� " & lrds�շѱ�׼.RecordCount & " ���е� " & llngItemCount & " ���շ���Ŀ��)", vbInformation, "ϵͳ��ʾ"
                        End If
                    End If
                
                Case �շ�_���ܿ���, �����շ�_���ܿ���
                    If cinb�շ�����(Index + 1).Enabled Then
                        cinb�շ�����(Index + 1).SetFocus
                    Else
                        cinb�շ�����(�շ�_�շѱ��).SetFocus
                    End If
                    If cind�ֵ�(�ֵ�_���ܿ���).Visible Then cinb�շ�����(Index).Text = cind�ֵ�(�ֵ�_���ܿ���).TextMatrix(cind�ֵ�(�ֵ�_���ܿ���).RowSel, 1)
                    mstr���ܿ��ұ�� = cind�ֵ�(�ֵ�_���ܿ���).TextMatrix(cind�ֵ�(�ֵ�_���ܿ���).RowSel, 0)
                Case ʵ�ս��
                    Call mobj����ͨ�ö���_BeforeOperate("�շ�", False)
                Case Else
                    '���շѽ������ƶ�����
                    If Index < �շ�_���ܿ��� Then
                        If cinb�շ�����(Index + 1).Enabled Then cinb�շ�����(Index + 1).SetFocus
                    End If
                    If Index > �շ�_�շѱ�׼ And Index < �����շ�_�շ���Ŀ Then
                        If Index = �����շ�_���� Then
                            If cdtp����(0).Enabled Then cdtp����(0).SetFocus
                        Else
                            If cinb�շ�����(Index + 1).Enabled Then cinb�շ�����(Index + 1).SetFocus
                        End If
                    End If
            End Select
        End Select
Exit Sub
errhandle:
    sfsub������ "�շѽ���", "frm�շ�", "cinb�շ�����_KeyPress", Err.Number, Err.Description
End Sub

'***************************����Ϊ�μ�����09-12***************************
Private Sub cinb�շ�����_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errhandle
    If Index = �շ�_���ѵ�λ Or Index = �����շ�_���ѵ�λ Then
        If Button = 2 Then
            mint�Ƿ��Ҽ� = 1
        Else
            mint�Ƿ��Ҽ� = 0
        End If
    End If
Exit Sub
errhandle:
    sfsub������ "�շѽ���", "frm�շ�", "cinb�շ�����_MouseDown", Err.Number, Err.Description
End Sub
'***************************����Ϊ�μ�����09-12***************************

'***************************����Ϊ�μ�����09-12***************************
Private Sub cinb�շ�����_Validate(Index As Integer, Cancel As Boolean)
On Error Resume Next
    mint�Ƿ��Ҽ� = 0
End Sub
'***************************����Ϊ�μ�����09-12***************************

Private Sub cind�ֵ�_DblClick(Index As Integer)
On Error Resume Next
    cinb�շ�����_KeyPress mintCurInput, vbKeyReturn
End Sub

Private Sub cind�ֵ�_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Select Case KeyCode
        Case vbKeyEscape
            cind�ֵ�(Index).Visible = False
            Me.KeyPreview = True
        Case vbKeyReturn
            cind�ֵ�_DblClick (Index)
        Case Else
    End Select
End Sub



Private Sub cind�ֵ�_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    mlngX = X
    mlngY = Y
End Sub

Private Sub cind�ֵ�_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errhandle
    If (Button = vbLeftButton) And Y < (cind�ֵ�(Index).RowHeight(0) * cind�ֵ�(Index).Rows - 1) And X < cind�ֵ�(Index).ColPos(cind�ֵ�(Index).Cols - 1) + cind�ֵ�(Index).ColWidth(cind�ֵ�(Index).Cols - 1) Then
        If cind�ֵ�(Index).Top > 0 And (cind�ֵ�(Index).Top + cind�ֵ�(Index).Height) < Me.Height And cind�ֵ�(Index).Left > 0 And (cind�ֵ�(Index).Left + cind�ֵ�(Index).Width) < Me.Width Then
            cind�ֵ�(Index).Move cind�ֵ�(Index).Left + X - mlngX, cind�ֵ�(Index).Top + Y - mlngY
        End If
    End If
    If cind�ֵ�(Index).Top <= 0 Then cind�ֵ�(Index).Top = 1
    If cind�ֵ�(Index).Left <= 0 Then cind�ֵ�(Index).Left = 1
    If cind�ֵ�(Index).Top + cind�ֵ�(Index).Height >= Me.Height Then cind�ֵ�(Index).Top = Me.Height - cind�ֵ�(Index).Height - 1
    If cind�ֵ�(Index).Left + cind�ֵ�(Index).Width >= Me.Width Then cind�ֵ�(Index).Left = Me.Width - cind�ֵ�(Index).Width - 1
Exit Sub
errhandle:
    sfsub������ "�շѽ���", "frm�շ�", "cind�ֵ�_MouseMove", Err.Number, Err.Description
End Sub

Private Sub cing�����嵥_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    Dim lcurMoney As Currency
    
    On Error GoTo errhandle
    ctlb������.Buttons("�շ�(&G)").Enabled = True
    Select Case Col
        Case �����嵥_����
            '�ж�������Ƿ���ֵ
            If Len(cing�����嵥(ctabShoufei.Tab).TextMatrix(Row, Col)) > 4 Then
                cing�����嵥(ctabShoufei.Tab).TextMatrix(Row, Col) = mstrUndoCount
            Else
                If IsNumeric(cing�����嵥(ctabShoufei.Tab).TextMatrix(Row, Col)) And Val(cing�����嵥(ctabShoufei.Tab).TextMatrix(Row, Col)) > 0 Then   '����ֵ
                 '������
                    cing�����嵥(ctabShoufei.Tab).TextMatrix(Row, �����嵥_���) = cing�����嵥(ctabShoufei.Tab).TextMatrix(Row, �����嵥_����) * cing�����嵥(ctabShoufei.Tab).TextMatrix(Row, �����嵥_����)
                Else                                                            '������ֵ
                    'Undo
                    cing�����嵥(ctabShoufei.Tab).TextMatrix(Row, Col) = mstrUndoCount
                End If
            End If
        Case �����嵥_����
            Dim lcur���� As Currency
            If mcur��С���� = mcur��󵥼� Then
                sffuncMsg "���շ���Ŀ�����Ѷ�,�����޸ģ�", sf����
                cing�����嵥(ctabShoufei.Tab).TextMatrix(Row, Col) = mstrUndoMoney
                If cinb�շ�����(mintCurInput).Enabled Then cinb�շ�����(mintCurInput).SetFocus
                ctlb������.Buttons("�շ�(&G)").Enabled = True
                Exit Sub
            End If
            
            If IsNumeric(cing�����嵥(ctabShoufei.Tab).TextMatrix(Row, Col)) Then
                If Val(cing�����嵥(ctabShoufei.Tab).TextMatrix(Row, Col)) > 0 Then
                        If Val(cing�����嵥(ctabShoufei.Tab).TextMatrix(Row, Col)) <= mcur��󵥼� And Val(cing�����嵥(ctabShoufei.Tab).TextMatrix(Row, Col)) >= mcur��С���� Then
                            cing�����嵥(ctabShoufei.Tab).TextMatrix(Row, �����嵥_���) = cing�����嵥(ctabShoufei.Tab).TextMatrix(Row, �����嵥_����) * cing�����嵥(ctabShoufei.Tab).TextMatrix(Row, �����嵥_����)
                        Else
                            sffuncMsg "����ĵ��۳�����Χ��", sf����
                            cing�����嵥(ctabShoufei.Tab).TextMatrix(Row, Col) = mstrUndoMoney
                        End If
                Else
                    cing�����嵥(ctabShoufei.Tab).TextMatrix(Row, Col) = mstrUndoMoney
                End If
            Else
                cing�����嵥(ctabShoufei.Tab).TextMatrix(Row, Col) = mstrUndoMoney
            End If
        Case Else
    End Select
    
    For i = 1 To cing�����嵥(ctabShoufei.Tab).Rows - 1
        lcurMoney = lcurMoney + Val(cing�����嵥(ctabShoufei.Tab).TextMatrix(i, �����嵥_���))
    Next
    mcur�ܽ�� = lcurMoney
    
    '�ж��Ƿ���ѡ�еķ�����Ϣ�вű仯

    If cchk�ڲ��շ�.Value = 0 Then
        cinb�շ�����(Ӧ�ս��) = lcurMoney * Val(cinb�շ�����(���۱���).Text)
        cinb�շ�����(Ӧ�ս���д) = FuncConvertToCapsStr(Val(cinb�շ�����(Ӧ�ս��)))
    End If
    
    '�������ݿ��е�ֵ
     If umfuncУ���û�Ȩ��("�շѹ���_�ڲ��շ���Ϣ�޸�") Then
        If mstr�շѱ�� = "" And cing�����嵥(ctabShoufei.Tab).TextMatrix(cing�����嵥(ctabShoufei.Tab).RowSel, 0) = "" Then
           sffuncMsg "������Ϣ�����޷�ɾ����"
           Exit Sub
        Else
           Dim lstr�շ���Ŀ��� As String       '��¼�շѱ��
           Dim lsing As Currency                '��¼����
           Dim lCurrency As Currency            '��¼���
           Dim lcount As Long                   '��¼����
           
           lstr�շ���Ŀ��� = cing�����嵥(ctabShoufei.Tab).TextMatrix(cing�����嵥(ctabShoufei.Tab).RowSel, 0)
           lsing = cing�����嵥(ctabShoufei.Tab).TextMatrix(cing�����嵥(ctabShoufei.Tab).RowSel, 2)
           lcount = cing�����嵥(ctabShoufei.Tab).TextMatrix(cing�����嵥(ctabShoufei.Tab).RowSel, 3)
           lCurrency = lsing * lcount
           sub�޸��շ���Ŀ���� mstr�շѱ��, lstr�շ���Ŀ���, lsing, lcount, lCurrency
           
          
           '�����޸��շ���Ŀ����ϸ��Ϣ,�޸ĵ��շ���Ŀ���ƣ��޸ĵĽ�� ʱ�䣺2002/09/17�����ߣ��켽��
           Dim lstr�շ���Ŀ���� As String
           Dim lstrIntro As String
           lstrIntro = ""
           lstr�շ���Ŀ���� = cing�����嵥(ctabShoufei.Tab).TextMatrix(cing�����嵥(ctabShoufei.Tab).RowSel, 1)
           lstrIntro = "�շ���Ŀ���ƣ�" & lstr�շ���Ŀ���� & "�����ۣ�" & lsing & "Ԫ" & "��������" & lcount & "����" & lCurrency & " Ԫ"
           
           '���ܣ����޸Ĵ�����Ϣ��ʽ���͡�ʱ�䣺2002/08/05�����ߣ��켽��
           sub��Ϣ���� mstr�շѱ��, "�շѱ��Ϊ��" & mstr�շѱ�� & "�ķ�����Ϣ�ĵ��ۻ������ѱ��޸ġ�" & lstrIntro
        End If
     End If
     
    'ˢ�½���
    For i = 1 To cing�շѻ�����Ϣ��.Rows - 1
        If cing�շѻ�����Ϣ��.Cell(flexcpText, i, 2) = mstr�շѱ�� Then
        cing�շѻ�����Ϣ��.TextMatrix(i, 5) = mcur�ܽ�� * Val(cinb�շ�����(���۱���).Text)
        End If
    Next
    
    Dim lbln�Ƿ���ѡ���� As Boolean
    lbln�Ƿ���ѡ���� = False
    
    For i = 1 To cing�շѻ�����Ϣ��.Rows - 1
        If cing�շѻ�����Ϣ��.Cell(flexcpChecked, i, 0) = 1 Then
            lbln�Ƿ���ѡ���� = True
            Exit For
        End If
    Next
    
    If lbln�Ƿ���ѡ���� = True Then
        sub��������ˢ��
    End If

errhandle:
    If cinb�շ�����(mintCurInput).Enabled Then cinb�շ�����(mintCurInput).SetFocus
    ctlb������.Buttons("�շ�(&G)").Enabled = True
    If Err.Number = 0 Then Exit Sub
    sfsub������ ������, ģ����, "cing�����嵥_AfterEdit", Err.Number, Err.Description
End Sub

'���ܣ����ڲ��շ������ӷ�����Ϣ���շ���Ŀ
'ע������: ��Ϊ������Ϣ�У��շ���Ŀ�е������ض�Ϊ1,����������ͬ
'          �����ڴ���ʱ��������Ӧ��
'ʱ�䣺2002/07/19
'���ߣ��켽��
Private Function func�����շ���Ŀ(ByVal para�շѱ�� As String, ByVal Para�շ���Ŀ��� As String, ByVal Para���� As Currency) As Boolean
On Error GoTo errHanler
    Dim lstrSql As String           '���������¼SQL���
    Dim lstr�շ����� As String      '���������¼�շ�����
    Dim lstr�շѱ�� As String      '���������¼�վݱ��
    Dim lstr������ As String        '���������¼����������
    Dim lstr���ѵ�λ���� As String  '���������¼���ѵ�λ����
    Dim lstr���ѵ�λ��� As String  '���������¼���ѵ�λ���
    Dim lstr�������� As String      '���������¼��������
    Dim lstr���ܿƾ����� As String  '���������¼���ܿƾ�����
    Dim lstr���ܿƱ�� As String    '���������¼���ܿƱ��
    Dim lstrҵ����� As String      '���������¼ҵ�����
    Dim lobjTemp As Object          '������ʱ������¼��

    '��ʼ������
    mblntemp = True
    func�����շ���Ŀ = False

    lstrSql = "select * from �շѹ���_������Ϣ�� where �շѱ��='" & para�շѱ�� & "'"
    Set lobjTemp = dafuncGetData(lstrSql)
    
    '��ȡ�շ���Ϣ
    If lobjTemp.RecordCount > 0 Then
        lstr�շ����� = lobjTemp("�շ�����")
        lstr�շѱ�� = lobjTemp("�շѱ��")
        lstr������ = IIf(IsNull(lobjTemp("������")), "��������", lobjTemp("������"))
        lstr���ѵ�λ���� = IIf(IsNull(lobjTemp("���ѵ�λ����")), "���굥λ", lobjTemp("���ѵ�λ����"))
        lstr���ѵ�λ��� = IIf(IsNull(lobjTemp("���ѵ�λ���")), "���굥λ���", lobjTemp("���ѵ�λ���"))
        lstr�������� = IIf(IsNull(lobjTemp("��������")), "�������ڲ���", lobjTemp("��������"))
        lstr���ܿƾ����� = IIf(IsNull(lobjTemp("���ܿ��Ҿ�����")), "��������", lobjTemp("���ܿ��Ҿ�����"))
        lstr���ܿƱ�� = IIf(IsNull(lobjTemp("���ܿ��ұ��")), "������", lobjTemp("���ܿ��ұ��"))
        lstrҵ����� = IIf(IsNull(lobjTemp("ҵ�����")), "������", lobjTemp("ҵ�����"))
        
        '�����еķ�����Ϣ����������շ���Ŀ��Ϣ�������ݿ��в���һ���շ���Ŀ����
        lstrSql = "insert into �շѹ���_������Ϣ�� (�շ�����,�շѱ��,�շ���Ŀ���,����,����," & _
                  "���,������,���ѵ�λ����,���ѵ�λ���,��������,���ܿ��Ҿ�����,���ܿ��ұ��,ҵ�����) " & _
                  " values ( '" & lstr�շ����� & "','" & lstr�շѱ�� & "','" & Para�շ���Ŀ��� & "'," & _
                   1 & " ," & Para���� & "," & Para���� & ",'" & lstr������ & "','" & _
                   lstr���ѵ�λ���� & "','" & lstr���ѵ�λ��� & "','" & lstr�������� & "','" & _
                   lstr���ܿƾ����� & "','" & lstr���ܿƱ�� & "','" & lstrҵ����� & "')"
        dafuncGetData (lstrSql)
         
    Else
        Exit Function
    End If
    func�����շ���Ŀ = True
Exit Function
errHanler:
    func�����շ���Ŀ = False
    sfsub������ ������, ģ����, "sub�����շ���Ŀ", Err.Number, Err.Description
End Function


Private Sub sub�޸��շ���Ŀ����(ByVal para�շѱ�� As String, ByVal Para��Ŀ��� As String, ByVal Para���� As Double, ByVal Para���� As Long, ByVal Para��� As Double)
On Error GoTo errhandler
    Dim lstrSql As String           '���������¼SQL���
    
    lstrSql = "update  �շѹ���_������Ϣ�� set ����= '" & Para���� & "'," & _
              " ����=convert(money,'" & Para���� & "'), ���=convert(money,'" & Para��� & "')" & _
              " where �շѱ��='" & para�շѱ�� & "' and �շ���Ŀ���='" & Para��Ŀ��� & "'"
    dafuncGetData (lstrSql)
Exit Sub
errhandler:
    sfsub������ ������, ģ����, "sub�޸��շ���Ŀ����", Err.Number, Err.Description
End Sub

Private Sub cing�����嵥_BeforeEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error Resume Next
    '���ܣ�����Ȩ�����������ڲ��շ���Ϣ���޸�.
    '���ߣ��켽��
    'ʱ�䣺2002/07/01
    If cchk�ڲ��շ�.Value = vbChecked And ctabShoufei.Tab = 0 Then
        If umfuncУ���û�Ȩ��("�շѹ���_�ڲ��շ���Ϣ�޸�") Then
            Cancel = False
        Else
            Cancel = True
            'sffuncMsg "û���޸��ڲ��շѵ�Ȩ�ޣ�"
            Exit Sub
        End If
    End If
    
    Select Case Col
        Case �����嵥_����
            ctlb������.Buttons("�շ�(&G)").Enabled = False
            mstrUndoCount = cing�����嵥(ctabShoufei.Tab).TextMatrix(Row, Col)
            
        Case �����嵥_����
            ctlb������.Buttons("�շ�(&G)").Enabled = False
            mrds�շ���Ŀ.MoveFirst
            'mrds�շ���Ŀ.find "�շ���Ŀ���=" & cind�ֵ�(�ֵ�_�շ���Ŀ).TextMatrix(cind�ֵ�(�ֵ�_�շ���Ŀ).RowSel, �ֵ�_�շ���Ŀ���)
            '�޸�:��ȡ�շѱ�ţ��ӽ����ϵı���л��. ʱ�䣺2002/02/28 �켽��
            mrds�շ���Ŀ.find "�շ���Ŀ���=" & cing�����嵥(ctabShoufei.Tab).TextMatrix(Row, 0)
            If mrds�շ���Ŀ.RecordCount > 0 Then
                mcur��С���� = mrds�շ���Ŀ("��С����").Value
                mcur��󵥼� = mrds�շ���Ŀ("��󵥼�").Value
            Else
                sffuncMsg "δ�ҵ����շ���Ŀ��������Ϣ����������Ϣ�����ѱ��޸Ļ�ɾ�������˳��շѽ��棬���½��룡"
            End If
            mstrUndoMoney = cing�����嵥(ctabShoufei.Tab).TextMatrix(Row, Col)
        Case Else
            ctlb������.Buttons("�շ�(&G)").Enabled = True
            Cancel = True
    End Select
End Sub



Private Sub cing�����嵥_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next
' ���ܣ����ڲ��շѲ�����ɾ��������Ϣ��ʱ�䣺2002/02/6 �켽��
If cchk�ڲ��շ�.Value = vbChecked And ctabShoufei.Tab = 0 And KeyCode = vbKeyDelete Then
    If umfuncУ���û�Ȩ��("�շѹ���_�ڲ��շ���Ϣ�޸�") Then
        Select Case KeyCode
        Case vbKeyDelete
            mobj����ͨ�ö���_BeforeOperate "ɾ��", False
        Case vbKeyEscape
            cind�ֵ�(�ֵ�_�շ���Ŀ).Visible = False
            cind�ֵ�(�ֵ�_�շѱ�׼).Visible = False
            Me.KeyPreview = True
        End Select
    Else
        sffuncMsg "û���޸��ڲ��շѵ�Ȩ�ޣ�"
    End If
Else
    Select Case KeyCode
        Case vbKeyDelete
            mobj����ͨ�ö���_BeforeOperate "ɾ��", False
        Case vbKeyEscape
            cind�ֵ�(�ֵ�_�շ���Ŀ).Visible = False
            cind�ֵ�(�ֵ�_�շѱ�׼).Visible = False
            Me.KeyPreview = True
    End Select
End If
End Sub

Private Sub cing�����嵥_LostFocus(Index As Integer)
On Error Resume Next
    ctlb������.Buttons("�շ�(&G)").Enabled = True
End Sub

Private Sub cing�շѻ�����Ϣ��_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error Resume Next
    If cing�շѻ�����Ϣ��.Cell(flexcpChecked, cing�շѻ�����Ϣ��.RowSel, 0) = 2 Then
        mcur������Ϣ�ܽ�� = mcur������Ϣ�ܽ�� - Val(cing�շѻ�����Ϣ��.TextMatrix(cing�շѻ�����Ϣ��.RowSel, ������Ϣ_���))
    Else
        mcur������Ϣ�ܽ�� = mcur������Ϣ�ܽ�� + Val(cing�շѻ�����Ϣ��.TextMatrix(cing�շѻ�����Ϣ��.RowSel, ������Ϣ_���))
    End If
    cinb�շ�����(Ӧ�ս��).Text = mcur������Ϣ�ܽ�� * (cinb�շ�����(���۱���).Text)
End Sub

Private Sub cing�շѻ�����Ϣ��_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error Resume Next
    If Col > 0 Then Cancel = True
End Sub

'���ܣ�ˢ�½����Ӧ�ս��
'���ߣ��켽��
'ʱ�䣺2002/07/22
Private Sub sub��������ˢ��()
On Error GoTo errhander
    Dim i As Integer
    Dim lcurMoney As Currency
    lcurMoney = 0
    For i = 1 To cing�շѻ�����Ϣ��.Rows - 1
        If cing�շѻ�����Ϣ��.Cell(flexcpChecked, i, 0) = 1 Then
            lcurMoney = lcurMoney + Val(cing�շѻ�����Ϣ��.TextMatrix(i, 5))
        End If
    Next
    
    mcur������Ϣ�ܽ�� = lcurMoney
    
    cinb�շ�����(Ӧ�ս��) = lcurMoney * Val(cinb�շ�����(���۱���).Text)
    cinb�շ�����(Ӧ�ս���д) = FuncConvertToCapsStr(Val(cinb�շ�����(Ӧ�ս��)))
Exit Sub
errhander:
    sfsub������ "�շѽ���", "frm�շ�", "sub��������ˢ��", Err.Number, Err.Description
End Sub



Private Sub cing�շѻ�����Ϣ��_Click()
    Dim lrds��ϸ���� As Recordset
    Dim lrds���۱��� As Recordset
    Dim i As Long
    On Error GoTo errhandle
    If cing�շѻ�����Ϣ��.RowSel < 1 Then Exit Sub
    cing�����嵥(�շ�).Rows = 1
    Set lrds��ϸ���� = mobj�շѹ���.funcExecute("select a.�շ�����,a.�շѱ��,a.�շ���Ŀ���,b.�շ���Ŀ����,a.����,a.����,b.������λ,a.���,a.���ѵ�λ���,a.���ѵ�λ����,a.������,d.Ƭ��,���ܿ������� = c.����" & _
                                                " from  �շѹ���_������Ϣ�� a left join �շѹ���_�շ���Ŀ�ֵ�� b on a.�շ���Ŀ��� = b.�շ���Ŀ���" & _
                                                " left join ϵͳ����_�����ֵ�� c on a.���ܿ��ұ��=c.���" & _
                                                " left join ��λ����_��λ������Ϣ�� d on a.���ѵ�λ���=d.������" & _
                                                " where a.�շ�״̬= 0 and a.�շѱ��='" & cing�շѻ�����Ϣ��.TextMatrix(cing�շѻ�����Ϣ��.RowSel, ������Ϣ_�շѱ��) & "'", "cls������Ϣ")
                                                
                                                
    '��¼��ǰ���շ�����
    mstr�շѱ�� = cing�շѻ�����Ϣ��.TextMatrix(cing�շѻ�����Ϣ��.RowSel, ������Ϣ_�շѱ��)
    
    If lrds��ϸ���� Is Nothing Then Exit Sub
    If lrds��ϸ����.BOF And lrds��ϸ����.EOF Then Exit Sub
    lrds��ϸ����.MoveFirst
    For i = 0 To lrds��ϸ����.RecordCount - 1
        cing�����嵥(�շ�).AddItem lrds��ϸ����("�շ���Ŀ���") & vbTab & _
                                   lrds��ϸ����("�շ���Ŀ����") & vbTab & _
                                   lrds��ϸ����("����") & vbTab & _
                                   lrds��ϸ����("����") & vbTab & _
                                   lrds��ϸ����("���")
        If Not lrds��ϸ����.EOF Then lrds��ϸ����.MoveNext
    Next
    lrds��ϸ����.MoveFirst
    mstr���ѵ�λ��� = lrds��ϸ����("���ѵ�λ���").Value
    '�켽�����޸�2001/12/20���������ڽ����϶��շѱ�ŵ���ʾ
    cinb�շ�����(�շ�_�շѱ��).Text = lrds��ϸ����("�շѱ��")
    cinb�շ�����(�շ�_���ܿ���).Text = lrds��ϸ����("���ܿ�������")
    
    
    '��ʾƬ����Ϣ
    If IIf(IsNull(lrds��ϸ����("Ƭ��").Value), "", lrds��ϸ����("Ƭ��").Value) = "" Then
        ClabƬ��.Caption = "Ƭ����(����)"
    Else
        ClabƬ��.Caption = "(" + lrds��ϸ����("Ƭ��").Value + ")"
    End If
    
    cinb�շ�����(�շ�_������).Text = lrds��ϸ����("������")
    cinb�շ�����(�շ�_���ѵ�λ).Text = IIf(IsNull(lrds��ϸ����("���ѵ�λ����").Value), "", lrds��ϸ����("���ѵ�λ����").Value)
    Set lrds���۱��� = mobj�շѹ���.funcExecute("select ���۱��� from �շѹ���_������Ϣ�� where ��λ���='" & mstr���ѵ�λ��� & "'", "cls������Ϣ")
    If lrds���۱���.BOF And lrds���۱���.EOF Then
        cinb�շ�����(���۱���).Text = "1.00"
    Else
        cinb�շ�����(���۱���).Text = Format(lrds���۱���("���۱���").Value, "0.00")
    End If
errhandle:
    If Err.Number = 0 Then Exit Sub
    sfsub������ "�շѽ���", "frm�շ�", "cing�շѻ�����Ϣ��_Click", Err.Number, Err.Description
End Sub



Private Sub ctabShoufei_Click(PreviousTab As Integer)
On Error GoTo errhandle
    cind�ֵ�(�ֵ�_�շѱ�׼).Visible = False
    cind�ֵ�(�ֵ�_�շ���Ŀ).Visible = False
    cind�ֵ�(�ֵ�_���ܿ���).Visible = False
    

    If PreviousTab = 1 Then
        Ccbo�շ���Ŀ����.ListIndex = 0
        Ccbo�շ���Ŀ����.Refresh
    Else
        Ccbo���շѴ���.ListIndex = 0
        Ccbo���շѴ���.Refresh
    End If
    
    If ctabShoufei.Tab = �շ� And cchk�ڲ��շ�.Value = 1 Then
        ctlb������.Buttons("��ѯ(&Q)").Enabled = True
        
        '�޸�:�ڴ������շѽ������,�����շ���Ŀ�Ŀؼ�Ӧ��Ϊ������ �켽�� 2002/11/26
        Ccbo�շ���Ŀ����.Enabled = False
        cinb�շ�����(5).Enabled = False
    Else
        ctlb������.Buttons("��ѯ(&Q)").Enabled = False
    End If
    sub�������
Exit Sub
errhandle:
    sfsub������ "�շѽ���", "ctabShoufei_Click", "cing�շѻ�����Ϣ��_Click", Err.Number, Err.Description
End Sub

Private Sub Ctim_Timer()
    Dim i As Long
    Dim j As Long
    Dim lobjRec As Object           '���������¼�����
    Dim lstrSql As String              '���������¼SQL���
    
    On Error GoTo errhandler
    
    '��cind�ֵ�(�ֵ�_�շ���Ŀ)
    Ctim.Enabled = False
    Me.MousePointer = 11
    Me.cstuShoufei.Panels(1) = "���ڼ��ػ������ݣ����Ժ�..."
    Me.Enabled = False
    
    '����:���շ���Ŀ������,��������,���շ���Ŀ�и����շ���Ŀ����������.
    'ʱ��:2002/06/21
    '����:�켽��
    lstrSql = "select �շ���Ŀ���,�շ���Ŀ���� from �շѹ���_�շ���Ŀ�ֵ�� where len(�շ���Ŀ���)=3 " & _
           " order by �շ���Ŀ��� "
    
    Set lobjRec = dafuncGetData(lstrSql)
        
    Do While Not lobjRec.EOF
        Ccbo�շ���Ŀ����.AddItem lobjRec("�շ���Ŀ����")
        Ccbo���շѴ���.AddItem lobjRec("�շ���Ŀ����")
        lobjRec.MoveNext
    Loop
    
    Ccbo�շ���Ŀ����.ListIndex = 0
    Ccbo�շ���Ŀ����.Refresh
        
errhandler:
    Me.Enabled = True
    cind�ֵ�(�ֵ�_�շ���Ŀ).Redraw = True
    subEnable�շ�����
    Me.cstuShoufei.Panels(1) = ""
    Me.MousePointer = 0
    Exit Sub
    
End Sub

'����:���շ���Ŀ�����м�������.
'����:Para�շ���Ŀ����
'�����켽��:
'ʱ��:2002/06/21

Private Sub sub�շ���Ŀ���(ByVal Para�շ���Ŀ���� As String)
On Error GoTo errhandler
    Dim lstrSql As String            '���������¼SQL���
    Dim lobjRec As Object            '���������¼���ݼ�
    Dim lstrtemp As String           '���������¼�շѱ��ǰ׺��
    Dim i As Integer                 '����ѭ������
    Dim j As Integer                 '����ѭ������
    Dim lInt As Integer              '�����¼������
    Dim lobjRecCount As Object       '�����¼��������
    Dim lInt���� As Integer
    Dim lbln��ʶ As Boolean
    
    '�շѴ�������Ϊ�մ�,�˳��ù���
    If Para�շ���Ŀ���� = "" Then
        Exit Sub
    End If
    
    '�����շ���Ŀ��������,��ȡ�շѱ��ǰ׺
    lstrSql = "select �շ���Ŀ��� from �շѹ���_�շ���Ŀ�ֵ�� where �շ���Ŀ����= '" & Para�շ���Ŀ���� & "'"
    Set lobjRec = dafuncGetData(lstrSql)
    lstrtemp = Left$(lobjRec("�շ���Ŀ���"), 3)
    
    '��ȡ��¼������
    lstrSql = "select count(*) as ��¼���� from �շѹ���_�շ���Ŀ�ֵ�� where left(�շ���Ŀ���,3)='" & lstrtemp & "'"
    Set lobjRecCount = dafuncGetData(lstrSql)
    lInt = lobjRecCount("��¼����")
    
    '�������м����շ���Ŀ
    cind�ֵ�(�ֵ�_�շ���Ŀ).Redraw = True
    cind�ֵ�(�ֵ�_�շ���Ŀ).Clear
    
    mrds�շ���Ŀ.MoveFirst
        
    If Not (mrds�շ���Ŀ Is Nothing) Then
        cind�ֵ�(�ֵ�_�շ���Ŀ).Cols = mrds�շ���Ŀ.Fields.Count

        cind�ֵ�(�ֵ�_�շ���Ŀ).Refresh
        For i = 0 To mrds�շ���Ŀ.Fields.Count - 1
            cind�ֵ�(�ֵ�_�շ���Ŀ).TextMatrix(0, i) = mrds�շ���Ŀ(i).Name
        Next
        
        lInt���� = 1
        lbln��ʶ = False
        If (Not mrds�շ���Ŀ.BOF) And (Not mrds�շ���Ŀ.EOF) Then
            cind�ֵ�(�ֵ�_�շ���Ŀ).Rows = lInt
            mrds�շ���Ŀ.MoveFirst
            For i = 0 To mrds�շ���Ŀ.RecordCount - 1
                
                If Left$(mrds�շ���Ŀ("�շ���Ŀ���"), 3) = lstrtemp Then
                    lbln��ʶ = True
                    For j = 0 To mrds�շ���Ŀ.Fields.Count - 1
                        '�����ܣ��������ݿ���ܴ��ڵĿ�ֵ,����ת��Ϊ""��ʱ�䣺2002/01/27,�켽����
                        cind�ֵ�(�ֵ�_�շ���Ŀ).TextMatrix(lInt����, j) = IIf(IsNull(mrds�շ���Ŀ(j)), "", mrds�շ���Ŀ(j))
                        cind�ֵ�(�ֵ�_�շ���Ŀ).AutoSize j
                    Next j
                    
                End If
                If Not mrds�շ���Ŀ.EOF Then mrds�շ���Ŀ.MoveNext
                If lbln��ʶ = True Then
                    lInt���� = lInt���� + 1
                End If
            Next i
        End If
        
    End If
    cind�ֵ�(�ֵ�_�շ���Ŀ).Width = 165
    For i = 0 To cind�ֵ�(�ֵ�_�շ���Ŀ).Cols - 1
        cind�ֵ�(�ֵ�_�շ���Ŀ).Width = cind�ֵ�(�ֵ�_�շ���Ŀ).Width + cind�ֵ�(�ֵ�_�շ���Ŀ).ColWidth(i)
    Next
    
    Exit Sub
errhandler:
    Me.Enabled = True
    cind�ֵ�(�ֵ�_�շ���Ŀ).Redraw = True
    subEnable�շ�����
    Me.cstuShoufei.Panels(1) = ""
    Me.MousePointer = 0
    Exit Sub
End Sub

Private Sub cupd�޸Ĵ��۱���_DownClick()
On Error Resume Next
    If Val(cinb�շ�����(���۱���).Text) > 0 Then
        cinb�շ�����(���۱���).Text = Format(CStr(Val(cinb�շ�����(���۱���).Text) - 0.01), "0.00")
    Else
        cinb�շ�����(���۱���).Text = "0.00"
    End If
End Sub
Private Sub cupd�޸Ĵ��۱���_GotFocus()
On Error Resume Next
    cind�ֵ�(�ֵ�_�շѱ�׼).Visible = False
    cind�ֵ�(�ֵ�_�շ���Ŀ).Visible = False
    cind�ֵ�(�ֵ�_���ܿ���).Visible = False
End Sub

Private Sub cupd�޸Ĵ��۱���_UpClick()
On Error GoTo errhandle
    If Val(cinb�շ�����(���۱���).Text) < 1 Then
        cinb�շ�����(���۱���).Text = Format(CStr(Val(cinb�շ�����(���۱���).Text) + 0.01), "0.00")
    Else
        cinb�շ�����(���۱���).Text = "1.00"
    End If
Exit Sub
errhandle:
    sfsub������ "�շѹ���������", "frm�շ�", "Form_UpClick", Err.Number, Err.Description
End Sub

Private Sub Form_Activate()
On Error GoTo errhandle
    If mblnʹ�� Then
        ctabShoufei.Tab = �շ�
    End If
Exit Sub
errhandle:
    sfsub������ "�շѹ���������", "frm�շ�", "Form_Activate", Err.Number, Err.Description
End Sub

Private Sub Form_Load()
Dim lcol������ As Collection
On Error GoTo errhandle
    If pblnInUse Then Exit Sub
    pblnInUse = True
    mblnʹ�� = True
        
    Set mobj����ͨ�ö��� = New cls����ͨ�ö���
    Set mobj����ͨ�ö���.Form = Me
    Set mobj����ͨ�ö���.c������ = ctlb������
    
    Set lcol������ = New Collection
    
    lcol������.Add "��ѯ(&Q)105"
    lcol������.Add "|"
    lcol������.Add "�շ�(&G)123"
    lcol������.Add "|"
    lcol������.Add "ɾ��"
    lcol������.Add "���"
    lcol������.Add "����(&T)122"
    lcol������.Add "|"
    lcol������.Add "�˳�"
    
    mobj����ͨ�ö���.subInitialize lcol������, ""
    Set lcol������ = Nothing
    
    '����:�ڽ����ʼ��ʱ,Ĭ��Ϊ�����ڲ��շ�,����cing�շѻ�����Ϣ���
    '�޸���:�켽��
    '�޸�ʱ��:2002/06/21
    
    cing�շѻ�����Ϣ��.Visible = False
    Frame5.Left = 120
    Frame5.Width = 10845
    cing�����嵥(0).Width = 10600
    
    Clabҵ�����.Enabled = False
    Ccboҵ�����.Enabled = False
    Ccboҵ�����.BackColor = &H80000000
    
    '����Ϊ�μ�����
    If Not func��ȡ��ʼ������ Then
        mblnʹ�� = False
        ctlb������.Buttons(1).Enabled = False
        ctlb������.Buttons(3).Enabled = False
        ctlb������.Buttons(5).Enabled = False
        ctlb������.Buttons(6).Enabled = False
        ctlb������.Buttons(7).Enabled = False
        ctabShoufei.Visible = False
        Frame6.Visible = False
        Frame7.Visible = False
        Exit Sub
    End If
    
    
    sub��ʼ������
    '����Ϊ�μ�����
    
    ccmdѡ��.Enabled = False
    Copt����.Enabled = False
    Copt����.Enabled = False
    Copt����.Enabled = False
    Cchk��������Ϣ��ѯ.Enabled = False
    ctlb������.Buttons(1).Enabled = False
    ctlb������.Buttons(7).Enabled = False
    mstr�շѱ�� = ""
    mblntemp = False
    '���³�ʼ�����붨ʱ����
    Ctim.Enabled = True
    
    '�޸ģ�2002-10-17����������ӡǰԤ����
    On Error Resume Next
    Set mobj���� = New cls�û���������
    mobj����.�û���� = um�û����
    mobj����.ҵ���� = "�շѹ���"
    If mobj����.������ֵ("��ӡǰԤ��") = "��" Then
        cchkԤ��.Value = 1
    End If
    
    Exit Sub
errhandle:
    sfsub������ "�շѹ���������", "frm�շ�", "Form_Load", Err.Number, Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    pblnInUse = False
    Set mrds�շ���Ŀ = Nothing
    Set mrds�շѱ�׼ = Nothing
    Set mobj����ͨ�ö��� = Nothing
    Set mobj�շѹ��� = Nothing
    Set mobjҵ������ = Nothing
    Set mobj��λ���� = Nothing
    
    '�޸ģ�2002-10-17�����������������
    mobj����.sub���Ǽ���ֵ "��ӡǰԤ��", IIf(cchkԤ��.Value = 1, "��", "��")
    
End Sub


'&  ---------------| sub������� |----------------------
'&  ��;��  �������
'&  ���ߣ�  Shadow
'&  ������ڣ�2001/4/1
Private Sub sub�������()
    Dim i As Integer
    '�켽����2001/12/21����������ؼ�״̬��ʶ
    Dim l״̬��ʶ As Integer
    
    On Error GoTo errhandle
    mcur�ܽ�� = 0
    mcur������Ϣ�ܽ�� = 0
    l״̬��ʶ = 0
    
    '�켽����2001/12/21��:Ϊ�������������棬�ı����ؼ����ԣ�����¼��ʶֵ
    If Frame4.Enabled = False Then
        Frame4.Enabled = True
        l״̬��ʶ = 1
    End If
  
  
    '�ڱ������Ҫ����շѱ��;�켽��;2002/9/30
    mstr���ѵ�λ��� = ""
  
    If ctabShoufei.Tab = �շ� Then
        For i = �շ�_�շѱ�� To �շ�_�շѱ�׼
            cinb�շ�����(i).Text = ""
        Next
        'cchk�ڲ��շ�.Value = 0
        cchkͬ���շ�.Value = 0
        cing�շѻ�����Ϣ��.Rows = 1
        cing�����嵥(�շ�).Rows = 1
        If cinb�շ�����(�շ�_�շѱ��).Enabled Then
            cinb�շ�����(�շ�_�շѱ��).SetFocus
        Else
            cinb�շ�����(�շ�_������).SetFocus
        End If
    Else
        For i = �����շ�_������ To �����շ�_�շѱ�׼
            cinb�շ�����(i).Text = ""
        Next
        
        With cdtp����(��Ժ)
            .Year = Year(Date)
            .Month = Month(Date)
            .Day = Day(Date)
        End With
        
        With cdtp����(��Ժ)
            .Year = Year(Date)
            .Month = Month(Date)
            .Day = Day(Date)
        End With
        cing�����嵥(�����շ�).Rows = 1
        If cinb�շ�����(�����շ�_������).Enabled Then cinb�շ�����(�����շ�_������).SetFocus
    End If
    
    cinb�շ�����(���۱���).Text = "1.00"
    
    For i = Ӧ�ս�� To ��������
        cinb�շ�����(i).Text = ""
    Next
    
    cind�ֵ�(�ֵ�_�շ���Ŀ).Visible = False
    cind�ֵ�(�ֵ�_�շѱ�׼).Visible = False
    cind�ֵ�(�ֵ�_���ܿ���).Visible = False

    cinb�շ�����(��������).Text = Date
    cinb�շ�����(���۱���).Text = "1.00"
    cinb�շ�����(�����շ�_���ܿ���).Text = um�û���������
    cinb�շ�����(�շ�_���ܿ���).Text = um�û���������
    
    '�켽����2001/12/21��:���ݱ�ʶֵ�ָ�����ؼ�����
    If l״̬��ʶ = 1 Then
        Frame4.Enabled = False
    Else
        Frame4.Enabled = True
    End If
Exit Sub
errhandle:
    sfsub������ ������, ģ����, "sub����շѽ���", Err.Number, Err.Description
End Sub


'&  ---------------| func��ѯ������Ϣ |----------------------
'&  ��;��  ��������ѯ��¼
'&  ���أ�  Recordset
'&  ���ߣ�  Shadow
'&  ������ڣ�2001/4/1
Private Function func��ѯ������Ϣ() As Recordset
    Dim lstr��ѯ���� As String
    Dim i As Integer
    
    On Error GoTo errhandle
        
    If ctabShoufei.Tab = �շ� Then
        For i = �շ�_�շѱ�� To �շ�_���ܿ���
            If cinb�շ�����(i).Text <> "" Then
                If lblCaption(i).Caption = "���ѵ�λ" Then
                    lstr��ѯ���� = lstr��ѯ���� & "���ѵ�λ���" & "='" & mstr���ѵ�λ��� & "' and "
                ElseIf lblCaption(i).Caption = "���ܿ���" Then
                    lstr��ѯ���� = lstr��ѯ���� & "���ܿ��ұ��" & "='" & mstr���ܿ��ұ�� & "' and "
                Else
                    lstr��ѯ���� = lstr��ѯ���� & lblCaption(i).Caption & "='" & cinb�շ�����(i).Text & "' and "
                End If
            End If
        Next
    Else
        For i = �����շ�_������ To �����շ�_���ܿ���
            If cinb�շ�����(i).Text <> "" Then lstr��ѯ���� = lstr��ѯ���� & lblCaption(i).Caption & "='" & cinb�շ�����(i).Text & "' and "
        Next
    End If
    
    If (lstr��ѯ���� = vbNullString) Or (lstr��ѯ���� = vbNullChar) Then
        sffuncMsg "����������������ѯ������", sf����
        Exit Function
    Else
        lstr��ѯ���� = lstr��ѯ���� & "�շ�״̬=0"
        Set func��ѯ������Ϣ = mobj�շѹ���.funcExecute("select a.�շ�����,a.�շѱ��,a.�շ���Ŀ���,�շ���Ŀ����=(select �շ���Ŀ���� from �շѹ���_�շ���Ŀ�ֵ�� where �շ���Ŀ���=a.�շ���Ŀ���),a.����,������λ=(select ������λ from �շѹ���_�շ���Ŀ�ֵ�� where �շ���Ŀ���=a.�շ���Ŀ���),a.����,a.���,a.�շ�״̬,a.���ѷ�ʽ,a.������,a.���ѵ�λ���,���ѵ�λ=(select ��λ���� from ��λ����_��λ������Ϣ�� where ������=a.���ѵ�λ���),a.��������,a.�˷�����,�շ��˱��=a.�շ���,�շ���=(select ���� from ϵͳ����_Ա��������Ϣ�� where ���=a.�շ���),�˷��˱��=a.�˷���,�˷���=(select ���� from ϵͳ����_Ա��������Ϣ�� where ���=a.�˷���) ,���ܿ��Ҿ����˱��=a.���ܿ��Ҿ�����,���ܿ��Ҿ�����=(select ���� from ϵͳ����_Ա��������Ϣ�� where ���=a.���ܿ��Ҿ�����),���ܿ��ұ��,���ܿ���=(select ���� from ϵͳ����_�����ֵ�� where ���=a.���ܿ��ұ��),���۱���  from �շѹ���_������Ϣ�� a where " & lstr��ѯ����, "cls������Ϣ")
                                            
    End If
    Exit Function
    
errhandle:
    Set func��ѯ������Ϣ = Nothing
    sfsub������ ������, ģ����, "func��ѯ������Ϣ", Err.Number, Err.Description
End Function

'&  ---------------| func�ռ����� |----------------------
'&  ��;��  �ӽ������ռ����ݲ���ϳɼ���,�����ҵ����� "mobj�շѹ���"
'&  ���أ�  Collection
'&  ���ߣ�  Shadow
'&  ������ڣ�2001/4/1
Private Function func�ռ�����() As Collection
    Dim i As Long
    Dim j As Integer
    Dim lcol��¼ As Collection
    Dim lcol���� As Collection
    Dim lintTab As Integer              'ctabShoufei�ĵ�ǰҳ
    Dim lrds���ܿ��� As Recordset
    On Error GoTo errhandle
    lintTab = ctabShoufei.Tab           'ȡ��ctabShoufei�ĵ�ǰҳ
    '��֯����
    If cing�����嵥(lintTab).Rows = 1 Then
        sffuncMsg "�޿��õķ�����Ϣ��", sf����
        GoTo WayOut
    Else
        Set lcol���� = New Collection
        For i = 1 To cing�����嵥(lintTab).Rows - 1
            Set lcol��¼ = New Collection
            '��һ�����ü�¼װ�뼯�� "lcol��¼"
            For j = 0 To cing�����嵥(lintTab).Cols - 1
                '��������е��ֶ�(���շ���Ŀ��š��������ۡ�����������������)
                lcol��¼.Add cing�����嵥(lintTab).TextMatrix(i, j), cing�����嵥(lintTab).TextMatrix(0, j)
            Next
            If lintTab = �շ� Then
                '����շ������ֶ�
                For j = �շ�_������ To �շ�_���ܿ���
                    If lblCaption(j).Caption = "���ѵ�λ" Then
                        lcol��¼.Add mstr���ѵ�λ���, "���ѵ�λ���"
                        lcol��¼.Add cinb�շ�����(�շ�_���ѵ�λ).Text, "���ѵ�λ����"
                    ElseIf lblCaption(j).Caption = "���ܿ���" Then
                        Set lrds���ܿ��� = mobj�շѹ���.funcExecute("select ��� from ϵͳ����_�����ֵ�� where ����='" & cinb�շ�����(�շ�_���ܿ���).Text & "'", "cls������Ϣ")
                        If lrds���ܿ���.BOF And lrds���ܿ���.EOF Then
                            sffuncMsg "��������ܿ��Ҳ������õĿ��ҷ�Χ�ڣ�", sf����
                            If cinb�շ�����(�շ�_���ܿ���).Enabled Then cinb�շ�����(�շ�_���ܿ���).SetFocus
                            Set lrds���ܿ��� = Nothing
                            Set func�ռ����� = Nothing
                            Exit Function
                        Else
                            mstr���ܿ��ұ�� = lrds���ܿ���("���").Value
                        End If
                        
                        lcol��¼.Add mstr���ܿ��ұ��, "���ܿ��ұ��"
                    Else
                        lcol��¼.Add cinb�շ�����(j).Text, lblCaption(j).Caption
                    End If
                Next
                lcol��¼.Add um�û����, "���ܿ��Ҿ�����"                     '������Ҫ�޸�
                
                '*****************����Ϊ�μ��޸ġ���0808*******************
                mrds���ѷ�ʽ.Filter = "����='" & cmb���ѷ�ʽ.Text & "'"
                mint���ѷ�ʽ��� = mrds���ѷ�ʽ("���").Value
                '*****************����Ϊ�μ��޸ġ���0808*******************
                
            Else
                '��������շ������ֶ�
                For j = �����շ�_������ To �����շ�_���ܿ���
                    If lblCaption(j).Caption = "���ѵ�λ" Then
                        lcol��¼.Add mstr���ѵ�λ���, "���ѵ�λ���"
                        lcol��¼.Add cinb�շ�����(�����շ�_���ѵ�λ).Text, "���ѵ�λ����"
                    ElseIf lblCaption(j).Caption = "���ܿ���" Then
                        Set lrds���ܿ��� = mobj�շѹ���.funcExecute("select ��� from ϵͳ����_�����ֵ�� where ����='" & cinb�շ�����(�����շ�_���ܿ���).Text & "'", "cls������Ϣ")
                        If lrds���ܿ���.BOF And lrds���ܿ���.EOF Then
                            sffuncMsg "��������ܿ��Ҳ������õĿ��ҷ�Χ�ڣ�", sf����
                            If cinb�շ�����(�շ�_���ܿ���).Enabled Then cinb�շ�����(�շ�_���ܿ���).SetFocus
                            Set lrds���ܿ��� = Nothing
                            Set func�ռ����� = Nothing
                            Exit Function
                        Else
                            mstr���ܿ��ұ�� = lrds���ܿ���("���").Value
                        End If
                        
                        lcol��¼.Add mstr���ܿ��ұ��, "���ܿ��ұ��"
                    Else
                        lcol��¼.Add cinb�շ�����(j).Text, lblCaption(j).Caption
                    End If
                Next
                
                lcol��¼.Add um�û����, "���ܿ��Ҿ�����"
                lcol��¼.Add 0, "�շ�״̬"
                lcol��¼.Add cdtp����(��Ժ).Value 'CDate(cdtp����(��Ժ).Year & "/" & cdtp����(��Ժ).Month & "/" & cdtp����(��Ժ).Day), "��Ժ����"
                lcol��¼.Add cdtp����(��Ժ).Value 'CDate(cdtp����(��Ժ).Year & "/" & cdtp����(��Ժ).Month & "/" & cdtp����(��Ժ).Day), "��Ժ����"
            End If
            '�����з��ü�¼װ�뼯�� "lcol����"
            lcol����.Add lcol��¼
        Next
    End If
    Set func�ռ����� = lcol����
    GoTo WayOut
    
errhandle:
    Set func�ռ����� = Nothing
    sfsub������ ������, ģ����, "func�ռ�����", Err.Number, Err.Description, True
WayOut:
    Set lcol��¼ = Nothing
    Set lcol���� = Nothing
    Set lrds���ܿ��� = Nothing
End Function

'�޸ģ�2002-6-25����������վݺš�
Private Function func�ռ�ȷ����Ϣ() As Collection
    Dim lstr�վݺ�  As String
    
    On Error Resume Next
    Set func�ռ�ȷ����Ϣ = New Collection
    With func�ռ�ȷ����Ϣ
        .Add mstr�շ�����, "�շ�����"
        .Add Val(cinb�շ�����(���۱���).Text), "���۱���"
        .Add mint���ѷ�ʽ���, "�շѷ�ʽ"
        .Add CDate(cinb�շ�����(��������).Text), "��������"
        .Add um�û����, "�շ���"
        .Add "", "�˷���"
        .Add CDate("1900/1/1"), "�˷�����"
    
        '�޸ģ�2002-6-25����������վݺš�
        lstr�վݺ� = mobj�շѹ���.func�����վݺ�
        .Add lstr�վݺ�, "�վݺ�"
    End With
    
End Function


'&  ---------------| func��ȡ��ʼ������ |----------------------
'&  ��;��  ��ȡ��ʼ������
'&  ���ߣ�  Shadow
'&  ������ڣ�2001/4/3
Private Function func��ȡ��ʼ������() As Boolean
    Dim lobjҵ������ As Object
    On Error GoTo errhandle
    '��ʼ��ҵ�񼰵�λ��������
    Set mobj�շѹ��� = CreateObject("�շ�ҵ�����.cls�շѹ���")
    Set mobjҵ������ = CreateObject("�շ�ҵ�����.clsҵ������")
    Set mobj��λ���� = CreateObject("��λ����ҵ��.ClsUnitInterface")
    Set lobjҵ������ = CreateObject("�շ����ݶ���.clsҵ������")
    mint���ۿ��� = lobjҵ������.���ۿ���
    mint��Ŀ���� = lobjҵ������.��Ŀ����
    Set mrds�շ���Ŀ = mobj�շѹ���.func��ѯ�շ���Ŀ("datalength(�շ���Ŀ���)=" & CStr(3 * mint��Ŀ����))
    If mrds�շ���Ŀ Is Nothing Then
        sffuncMsg "��ȡ�շ���Ŀʧ�ܣ�����ϵͳ����Ա��ϵ��", sf����
        func��ȡ��ʼ������ = False
        Exit Function
    End If
    If mrds�շ���Ŀ.BOF Or mrds�շ���Ŀ.EOF Then
        sffuncMsg """�շ���Ŀ""��δ���ã������ú�""�շ���Ŀ""��ʹ���շѹ��ܡ�", sf����
        func��ȡ��ʼ������ = False
        Exit Function
    End If
    Set mrds�շѱ�׼ = mobj�շѹ���.funcExecute("select �շѱ�׼����,���Ƿ� from �շѹ���_�շѱ�׼��Ϣ�� group by ���Ƿ�,�շѱ�׼����", "cls������Ϣ")
    Set mrds���ܿ��� = dafuncGetData("select * from ϵͳ����_�����ֵ��")
    If mrds���ܿ��� Is Nothing Then
        sffuncMsg "�޷���ϵͳ�� ""ϵͳ����_�����ֵ��""��ȡ��ʼ�����ݡ�" & vbCrLf & "�շѹ��ܽ�����ʹ�ã�����ϵͳ����Ա��ϵ��", sf����
        func��ȡ��ʼ������ = False
        Exit Function
    End If
    If mrds���ܿ���.BOF Or mrds���ܿ���.EOF Then
        sffuncMsg """ϵͳ����_�����ֵ��"" ��δ����,�շѹ��ܽ�����ʹ�ã�" & vbCrLf & "����ϵͳ����Ա��ϵ��"
        func��ȡ��ʼ������ = False
        Exit Function
    End If
    Set mrds���ѷ�ʽ = mobjҵ������.func��ȡ�ֵ����Ϣ("�շѷ�ʽ�ֵ��")
    If mrds���ѷ�ʽ Is Nothing Then
        sffuncMsg "��ȡ�շѷ�ʽʧ��,����ϵͳ����Ա��ϵ��", sf����
        func��ȡ��ʼ������ = False
        Exit Function
    End If
    If mrds���ѷ�ʽ.BOF And mrds���ѷ�ʽ.EOF Then
        sffuncMsg "���ѷ�ʽ��δ���ã��������úý��ѷ�ʽ��ʹ���շѹ��ܣ�", sf����
        func��ȡ��ʼ������ = False
        Exit Function
    End If
    func��ȡ��ʼ������ = True
    GoTo WayOut
errhandle:
    If Err.Number = 9999 Then
        sffuncMsg "��ȱ:" & Mid$(Err.Description, InStr(Err.Description, ":") + 1) & vbCrLf & "�����úú���ʹ���շѹ��ܡ�"
    End If
    func��ȡ��ʼ������ = False
    sfsub������ ������, ģ����, "func��ȡ��ʼ������", Err.Number, Err.Description
WayOut:
    
End Function


'&  ---------------| sub��ʼ������ |----------------------
'&  ��;��  ��ʼ���շѽ���
'&  ���أ�  ��
'&  ���ߣ�  Shadow
'&  ������ڣ�2001/4/3
Private Sub sub��ʼ������()
On Error GoTo errhandle
    mstrUndoCount = ""
    mstrUndoMoney = ""
    mintCurInput = 0
    mlngX = 0
    mlngY = 0
    mstr���ѵ�λ��� = ""
    mstr���ܿ��ұ�� = ""
    mint���ѷ�ʽ��� = 0
    mstr�շ����� = ""
    mcur�ܽ�� = 0
    mcur������Ϣ�ܽ�� = 0
    Dim i As Long
    Dim j As Long
    '��ʼ�� "cing�շѻ�����Ϣ��"
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '�޸��ˣ��켽��
    '���ܣ���ʼ��cing�շѻ�����Ϣ��
    'ʱ�䣺2001-12-20
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    With cing�շѻ�����Ϣ��
        .Cols = 6
        .Rows = 1
        .TextMatrix(0, ������Ϣ_ѡ��) = "ѡ��"
        .ColWidth(������Ϣ_ѡ��) = 480
        .ColAlignment(������Ϣ_ѡ��) = flexAlignCenterCenter
       
        .TextMatrix(0, ������Ϣ_�շ�����) = "�շ�����"
        .ColWidth(������Ϣ_�շ�����) = 1440
'        .ColAlignment(������Ϣ_�շ�����) = flexAlignCenterCenter
         
       
        .TextMatrix(0, ������Ϣ_�շѱ��) = "�շѱ��"
        .ColWidth(������Ϣ_�շѱ��) = 1440
'        .ColAlignment(������Ϣ_�շѱ��) = flexAlignCenterCenter
        
        .TextMatrix(0, ������Ϣ_������) = "������"
        .ColWidth(������Ϣ_������) = 880
'        .ColAlignment(������Ϣ_������) = flexAlignCenterCenter
        
        .TextMatrix(0, ������Ϣ_���ѵ�λ����) = "���ѵ�λ����"
        .ColWidth(������Ϣ_���ѵ�λ����) = 1800
'        .ColAlignment(������Ϣ_���ѵ�λ����) = flexAlignCenterCenter
          
        .TextMatrix(0, ������Ϣ_���) = "���"
        .ColWidth(������Ϣ_���) = 1000
'        .ColAlignment(������Ϣ_���) = flexAlignCenterCenter
    End With
    
    '��ʼ�� "cing�����嵥"
    With cing�����嵥(�շ�)
        .Cols = 5
        .Rows = 1
        .TextMatrix(0, �����嵥_�շ���Ŀ���) = "�շ���Ŀ���"
        .ColWidth(�����嵥_�շ���Ŀ���) = 1310
        .ColAlignment(�����嵥_�շ���Ŀ���) = flexAlignCenterCenter
        
        .TextMatrix(0, �����嵥_�շ���Ŀ����) = "�շ���Ŀ����"
        .ColWidth(�����嵥_�շ���Ŀ����) = 1320
'        .ColAlignment(�����嵥_�շ���Ŀ����) = flexAlignCenterCenter
        
        .TextMatrix(0, �����嵥_����) = "����"
        .ColWidth(�����嵥_����) = 480
'        .ColAlignment(�����嵥_����) = flexAlignCenterCenter
        
        .TextMatrix(0, �����嵥_����) = "����"
        .ColWidth(�����嵥_����) = 500
'        .ColAlignment(�����嵥_����) = flexAlignCenterCenter
        
        
        '.TextMatrix(0, �����嵥_������λ) = "������λ"
        '.ColWidth(�����嵥_������λ) = 900
        '.ColAlignment(�����嵥_������λ) = flexAlignCenterCenter
        
        .TextMatrix(0, �����嵥_���) = "���"
        .ColWidth(�����嵥_���) = 570
'        .ColAlignment(�����嵥_���) = flexAlignCenterCenter
    End With
    '�����շ�������
    cing�շѻ�����Ϣ��.ColHidden(1) = True
    With cing�����嵥(�����շ�)
        .Cols = 5
        .Rows = 1
        .TextMatrix(0, �����嵥_�շ���Ŀ���) = "�շ���Ŀ���"
        .ColWidth(�����嵥_�շ���Ŀ���) = 1410
        .ColAlignment(�����嵥_�շ���Ŀ���) = flexAlignCenterCenter
        
        .TextMatrix(0, �����嵥_�շ���Ŀ����) = "�շ���Ŀ����"
        .ColWidth(�����嵥_�շ���Ŀ����) = 1410
'        .ColAlignment(�����嵥_�շ���Ŀ����) = flexAlignCenterCenter
        
        .TextMatrix(0, �����嵥_����) = "����"
        .ColWidth(�����嵥_����) = 480
'        .ColAlignment(�����嵥_����) = flexAlignCenterCenter
        
        .TextMatrix(0, �����嵥_����) = "����"
        .ColWidth(�����嵥_����) = 480
'        .ColAlignment(�����嵥_����) = flexAlignCenterCenter
        
        
        '.TextMatrix(0, �����嵥_������λ) = "������λ"
        '.ColWidth(�����嵥_������λ) = 900
        '.ColAlignment(�����嵥_������λ) = flexAlignCenterCenter
        
        .TextMatrix(0, �����嵥_���) = "���"
        .ColWidth(�����嵥_���) = 570
'        .ColAlignment(�����嵥_���) = flexAlignCenterCenter
    End With
    
    '��ʼ�� "cind�ֵ�(�ֵ�_�շ���Ŀ)"
        
    With cind�ֵ�(�ֵ�_�շ���Ŀ)
        .Rows = 1
        .Cols = 1
    End With
    
'    '��cind�ֵ�(�ֵ�_�շ���Ŀ)
'    If Not (mrds�շ���Ŀ Is Nothing) Then
'        cind�ֵ�(�ֵ�_�շ���Ŀ).Cols = mrds�շ���Ŀ.Fields.Count
'        cind�ֵ�(�ֵ�_�շ���Ŀ).Refresh
'        For i = 0 To mrds�շ���Ŀ.Fields.Count - 1
'            cind�ֵ�(�ֵ�_�շ���Ŀ).TextMatrix(0, i) = mrds�շ���Ŀ(i).Name
'        Next
'
'        If (Not mrds�շ���Ŀ.BOF) And (Not mrds�շ���Ŀ.EOF) Then
'            cind�ֵ�(�ֵ�_�շ���Ŀ).Rows = mrds�շ���Ŀ.RecordCount + 1
'            mrds�շ���Ŀ.MoveFirst
'            For i = 0 To mrds�շ���Ŀ.RecordCount - 1
'                For j = 0 To mrds�շ���Ŀ.Fields.Count - 1
'                    '�����ܣ��������ݿ���ܴ��ڵĿ�ֵ,����ת��Ϊ""��ʱ�䣺2002/01/27,�켽����
'                    cind�ֵ�(�ֵ�_�շ���Ŀ).TextMatrix(i + 1, j) = IIf(IsNull(mrds�շ���Ŀ(j)), "", mrds�շ���Ŀ(j))
'                    cind�ֵ�(�ֵ�_�շ���Ŀ).AutoSize j
'                Next j
'                If Not mrds�շ���Ŀ.EOF Then mrds�շ���Ŀ.MoveNext
'            Next i
'        End If
'    End If
'    cind�ֵ�(�ֵ�_�շ���Ŀ).Width = 165
'    For i = 0 To cind�ֵ�(�ֵ�_�շ���Ŀ).Cols - 1
'        cind�ֵ�(�ֵ�_�շ���Ŀ).Width = cind�ֵ�(�ֵ�_�շ���Ŀ).Width + cind�ֵ�(�ֵ�_�շ���Ŀ).ColWidth(i)
'    Next
    
    Dim llngStyle As Long
    
    '��ʼ�� "cind�ֵ�(�ֵ�_�շѱ�׼)"
    If Not (mrds�շѱ�׼ Is Nothing) Then
        cind�ֵ�(�ֵ�_�շѱ�׼).Cols = mrds�շѱ�׼.Fields.Count + 1
        For i = 1 To mrds�շѱ�׼.Fields.Count
            cind�ֵ�(�ֵ�_�շѱ�׼).TextMatrix(0, i) = mrds�շѱ�׼(i - 1).Name
        Next
        cind�ֵ�(�ֵ�_�շѱ�׼).TextMatrix(0, 0) = "���"
        
        If Not (mrds�շѱ�׼.BOF And mrds�շѱ�׼.EOF) Then
            cind�ֵ�(�ֵ�_�շѱ�׼).Rows = mrds�շѱ�׼.RecordCount + 1
            mrds�շѱ�׼.MoveFirst
            For i = 1 To mrds�շѱ�׼.RecordCount
                cind�ֵ�(�ֵ�_�շѱ�׼).TextMatrix(i, 0) = i
                For j = 1 To mrds�շѱ�׼.Fields.Count
                    '�����ܣ��������ݿ���ܴ��ڵĿ�ֵ,����ת��Ϊ""��ʱ�䣺2002/01/27,�켽����
                    cind�ֵ�(�ֵ�_�շѱ�׼).TextMatrix(i, j) = IIf(IsNull(mrds�շѱ�׼(j - 1)), "", mrds�շѱ�׼(j - 1))
                    cind�ֵ�(�ֵ�_�շѱ�׼).AutoSize j
                Next j
                If Not mrds�շѱ�׼.EOF Then mrds�շѱ�׼.MoveNext
            Next i
        End If
        
        cind�ֵ�(�ֵ�_�շѱ�׼).Width = 165
        For i = 0 To cind�ֵ�(�ֵ�_�շѱ�׼).Cols - 1
            cind�ֵ�(�ֵ�_�շѱ�׼).Width = cind�ֵ�(�ֵ�_�շѱ�׼).Width + cind�ֵ�(�ֵ�_�շѱ�׼).ColWidth(i)
        Next
    End If
    
    '��ʼ�� "cind�ֵ�_���ܿ���"
    If Not (mrds���ܿ��� Is Nothing) Then
        cind�ֵ�(�ֵ�_���ܿ���).Cols = mrds���ܿ���.Fields.Count
        For i = 0 To mrds���ܿ���.Fields.Count - 1
            cind�ֵ�(�ֵ�_���ܿ���).TextMatrix(0, i) = mrds���ܿ���(i).Name
        Next

        If (Not mrds���ܿ���.BOF) And (Not mrds���ܿ���.EOF) Then
            cind�ֵ�(�ֵ�_���ܿ���).Rows = mrds���ܿ���.RecordCount + 1
            mrds���ܿ���.MoveFirst
            For i = 0 To mrds���ܿ���.RecordCount - 1
                For j = 0 To mrds���ܿ���.Fields.Count - 1
                    '�����ܣ��������ݿ���ܴ��ڵĿ�ֵ,����ת��Ϊ""��ʱ�䣺2002/01/27,�켽����
                    cind�ֵ�(�ֵ�_���ܿ���).TextMatrix(i + 1, j) = IIf(IsNull(mrds���ܿ���(j)), "", mrds���ܿ���(j))
                    cind�ֵ�(�ֵ�_���ܿ���).AutoSize j
                Next j
                If Not mrds���ܿ���.EOF Then mrds���ܿ���.MoveNext
            Next i
        End If
    End If
    
    llngStyle = GetWindowLong(cind�ֵ�(�ֵ�_�շ���Ŀ).hwnd, GWL_STYLE)
    SetWindowLong cind�ֵ�(�ֵ�_�շ���Ŀ).hwnd, GWL_STYLE, llngStyle Or WS_THICKFRAME
    SetWindowLong cind�ֵ�(�ֵ�_�շѱ�׼).hwnd, GWL_STYLE, llngStyle Or WS_THICKFRAME
    SetWindowLong cind�ֵ�(�ֵ�_���ܿ���).hwnd, GWL_STYLE, llngStyle Or WS_THICKFRAME
    
        '�޸�(�켽��):
    '����:��Ȩ�޿��ƽ����ϴ��ۿؼ�������
    'ʱ��:2001-12-19
    If umfuncУ���û�Ȩ��("�շѹ���_����") Then
        Frame6.Enabled = True
        Frame6.Caption = "�������"
        lblCaption(19).Enabled = True
        cinb�շ�����(19).Enabled = True
        cupd�޸Ĵ��۱���.Enabled = True
        cchk��ӡ���۱���.Enabled = True
        Select Case mint���ۿ���
            Case 0
                cupd�޸Ĵ��۱���.Enabled = False
                cinb�շ�����(19).Enabled = False
            Case 1
                cupd�޸Ĵ��۱���.Enabled = True
                cinb�շ�����(19).Enabled = True
            Case 2
                cupd�޸Ĵ��۱���.Enabled = False
                cinb�շ�����(19).Enabled = False
            Case Else
        End Select
        
    Else
        Frame6.Caption = "�������(��Ȩ��)"
        Frame6.Enabled = False
        lblCaption(19).Enabled = False
        cinb�շ�����(19).Enabled = False
        cupd�޸Ĵ��۱���.Enabled = False
        cchk��ӡ���۱���.Enabled = False
    End If
        
    '��ʼ�����ѷ�ʽ�б�
    If Not (mrds���ѷ�ʽ Is Nothing) Then
        If Not (mrds���ѷ�ʽ.BOF And mrds���ѷ�ʽ.EOF) Then
            mrds���ѷ�ʽ.MoveFirst
            For i = 0 To mrds���ѷ�ʽ.RecordCount - 1
                cmb���ѷ�ʽ.AddItem mrds���ѷ�ʽ("����")
                If Not mrds���ѷ�ʽ.EOF Then mrds���ѷ�ʽ.MoveNext
            Next
        End If
        cmb���ѷ�ʽ.ListIndex = 0
    End If
    cinb�շ�����(��������).Text = Date
    cinb�շ�����(�շ�_���ܿ���).Text = um�û���������
    cinb�շ�����(�����շ�_���ܿ���).Text = um�û���������
    
    Exit Sub
errhandle:
    sfsub������ "�շѹ���������", "frm�շ�", "sub��ʼ������", Err.Number, Err.Description
End Sub

'&  ---------------| subDisable�շ����� |----------------------
'&  ��;��  ʹ�շѽ����ϵ������ͱ��ʧЧ
'&  ���ߣ�  Shadow
'&  ������ڣ�2001/4/3
Private Sub subDisable�շ�����()
On Error Resume Next
    Dim i As Long
    For i = �շ�_�շѱ�� To �շ�_�շѱ�׼
        cinb�շ�����(i).Enabled = False
    Next
    cing�շѻ�����Ϣ��.Enabled = False
    cing�����嵥(�շ�).Enabled = False
End Sub

'&  ---------------| subEnable�շ����� |----------------------
'&  ��;��  ʹ�շѽ����ϵ������ͱ����Ч
'&  ���أ�  ��
'&  ���ߣ�  Shadow
'&  ������ڣ�2001/4/3

Private Sub subEnable�շ�����()
    Dim i As Long
    
    On Error Resume Next
    If cchk�ڲ��շ�.Value = 1 Then
        For i = �շ�_�շѱ�� To �շ�_�շѱ�׼
            cinb�շ�����(i).Enabled = True
        Next
        For i = �շ�_�շ���Ŀ To �շ�_�շѱ�׼
            cinb�շ�����(i).Enabled = False
        Next
    Else
        For i = �շ�_�շ���Ŀ To �շ�_�շѱ�׼
            cinb�շ�����(i).Enabled = True
        Next
        cinb�շ�����(�շ�_�շѱ��).Enabled = False
    End If
    
    If cchkͬ���շ�.Value = 1 Then cing�շѻ�����Ϣ��.Editable = True
    cing�����嵥(�շ�).Enabled = True
    Err.Clear
End Sub

'&  ---------------| subDisable�����շ����� |----------------------
'&  ��;��  ʹ�����շѽ����ϵ������ͱ��ʧЧ
'&  ���ߣ�  Shadow
'&  ������ڣ�2001/4/3
Private Sub subDisable�����շ�����()
On Error Resume Next
    Dim i As Long
    For i = �����շ�_������ To �����շ�_�շѱ�׼
        cinb�շ�����(i).Enabled = False
    Next
    For i = ��Ժ To ��Ժ
        cdtp����(i).Enabled = False
    Next
    cing�����嵥(�����շ�).Enabled = False
End Sub

'&  ---------------| subEnable�����շ����� |----------------------
'&  ��;��  ʹ�����շѽ����ϵ������ͱ����Ч
'&  ���ߣ�  Shadow
'&  ������ڣ�2001/4/3
Private Sub subEnable�����շ�����()
On Error Resume Next
    Dim i As Long
    For i = �����շ�_������ To �����շ�_�շѱ�׼
        cinb�շ�����(i).Enabled = True
    Next
    For i = ��Ժ To ��Ժ
        cdtp����(i).Enabled = True
    Next
    cing�����嵥(�����շ�).Enabled = True
End Sub

Private Sub mobj����ͨ�ö���_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    Dim lrds������Ϣ As Recordset
    Dim lcol���� As Collection
    Dim lcol�շ�ȷ����Ϣ As Collection
    Dim lstr���۷��� As String
    Dim lstr�շѱ�� As String
    Dim i As Long, j As Long
    Dim lstr�շѱ����() As String
    Dim lint�շѱ������ As Integer
    
    On Error GoTo errhandle
    Select Case Operate

'&  =============================================| �շ� |============================================
        Case "�շ�"
            If Not ValidateData Then Exit Sub
            
            ctlb������.Buttons(9).Enabled = False
            
            mrds���ѷ�ʽ.Filter = "����='" & cmb���ѷ�ʽ.Text & "'"
            mint���ѷ�ʽ��� = mrds���ѷ�ʽ("���").Value
            
            If cing�����嵥(ctabShoufei.Tab).Rows = 1 Then
                sffuncMsg "�޿��÷�����Ϣ", sf����
                GoTo WayOut
            End If
            
            If IsNumeric(cinb�շ�����(19).Text) Then
                If CDbl(cinb�շ�����(19).Text) = 0 Then
                    MsgBox "���۱��ʲ���Ϊ0��", vbOKOnly & vbExclamation, "ϵͳ��ʾ"
                    Cancel = True
                    cinb�շ�����(19).Text = "1.00"
                    Exit Sub
                End If
            Else
                MsgBox "���۱���¼�벻��ȷ��", vbOKOnly & vbExclamation, "ϵͳ��ʾ"
                Cancel = True
                Exit Sub
            End If
            
            If cchk�ڲ��շ�.Value = 0 Or ctabShoufei.Tab = �����շ� Then
                
                'ֱ���շ�
                Set lcol���� = func�ռ�����
                If lcol���� Is Nothing Then
                    sffuncMsg "�ռ�����ʧ�ܣ�", sf����
                    GoTo WayOut
                End If
                lstr���۷��� = mobj�շѹ���.func����_���ݼ���(lcol����)
                
                '�޸ģ����϶Բ��˳���Ժ���ڵ��жϣ�ʱ�䣺2002/02/05���켽��
                If cdtp����(0).Value > cdtp����(1).Value Then
                    sffuncMsg "��Ժ���ڲ���С����Ժ����,����������ʱ�䣡"
                    GoTo WayOut
                End If
                
                If lstr���۷��� = "" Then
                    sffuncMsg "ִ�л��۲���ʱʧ�ܣ�", sf����
                    GoTo WayOut
                Else
                    lstr�շѱ�� = Mid(lstr���۷���, InStr(lstr���۷���, ";") + 1)
                    mstr�շ����� = lstr�շѱ��
                End If
                
                '�����շ���Ϣ��
                '�޸ģ�2002-6-25����Ϊ�վ�ȷ����Ϣ��Ҫ�����վݺţ����������������
                On Error GoTo errTransHandler
                dasubBeginTran
                
                Set lcol�շ�ȷ����Ϣ = func�ռ�ȷ����Ϣ '�޸ģ�2002-6--25�����������վݺš�
                
                Call mobj�շѹ���.func�շ�(mstr�շ�����, lcol�շ�ȷ����Ϣ)
                
                dasubCommitTran
                On Error GoTo errhandle
                
                ctlb������.Buttons("�շ�(&G)").Enabled = False
                
                MsgBox "�շ���ɣ���ȴ���ӡƱ�ݣ�", vbInformation, "ϵͳ��ʾ"
                mcur�ܽ�� = 0
                cinb�շ�����(Ӧ�ս��) = "0"
                If Not (cchk�ڲ��շ�.Value = 1 And ctabShoufei.Tab = �շ�) Then sub�������
                If ctabShoufei.Tab = �շ� Then
                    If cinb�շ�����(�շ�_������).Enabled Then
                        cinb�շ�����(�շ�_������).SetFocus
                    Else
                        cinb�շ�����(�շ�_�շѱ��).SetFocus
                    End If
                Else
                    If cinb�շ�����(�����շ�_������).Enabled Then cinb�շ�����(�����շ�_������).SetFocus
                End If
            
            ElseIf (cchk�ڲ��շ�.Value = 1) And (ctabShoufei.Tab = �շ�) Then
                
                '�ڲ��շ�
                For i = 1 To cing�շѻ�����Ϣ��.Rows - 1
                    If cing�շѻ�����Ϣ��.Cell(flexcpChecked, i, 0) = 1 Then
                        lint�շѱ������ = lint�շѱ������ + 1
                        ReDim Preserve lstr�շѱ����(lint�շѱ������)
                        lstr�շѱ����(lint�շѱ������) = cing�շѻ�����Ϣ��.TextMatrix(i, 2)
                    End If
                Next
                
                If lint�շѱ������ = 0 Then
                    sffuncMsg "��ѡ�е��շ���Ϣ��", sf����
                    GoTo WayOut
                End If
                                                
                mstr�շ����� = lstr�շѱ����(1)
                
                '�����շ���Ϣ��
                '�޸ģ�2002-6-25����Ϊ�վ�ȷ����Ϣ��Ҫ�����վݺţ����������������
                On Error GoTo errTransHandler
                dasubBeginTran
                
                Set lcol�շ�ȷ����Ϣ = func�ռ�ȷ����Ϣ '�޸ģ�2002-6--25�����������վݺš�
                
                For i = 1 To UBound(lstr�շѱ����)
                    Call mobj�շѹ���.func�շ�(lstr�շѱ����(i), lcol�շ�ȷ����Ϣ)
                Next
                
                dasubCommitTran
                On Error GoTo errhandle
                
                'ReDim lstr�շѱ����(0)
                For i = cing�շѻ�����Ϣ��.Rows - 1 To 1 Step -1
                    If i <= cing�շѻ�����Ϣ��.Rows - 1 Then
                        If cing�շѻ�����Ϣ��.Cell(flexcpChecked, i, 0) = 1 Then
                            cing�շѻ�����Ϣ��.RemoveItem i
                        End If
                    End If
                Next
                cing�����嵥(ctabShoufei.Tab).Rows = 1
                ctlb������.Buttons("�շ�(&G)").Enabled = False
                MsgBox "�շ���ɣ���ȴ���ӡƱ�ݣ�", vbInformation, "ϵͳ��ʾ"
                mcur������Ϣ�ܽ�� = 0
                cinb�շ�����(Ӧ�ս��) = "0"
                cinb�շ�����(ʵ�ս��).Text = "0"
                If Not (cchk�ڲ��շ�.Value = 1 And ctabShoufei.Tab = �շ�) Then sub�������
                If cinb�շ�����(�շ�_�շѱ��).Enabled Then cinb�շ�����(�շ�_�շѱ��).SetFocus
            End If
'&  =====================================================| ��ӡ |=================================================
                Dim lrdsreturn As Recordset
                Dim llngFieldCounter As Long
                Dim llngRecordCounter As Long
                Dim lstr��ʽ�ļ��� As String
                Dim lcol�����¼ As Collection
                Dim lcol������� As Collection
                Dim lrds���ظ�ʽ�ļ��� As Recordset
                Dim lobj���ܼ�¼ As Object
                
                'ȡ���շѼ�¼���е�Ʊ����������
                Set lrdsreturn = mobj�շѹ���.funcExecute("select b.Ʊ�����ͱ�� from �շѹ���_�շ���Ŀ�ֵ�� b, �շѹ���_������Ϣ�� c " & _
                                                                "Where b.�շ���Ŀ��� = c.�շ���Ŀ��� and c.�շ����� ='" & _
                                                                mstr�շ����� & "' group by b.Ʊ�����ͱ��", "cls������Ϣ")
                
                If lrdsreturn Is Nothing Then
                    sffuncMsg "δ�������շ���Ŀ��Ʊ��������Ϣ,�޷����д�ӡ��", sf����
                    If Not (cchk�ڲ��շ�.Value = 1 And ctabShoufei.Tab = �շ�) Then sub�������
                    GoTo WayOut
                End If

                If (lrdsreturn.BOF And lrdsreturn.EOF) Then
                    sffuncMsg "δ�������շ���Ŀ��Ʊ��������Ϣ,�޷����д�ӡ��", sf����
                    If Not (cchk�ڲ��շ�.Value = 1 And ctabShoufei.Tab = �շ�) Then sub�������
                Else
                    lrdsreturn.MoveFirst
                End If
                
                '��Ʊ������ȡ��������Ϣ
                For i = 0 To lrdsreturn.RecordCount - 1
                                                            
                    Set lrds������Ϣ = mobj�շѹ���.funcExecute("select * from �շѹ���_��ӡ������Ϣ where Ʊ�����ͱ��=" & lrdsreturn("Ʊ�����ͱ��") & " and �շ�����='" & mstr�շ����� & "'", "cls������Ϣ")

                    Set lcol������� = New Collection
                    If lrds������Ϣ Is Nothing Then
                        sffuncMsg "�޿ɴ�ӡ��Ϣ��", sf����
                        If Not (cchk�ڲ��շ�.Value = 1 And ctabShoufei.Tab = �շ�) Then sub�������
                        GoTo WayOut
                    End If
                    
                    If lrds������Ϣ.BOF And lrds������Ϣ.EOF Then
                        sffuncMsg "�޿ɴ�ӡ��Ϣ��", sf����
                        If Not (cchk�ڲ��շ�.Value = 1 And ctabShoufei.Tab = �շ�) Then sub�������
                        GoTo WayOut
                    End If
                    
                     '*******����Ϊ�μ�����-0823**********
                    Dim lstr���ѵ�λ As String
                    Dim lstr������ As String
                    If IIf(IsNull(lrds������Ϣ("���ѵ�λ����").Value), "", lrds������Ϣ("���ѵ�λ����")) <> "" Then
                        lstr���ѵ�λ = lrds������Ϣ("���ѵ�λ����").Value
                    Else
                        lstr���ѵ�λ = ""
                    End If
                    If IIf(IsNull(lrds������Ϣ("������").Value), "", lrds������Ϣ("������")) <> "" Then
                        lstr������ = lrds������Ϣ("������").Value
                    Else
                        lstr������ = ""
                    End If
                    '********����Ϊ�μ�����-0823**********
                    
                    
                    '�޸ģ�2002-9-29������ϲ���ӡ��
                    Set lobj���ܼ�¼ = mobj�շѹ���.funcExecute("select �շ���Ŀ���,����=avg(����),����=sum(����),���=sum(���) from �շѹ���_��ӡ������Ϣ " _
                                & "where Ʊ�����ͱ��=" & lrdsreturn("Ʊ�����ͱ��") & " and �շ�����='" & mstr�շ����� _
                                & "' group by �շ�����,�շ���Ŀ���", "cls������Ϣ")
                    
                    For llngRecordCounter = 0 To lobj���ܼ�¼.RecordCount - 1
                        
                        '�޸ģ�2002-9-29�������ȡ��ǰ��Ŀ����ϸ��Ϣ��
                        Set lrds������Ϣ = mobj�շѹ���.funcExecute("select * from �շѹ���_��ӡ������Ϣ where Ʊ�����ͱ��=" & lrdsreturn("Ʊ�����ͱ��") & " and �շ�����='" & mstr�շ����� & "' AND �շ���Ŀ���='" & lobj���ܼ�¼("�շ���Ŀ���") & "'", "cls������Ϣ")
                        
                        Set lcol�����¼ = New Collection
                        
                        '��һ����¼��ӵ�һ������(lcol�����¼)
                        For llngFieldCounter = 0 To lrds������Ϣ.Fields.Count - 1
                            If lrds������Ϣ.Fields(llngFieldCounter).Name = "���ѵ�λ����" Or lrds������Ϣ.Fields(llngFieldCounter).Name = "������" Then
                                If lrds������Ϣ.Fields(llngFieldCounter).Name = "���ѵ�λ����" Then lcol�����¼.Add lstr���ѵ�λ, "���ѵ�λ����"
                                If lrds������Ϣ.Fields(llngFieldCounter).Name = "������" Then lcol�����¼.Add lstr������, "������"
                            ElseIf lrds������Ϣ.Fields(llngFieldCounter).Name <> "����" And lrds������Ϣ.Fields(llngFieldCounter).Name <> "����" And lrds������Ϣ.Fields(llngFieldCounter).Name <> "���" Then
                                '�޸ģ�2002-9-29��������ۡ������������ʾ�������ݡ�
                                lcol�����¼.Add lrds������Ϣ(llngFieldCounter).Value, lrds������Ϣ.Fields(llngFieldCounter).Name
                            End If
                        Next
                        '�޸ģ�2002-9-29��������ۡ������������ʾ�������ݡ�
                        lcol�����¼.Add Format(lobj���ܼ�¼("����").Value, "0.00"), "����"
                        lcol�����¼.Add lobj���ܼ�¼("����").Value, "����"
                        lcol�����¼.Add Format(lobj���ܼ�¼("���").Value, "0.00"), "���"
                        
'                        '��������շ���Ϣ
                        lcol�����¼.Add cinb�շ�����(�����շ�_����).Text, "����"
                        lcol�����¼.Add cinb�շ�����(�����շ�_�Ա�).Text, "�Ա�"
                        lcol�����¼.Add cinb�շ�����(�����շ�_סԺ��).Text, "סԺ��"
                        lcol�����¼.Add cinb�շ�����(�����շ�_����).Text, "����"
                        lcol�����¼.Add CStr(cdtp����(��Ժ).Year) & CStr(cdtp����(��Ժ).Month) & CStr(cdtp����(��Ժ).Day), "��Ժ����"
                        lcol�����¼.Add CStr(cdtp����(��Ժ).Year) & CStr(cdtp����(��Ժ).Month) & CStr(cdtp����(��Ժ).Day), "��Ժ����"
                        lcol�����¼.Add cinb�շ�����(�����շ�_��Ժ����Ա).Text, "��Ժ����Ա"
                        lcol�����¼.Add cinb�շ�����(�����շ�_����ҽ��).Text, "����ҽ��"
                        
                        '�����ü�¼���뼯��lcol�������
                        lcol�������.Add lcol�����¼
                        
                        '����¼������
'                        If Not lrds������Ϣ.EOF Then lrds������Ϣ.MoveNext
                        If Not lobj���ܼ�¼.EOF Then lobj���ܼ�¼.MoveNext
                        
                    Next
                    
                    '����Ʊ�ݸ�ʽ��
                    If ctabShoufei.Tab = �շ� Then
                        Set lrds���ظ�ʽ�ļ��� = mobj�շѹ���.funcExecute("select * from �շѹ���_Ʊ��������Ϣ�� where Ʊ�����ͱ��='" & lrdsreturn("Ʊ�����ͱ��") & "' and ��Ӧҵ��='һ��'", "cls������Ϣ")
                    Else
                        Set lrds���ظ�ʽ�ļ��� = mobj�շѹ���.funcExecute("select * from �շѹ���_Ʊ��������Ϣ�� where Ʊ�����ͱ��='" & lrdsreturn("Ʊ�����ͱ��") & "' and ��Ӧҵ��='����'", "cls������Ϣ")
                    End If
                    If lrds���ظ�ʽ�ļ��� Is Nothing Then
                        sffuncMsg "δ���ҵ�Ʊ�ݸ�ʽ�ļ���", sf����
                    End If
                        
                    '��ʼ��ӡƱ�ݡ�
                    If lrds���ظ�ʽ�ļ���.BOF And lrds���ظ�ʽ�ļ���.EOF Then
                        sffuncMsg "δ���ҵ�Ʊ�ݸ�ʽ�ļ���", sf����
                    Else
                        lstr��ʽ�ļ��� = lrds���ظ�ʽ�ļ���("Ʊ�ݸ�ʽ�ļ�����")
                        Call mobj�շѹ���.sub��ӡƱ��(lcol�������, App.Path & "\" & lstr��ʽ�ļ���, IIf(cchkԤ��.Value = 1, True, False), cchk��ӡ���۱���.Value, lrds���ظ�ʽ�ļ���("�������").Value)
                    End If
                    If Not lrdsreturn.EOF Then lrdsreturn.MoveNext
                Next
                
                If Not (cchk�ڲ��շ�.Value = 1 And ctabShoufei.Tab = �շ�) Then sub�������
                ctlb������.Buttons("�շ�(&G)").Enabled = True
                
                ' �ָ��˳�����(�켽��2002_1_10)
                ctlb������.Buttons(9).Enabled = True
                Set lcol�����¼ = Nothing
                Set lcol������� = Nothing
                           
'&  ======================================| ��ѯ |============================================
        Case "��ѯ"

            '�޸ģ�2001-11-22���������ѵ�λ�����ܿ��ң���ѯ��
            Dim lobjRec As Object
            Dim lstr�����շѱ�� As String
           ' Dim lstr������ As String
            Dim lstr��λ���� As String
            Dim lstr���ܿ��ұ�� As String
            Dim lstrSql As String
            Dim lstrʱ����� As String
            Dim lstr���� As String
            Dim lstr���� As String
            Dim lstr���� As String
            Dim lTime As String
            Dim lstrҵ����� As String          '���������¼ҵ�����
            
            '��������Ŀ����
            Sub�����շ���Ŀ����
            
            '����ڲ��շ������е�����
            cing�շѻ�����Ϣ��.Rows = 1
            '��ѯ�����¼���շѱ�š����ܿ��ҡ����ѵ�λ��š�
            lstr�����շѱ�� = cinb�շ�����(�շ�_�շѱ��).Text
            If cinb�շ�����(�շ�_���ѵ�λ) = "" Then
                lstr��λ���� = ""
            Else
                lstr��λ���� = Trim(cinb�շ�����(�շ�_���ѵ�λ))
            End If
            
            If cinb�շ�����(�շ�_������) = "" Then
                lstr������ = ""
            Else
                lstr������ = Trim(cinb�շ�����(�շ�_������))
            End If
            
            If cinb�շ�����(�շ�_���ܿ���) = "" Then
                lstr���ܿ��ұ�� = ""
            Else
                lstr���ܿ��ұ�� = mstr���ܿ��ұ��
            End If
            If lstr�����շѱ�� = "" Then
                cing�շѻ�����Ϣ��.Rows = 1
            End If
            
            lstrҵ����� = Ccboҵ�����.Text
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '�޸��ˣ��켽��
            '���ܣ����������ɲ����
            'ʱ�䣺2001/12/20
            '
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            If Cchk��������Ϣ��ѯ.Value = Unchecked Then
                If Copt����.Value = True Then lstrʱ����� = "����"
                If Copt����.Value = True Then lstrʱ����� = "����"
                If Copt����.Value = True Then lstrʱ����� = "����"
                
                '���Ӷ�ҵ�����Ĳ�ѯ
                
                If lstrҵ����� = "����ҵ��" Then
                    lstrSql = "select a.���۱���,a.�շ�����,a.�շѱ��,a.�շ���Ŀ���,b.�շ���Ŀ����,a.����,a.����,b.������λ,a.���,a.���ѵ�λ����,a.������,d.Ƭ��,c.���� as ���ܿ�������" & _
                        " from �շѹ���_������Ϣ�� a left join �շѹ���_�շ���Ŀ�ֵ�� b on a.�շ���Ŀ��� = b.�շ���Ŀ��� " & _
                        " left join ϵͳ����_�����ֵ�� c on a.���ܿ��ұ��=c.��� " & _
                        " left join ��λ����_��λ������Ϣ�� d  on a.���ѵ�λ���=d.������" & _
                        " where a.�շ�״̬= 0 "
                Else
                    lstrSql = "select a.���۱���,a.�շ�����,a.�շѱ��,a.�շ���Ŀ���,b.�շ���Ŀ����,a.����,a.����,b.������λ,a.���,a.���ѵ�λ����,a.������,d.Ƭ��, c.���� as ���ܿ�������" & _
                        " from �շѹ���_������Ϣ�� a left join �շѹ���_�շ���Ŀ�ֵ�� b on a.�շ���Ŀ��� = b.�շ���Ŀ��� " & _
                        " left join ϵͳ����_�����ֵ�� c on a.���ܿ��ұ��=c.��� " & _
                        " left join ��λ����_��λ������Ϣ�� d  on a.���ѵ�λ���=d.������" & _
                        " where a.�շ�״̬= 0  and a.ҵ�����= '" & lstrҵ����� & "'"
                End If

                Select Case lstrʱ�����
                    Case "����" '��ʾ������շѼ�¼��
                        lTime = Format(Now(), "yyyy-mm-dd")
                        lstrSql = lstrSql & " and left(convert(varchar(30),a.��������,120),10)='" & lTime & "'"
                    Case "����" '��ʾ���µ��շѼ�¼��
                        lTime = Format(Now(), "yyyy-mm")
                        lstrSql = lstrSql & " and left(convert(varchar(30),a.��������,120),7)='" & lTime & "'"
                    Case "����" '��ʾ���е��շѼ�¼��
                        lstrSql = lstrSql
                End Select
                
            
            Else
            
                If lstrҵ����� = "����ҵ��" Then
                    lstrSql = "select a.���۱���,a.�շ�����,a.�շѱ��,a.�շ���Ŀ���,b.�շ���Ŀ����,a.����,a.����,b.������λ,a.���,a.���ѵ�λ����,a.������,d.Ƭ��,���ܿ������� = c.����" & _
                        " from �շѹ���_������Ϣ�� a left join �շѹ���_�շ���Ŀ�ֵ�� b on a.�շ���Ŀ��� = b.�շ���Ŀ��� " & _
                        " left join ϵͳ����_�����ֵ�� c on a.���ܿ��ұ��=c.��� " & _
                        " left join ��λ����_��λ������Ϣ�� d on a.���ѵ�λ���=d.������" & _
                        " where a.�շ�״̬= 0 and �շѱ��=" & IIf(lstr�����շѱ�� = "", "�շѱ��", "'" & lstr�����շѱ�� & "'") & _
                        " and a.���ܿ��ұ��=" & IIf(lstr���ܿ��ұ�� = "" Or lstr�����շѱ�� <> "", "a.���ܿ��ұ��", "'" & lstr���ܿ��ұ�� & "'") & _
                        " and a.���ѵ�λ����=" & IIf(lstr��λ���� = "" Or lstr�����շѱ�� <> "", "a.���ѵ�λ����", "'" & lstr��λ���� & "'") & _
                        " and a.������= " & IIf(lstr������ = "" Or lstr�����շѱ�� <> "", "a.������", "'" & lstr������ & "'")
                Else
                    lstrSql = "select a.���۱���,a.�շ�����,a.�շѱ��,a.�շ���Ŀ���,b.�շ���Ŀ����,a.����,a.����,b.������λ,a.���,a.���ѵ�λ����,a.������,d.Ƭ��,���ܿ������� = c.����" & _
                        " from �շѹ���_������Ϣ�� a left join �շѹ���_�շ���Ŀ�ֵ�� b on a.�շ���Ŀ��� = b.�շ���Ŀ��� " & _
                        " left join ϵͳ����_�����ֵ�� c on a.���ܿ��ұ��=c.��� " & _
                        " left join ��λ����_��λ������Ϣ�� d on a.���ѵ�λ���=d.������" & _
                        " where a.ҵ�����= '" & lstrҵ����� & "'" & _
                        " and a.�շ�״̬= 0 and �շѱ��=" & IIf(lstr�����շѱ�� = "", "�շѱ��", "'" & lstr�����շѱ�� & "'") & _
                        " and a.���ܿ��ұ��=" & IIf(lstr���ܿ��ұ�� = "" Or lstr�����շѱ�� <> "", "a.���ܿ��ұ��", "'" & lstr���ܿ��ұ�� & "'") & _
                        " and a.���ѵ�λ����=" & IIf(lstr��λ���� = "" Or lstr�����շѱ�� <> "", "a.���ѵ�λ����", "'" & lstr��λ���� & "'") & _
                        " and a.������= " & IIf(lstr������ = "" Or lstr�����շѱ�� <> "", "a.������", "'" & lstr������ & "'")
                End If
            End If
            
            '�����������
            lstrSql = lstrSql + " order by  a.�շѱ�� desc"
            Set lrds������Ϣ = mobj�շѹ���.funcExecute(lstrSql, "cls������Ϣ")
                        
            If (lrds������Ϣ Is Nothing) Then
                'cing�����嵥(0).Clear
                
                Clab�շ���Ŀ����.Enabled = False
                Ccbo�շ���Ŀ����.Enabled = False
                lblCaption(5).Enabled = False
                cinb�շ�����(5).Enabled = False
            
                cing�����嵥(0).Rows = 1
                sffuncMsg "�޷��������ķ�����Ϣ��", sf����
                GoTo WayOut
            ElseIf (lrds������Ϣ.BOF And lrds������Ϣ.EOF) Then
                'cing�����嵥(0).Clear
                Clab�շ���Ŀ����.Enabled = False
                Ccbo�շ���Ŀ����.Enabled = False
                lblCaption(5).Enabled = False
                cinb�շ�����(5).Enabled = False
                cing�����嵥(0).Rows = 1
                sffuncMsg "�޷��������ķ�����Ϣ��", sf����
                GoTo WayOut
            Else
                lrds������Ϣ.MoveFirst
                If cchk�ڲ��շ�.Value = 0 Then
                    cing�շѻ�����Ϣ��.Rows = 1
                End If
                cing�����嵥(�շ�).Rows = 1
                If lrds������Ϣ("���ѵ�λ����").Value <> vbNullString Then
                    cinb�շ�����(�շ�_���ѵ�λ) = lrds������Ϣ("���ѵ�λ����")
                End If
                
                cinb�շ�����(�շ�_������) = lrds������Ϣ("������")
                cinb�շ�����(�շ�_���ܿ���) = lrds������Ϣ("���ܿ�������")
                cinb�շ�����(���۱���).Text = lrds������Ϣ("���۱���")
                
                
                '��ʾƬ����Ϣ
                If IIf(IsNull(lrds������Ϣ("Ƭ��")), "", lrds������Ϣ("Ƭ��")) = "" Then
                    ClabƬ��.Caption = "Ƭ����(����)"
                Else
                    ClabƬ��.Caption = "(" + lrds������Ϣ("Ƭ��") + ")"
                End If
                
                '�޸ģ�2001/12/20���켽�������������շѱ�ŵ���ʾ
                cinb�շ�����(�շ�_�շѱ��).Text = lrds������Ϣ("�շѱ��")
                cinb�շ�����(�շ�_���ܿ���).Text = lrds������Ϣ("���ܿ�������")
                
                '�����������Ŀ
                For i = 0 To lrds������Ϣ.RecordCount - 1
                    'lrds������Ϣ("������λ") & vbTab &
                    '���շѱ����ͬ,����"cing�����嵥"�����Ŀ,���ۼ�"cing�շѻ�����Ϣ��"�еĽ��
                    If lrds������Ϣ("�շѱ��") = cing�շѻ�����Ϣ��.TextMatrix(cing�շѻ�����Ϣ��.Rows - 1, ������Ϣ_�շѱ��) Then
                        cing�����嵥(�շ�).AddItem lrds������Ϣ("�շ���Ŀ���") & vbTab & _
                                                   lrds������Ϣ("�շ���Ŀ����") & vbTab & _
                                                   lrds������Ϣ("����") & vbTab & _
                                                   lrds������Ϣ("����") & vbTab & _
                                                   lrds������Ϣ("���")
                       '�ۼƽ��
                        cing�շѻ�����Ϣ��.TextMatrix(cing�շѻ�����Ϣ��.Rows - 1, ������Ϣ_���) = _
                        CStr(Val(cing�շѻ�����Ϣ��.TextMatrix(cing�շѻ�����Ϣ��.Rows - 1, ������Ϣ_���)) + lrds������Ϣ("���"))
                        mcur�ܽ�� = cing�շѻ�����Ϣ��.TextMatrix(cing�շѻ�����Ϣ��.Rows - 1, ������Ϣ_���)
                    Else
                    '���շѱ�Ų���ͬ,����"cing�շѻ�����Ϣ��"�����Ŀ,�����"cing������Ϣ(�շ�)"�е���Ŀ,��������ӡ�
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    '�켽����2001/12/20���޸�
                    '���ܣ��ڱ������ʾ�ڲ��շѵ���Ϣ
                    'lrds������Ϣ("�շ�����")
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        cing�շѻ�����Ϣ��.AddItem vbTab & lrds������Ϣ("�շ�����") & vbTab & lrds������Ϣ("�շѱ��") & vbTab & lrds������Ϣ("������") & _
                                                   vbTab & lrds������Ϣ("���ѵ�λ����") & vbTab & lrds������Ϣ("���")
                        cing�շѻ�����Ϣ��.Cell(flexcpChecked, cing�շѻ�����Ϣ��.Rows - 1, 0) = 2
                        cing�����嵥(�շ�).Rows = 1
                        cing�����嵥(�շ�).AddItem lrds������Ϣ("�շ���Ŀ���") & vbTab & _
                                                   lrds������Ϣ("�շ���Ŀ����") & vbTab & _
                                                   lrds������Ϣ("����") & vbTab & _
                                                   lrds������Ϣ("����") & vbTab & _
                                                   lrds������Ϣ("���")
                        '������ʾ�շ�������
                        cing�շѻ�����Ϣ��.ColHidden(1) = True
                                               
                    End If
                    If Not lrds������Ϣ.EOF Then lrds������Ϣ.MoveNext
                Next
                
                '��¼��ǰ���շ�����
                mstr�շѱ�� = cing�շѻ�����Ϣ��.TextMatrix(cing�շѻ�����Ϣ��.RowSel, ������Ϣ_�շѱ��)
                If cing�շѻ�����Ϣ��.Rows > 1 Then
                    mcur������Ϣ�ܽ�� = 0
                    For i = 1 To cing�շѻ�����Ϣ��.Rows - 1
                        If cing�շѻ�����Ϣ��.Cell(flexcpChecked, i, 0) = 1 Then
                            mcur������Ϣ�ܽ�� = (mcur������Ϣ�ܽ�� + cing�շѻ�����Ϣ��.TextMatrix(i, ������Ϣ_���))
                        End If
                    Next
                    
                    '���ܣ�����Ԫ�ؿ���,�켽��,2002/07/22
                    Clab�շ���Ŀ����.Enabled = True
                    Ccbo�շ���Ŀ����.Enabled = True
                    lblCaption(5).Enabled = True
                    cinb�շ�����(5).Enabled = True
                    
                End If
                cinb�շ�����(Ӧ�ս��).Text = mcur������Ϣ�ܽ�� * Val(cinb�շ�����(���۱���).Text)
            End If
                        
            If cinb�շ�����(ʵ�ս��).Enabled Then cinb�շ�����(ʵ�ս��).SetFocus
            
        Case "ɾ��"
            
            If Not cing�����嵥(ctabShoufei.Tab).Enabled Then Exit Sub
            If cing�����嵥(ctabShoufei.Tab).RowSel < 1 Then Exit Sub
            ' ���ܣ��ط�����Ϣ��ɾ���ȣ��ڲ��շѲŲ���ɾ���������������ɾ����ʱ�䣺2002/02/06 �켽��
            If cchk�ڲ��շ�.Value = vbChecked And ctabShoufei.Tab = 0 Then
            
                If Not umfuncУ���û�Ȩ��("�շѹ���_�ڲ��շ���Ϣ�޸�") Then
                    sffuncMsg "û���޸��ڲ��շѵ�Ȩ�ޣ�"
                    GoTo WayOut
                End If
                
                If cing�����嵥(ctabShoufei.Tab).Rows - 1 = 1 Then
                     sffuncMsg "������Ϣ�����һ���շ���Ŀ������ɾ����Ҫ��ɾ����ֻ�б������ķ�����Ϣ��"
                     Exit Sub
                Else
                    '��ȡ�ܽ��
                    Dim lcurMoney As Currency
                    For i = 1 To cing�����嵥(ctabShoufei.Tab).Rows - 1
                        lcurMoney = lcurMoney + Val(cing�����嵥(ctabShoufei.Tab).TextMatrix(i, �����嵥_���))
                    Next
                    mcur�ܽ�� = lcurMoney
                    mcur�ܽ�� = mcur�ܽ�� - Val(cing�����嵥(ctabShoufei.Tab).TextMatrix(cing�����嵥(ctabShoufei.Tab).RowSel, �����嵥_���))
                    
                    Dim lstrtemp As String
                    lstrtemp = cing�����嵥(ctabShoufei.Tab).TextMatrix(cing�����嵥(ctabShoufei.Tab).RowSel, 0)
                    cing�����嵥(ctabShoufei.Tab).RemoveItem cing�����嵥(ctabShoufei.Tab).RowSel
                   
                    If mstr�շѱ�� = "" And lstrtemp = "" Then
                        sffuncMsg "������Ϣ�����޷�ɾ����"
                        Exit Sub
                    Else
                        Dim lstr�շ���Ŀ��� As String
                        lstr�շ���Ŀ��� = lstrtemp
                        subɾ��������Ϣ�շ���Ŀ mstr�շѱ��, lstr�շ���Ŀ���
                        
                        '�޸ģ������˶�ɾ���շ�����ϸ��Ϣ��˵�� �г��շ���Ŀ����
                        Dim lstr�շ���Ŀ���� As String
                        
                        lstr�շ���Ŀ���� = cing�����嵥(ctabShoufei.Tab).TextMatrix(cing�����嵥(ctabShoufei.Tab).RowSel, 1)
                        
                        '���ܣ����޸Ĵ�����Ϣ��ʽ���� ʱ�䣺2002/08/05 ���ߣ��켽��
                        sub��Ϣ���� mstr�շѱ��, "�շѱ��Ϊ��" & mstr�շѱ�� & "�ķ�����Ϣ���շ���Ŀ��" & lstr�շ���Ŀ���� & " �ѱ�ɾ����"
                    End If
                     
                    '���½�����ʾ
        
                    For i = 1 To cing�շѻ�����Ϣ��.Rows - 1
                        If cing�շѻ�����Ϣ��.Cell(flexcpText, i, 2) = mstr�շѱ�� Then
                            cing�շѻ�����Ϣ��.TextMatrix(i, 5) = mcur�ܽ�� * Val(cinb�շ�����(���۱���).Text)
                        End If
                    Next
                    
                    Dim lbln�Ƿ���ѡ���� As Boolean
                    lbln�Ƿ���ѡ���� = False
                    
                    For i = 1 To cing�շѻ�����Ϣ��.Rows - 1
                        If cing�շѻ�����Ϣ��.Cell(flexcpChecked, i, 0) = 1 Then
                            lbln�Ƿ���ѡ���� = True
                            Exit For
                        End If
                    Next
                    
                    If lbln�Ƿ���ѡ���� = True Then
                        sub��������ˢ��
                    End If
                    
                    'cinb�շ�����(Ӧ�ս��).Text = mcur�ܽ�� * Val(cinb�շ�����(���۱���).Text)
                End If
                
            Else
                mcur�ܽ�� = mcur�ܽ�� - Val(cing�����嵥(ctabShoufei.Tab).TextMatrix(cing�����嵥(ctabShoufei.Tab).RowSel, �����嵥_���))
                cing�����嵥(ctabShoufei.Tab).RemoveItem cing�����嵥(ctabShoufei.Tab).RowSel
                cinb�շ�����(Ӧ�ս��).Text = mcur�ܽ�� * Val(cinb�շ�����(���۱���).Text)
            End If
        Case "���"
            sub�������
            
        Case "����"
            '����:���Ӷ��ڲ��շ���Ϣ�ı��ϴ���.
            'ʱ��:2002/07/01
            '����:�켽��
            For i = 1 To cing�շѻ�����Ϣ��.Rows - 1
                If cing�շѻ�����Ϣ��.Cell(flexcpChecked, i, 0) = 1 Then
                    lint�շѱ������ = lint�շѱ������ + 1
                    ReDim Preserve lstr�շѱ����(lint�շѱ������)
                    lstr�շѱ����(lint�շѱ������) = cing�շѻ�����Ϣ��.TextMatrix(i, 2)
                End If
            Next
                
            If lint�շѱ������ = 0 Then
                sffuncMsg "��ѡ�е��շ���Ϣ��", sf����
                GoTo WayOut
            End If
            
            If MsgBox("��ȷ��Ҫ����ѡ���еķ�����Ϣ��", vbQuestion + vbYesNo + vbDefaultButton2, "ϵͳ��ʾ") = vbYes Then
                
                
                'ѭ�����Ϸ�����Ϣ
                For i = 1 To UBound(lstr�շѱ����)
                
                    '���ܣ����޸Ĵ�����Ϣ��ʽ���� �켽�� 2002/08/05
                    sub��Ϣ���� lstr�շѱ����(i), "�շѱ��Ϊ��" & lstr�շѱ����(i) & "�ķ�����Ϣ�ѱ����ϣ�"
                    sub���Ϸ�����Ϣ lstr�շѱ����(i)
                Next

                mobj����ͨ�ö���_BeforeOperate "��ѯ", False
                
            End If
 
        Case "�˳�"
            Set lrds������Ϣ = Nothing
        Case Else
    End Select
    GoTo WayOut
    
errhandle:
    If Err.Number = 94 Then Resume Next
    If Err.Number = 5 Then GoTo WayOut
    If Err.Number = 20513 Then
        MsgBox "��ӡ������", vbExclamation, "ϵͳ��ʾ"
        GoTo WayOut
    End If
    
    sfsub������ ������, ģ����, "mobj����ͨ�ö���_BeforeOperate", Err.Number, Err.Description
    
WayOut:
    ctlb������.Buttons("�շ�(&G)").Enabled = True
    ctlb������.Buttons(9).Enabled = True
    Set lcol���� = Nothing
    Set lrds������Ϣ = Nothing
    If cinb�շ�����(�շ�_�շѱ��).Enabled And Operate <> "�˳�" Then cinb�շ�����(�շ�_�շѱ��).SetFocus
    Exit Sub
'    Resume
    
errTransHandler:
    dasubRollBack
    GoTo errhandle
End Sub

'�����շѱ��ɾ��������Ϣ���շ���Ŀ.
'ʱ��:2002/07/02
'����:�켽��
Private Sub subɾ��������Ϣ�շ���Ŀ(ByVal para�շѱ�� As String, ByVal Para��Ŀ��� As String)
On Error GoTo errhandler
    Dim lstrSql As String     '���������¼SQL���
    
    lstrSql = "delete from �շѹ���_������Ϣ�� where �շѱ��='" & para�շѱ�� & "'" & _
              " and �շ���Ŀ���='" & Para��Ŀ��� & "'"
    dafuncGetData (lstrSql)
Exit Sub
errhandler:
     sfsub������ "�շѽ���", "frm�շ�", "subɾ��������Ϣ�շ���Ŀ", Err.Number, Err.Description
End Sub


'�����շѱ��ɾ��������Ϣ.
'ʱ��:2002/0607/01
'����:�켽��
Public Sub sub���Ϸ�����Ϣ(ByVal para�շѱ��)
On Error GoTo errhandler
    Dim lstrSql As String           '���������¼SQL���
    
    lstrSql = "delete from �շѹ���_������Ϣ�� where �շѱ��='" & para�շѱ�� & "'"
    dafuncGetData (lstrSql)
Exit Sub
errhandler:
    sfsub������ "�շѽ���", "frm�շ�", "sub���Ϸ�����Ϣ", Err.Number, Err.Description
End Sub


'����: ���������ת��Ϊ����ҵĴ�д�ַ���
'����: money       ���
'���: FuncConvertToCapsStr     ת���Ĵ�д�ַ���
'����޸�ʱ��: 96.6.11
'--------------------------------------------------
Public Function FuncConvertToCapsStr(Money As Currency) As String
On Error GoTo errhandle
    Const digit_str = "��Ҽ��������½��ƾ�"
    Const unit_str = "Ǫ��ʰ��Ǫ��ʰԪ�Ƿ�"
    Dim money_str As String
    
    If Money > 99999999.99 Then
        FuncConvertToCapsStr = ""
    ElseIf Money = 0 Then
        FuncConvertToCapsStr = "��Ԫ��"
    Else
        Dim temp_str As String
        Dim i, j As Integer
        
        If Money < 0 Then
            money_str = "��"
            Money = -Money
        Else
            money_str = ""
        End If
        
        temp_str = Format(Money, "00000000.00")
        
        'ת����������
        For i = 1 To 8
            If Mid(temp_str, i, 1) <> "0" Then Exit For
        Next
        For i = i To 8
            j = CInt(Mid(temp_str, i, 1))
            If j > 0 Then
                money_str = money_str & Mid(digit_str, j + 1, 1) & Mid(unit_str, i, 1)
            Else
                If i = 4 Then
                    money_str = money_str & "��"
                ElseIf i = 8 Then
                    money_str = money_str & "Ԫ"
                ElseIf Mid(temp_str, i + 1, 1) <> "0" Then
                    money_str = money_str & Mid(digit_str, j + 1, 1)
                End If
            End If
        Next
        
        'ת��С������
        If Right(temp_str, 2) = "00" Then
            money_str = money_str & "��"
        Else
            'ת����
            j = CInt(Mid(temp_str, 10, 1))
            If j > 0 Then
                money_str = money_str & Mid(digit_str, j + 1, 1) & "��"
            Else
                money_str = money_str & "��"
            End If
            'ת����
            j = CInt(Mid(temp_str, 11, 1))
            If j > 0 Then
                money_str = money_str & Mid(digit_str, j + 1, 1) & "��"
            Else
                money_str = money_str & "��"
            End If
        End If
        
        FuncConvertToCapsStr = money_str
    End If
Exit Function
errhandle:
    sfsub������ "�շѽ���", "frm�շ�", " FuncConvertToCapsStr()", Err.Number, Err.Description
End Function

Private Function func�����Ŀ�Ƿ���ѡ(ByVal Para�շ���Ŀ��� As String) As Boolean
On Error GoTo errhandle
    Dim i As Long
    func�����Ŀ�Ƿ���ѡ = False
    If cing�����嵥(ctabShoufei.Tab).Rows = 1 Then
        func�����Ŀ�Ƿ���ѡ = False
        Exit Function
    End If
    
    For i = 1 To cing�����嵥(ctabShoufei.Tab).Rows - 1
        If Para�շ���Ŀ��� = cing�����嵥(ctabShoufei.Tab).TextMatrix(i, �����嵥_�շ���Ŀ���) Then
            func�����Ŀ�Ƿ���ѡ = True
            Exit Function
        End If
    Next
Exit Function
errhandle:
    sfsub������ "�շѽ���", "frm�շ�", " func�����Ŀ�Ƿ���ѡ()", Err.Number, Err.Description
End Function

Private Function ValidateData() As Boolean
On Error GoTo errhandle
    Select Case ctabShoufei.Tab
        Case �շ�
            
            If cchk�ڲ��շ�.Value = 0 Then
                If cinb�շ�����(�շ�_������).Text = vbNullString And cinb�շ�����(�շ�_���ѵ�λ) = vbNullString Then
                    ValidateData = False
                    sffuncMsg """������"" �� ""���ѵ�λ"" ������������֮һ��", sf����
                Else
                    ValidateData = True
                End If
                If cinb�շ�����(�շ�_���ܿ���).Text = "" Then
                    ValidateData = False
                    sffuncMsg "���������ܿ��ң�", sf����
                End If
            Else
                ValidateData = True
            End If
        Case �����շ�
            If cinb�շ�����(�����շ�_������).Text = vbNullString And cinb�շ�����(�����շ�_���ѵ�λ) = vbNullString Then
                ValidateData = False
                sffuncMsg """������"" �� ""���ѵ�λ"" ������������֮һ��", sf����
            Else
                ValidateData = True
            End If
            If cinb�շ�����(�����շ�_���ܿ���).Text = "" Then
                ValidateData = False
                sffuncMsg "���������ܿ��ң�", sf����
            End If
    End Select
Exit Function
errhandle:
    sfsub������ "�շѽ���", "frm�շ�", " ValidateData()", Err.Number, Err.Description
End Function


Private Function funcƥ���շ���Ŀ(ByVal paraValue As String) As Long
    Dim i As Long
    
    On Error GoTo errhandle
    
    funcƥ���շ���Ŀ = 0
    
    If paraValue = vbNullString Then Exit Function
    '�� paraValue ��λΪ����,��ƥ����
    If Asc(paraValue) >= Asc("0") And Asc(paraValue) <= Asc("9") Then
        For i = 1 To cind�ֵ�(�ֵ�_�շ���Ŀ).Rows - 1
            If Left(cind�ֵ�(�ֵ�_�շ���Ŀ).TextMatrix(i, �ֵ�_�շ���Ŀ���), Len(paraValue)) = paraValue Then
                cind�ֵ�(�ֵ�_�շ���Ŀ).TopRow = i
                cind�ֵ�(�ֵ�_�շ���Ŀ).Select i, 0
                funcƥ���շ���Ŀ = i
                Exit Function
            End If
        Next
    End If
    
    '�� paraValue ��λΪ��ĸ,��ƥ�����Ƿ�
    If Asc(paraValue) >= Asc("A") And Asc(paraValue) <= Asc("z") Then
        For i = 1 To cind�ֵ�(�ֵ�_�շ���Ŀ).Rows - 1
            If Left(cind�ֵ�(�ֵ�_�շ���Ŀ).TextMatrix(i, �ֵ�_���Ƿ�), Len(paraValue)) = paraValue Then
                cind�ֵ�(�ֵ�_�շ���Ŀ).TopRow = i
                cind�ֵ�(�ֵ�_�շ���Ŀ).Select i, 0
                funcƥ���շ���Ŀ = i
                Exit Function
            End If
        Next
    End If
    
    '�������ƥ������

    For i = 1 To cind�ֵ�(�ֵ�_�շ���Ŀ).Rows - 1
        If Left(cind�ֵ�(�ֵ�_�շ���Ŀ).TextMatrix(i, �ֵ�_�շ���Ŀ����), Len(paraValue)) = paraValue Then
            cind�ֵ�(�ֵ�_�շ���Ŀ).TopRow = i
            cind�ֵ�(�ֵ�_�շ���Ŀ).Select i, 0
            funcƥ���շ���Ŀ = i
            Exit Function
        End If
    Next
    
errhandle:
    If Err.Number = 0 Then Exit Function
    funcƥ���շ���Ŀ = 0
    sfsub������ "�շѽ���", "frm�շ�", "funcƥ���շ���Ŀ", Err.Number, Err.Description
End Function

Private Function funcƥ���շѱ�׼(ByVal paraValue As String) As Long
On Error GoTo errhandle
    Dim i As Long
    funcƥ���շѱ�׼ = 0
    If paraValue = vbNullString Then Exit Function
    '�� paraValue ��λΪ����,��ƥ����
    If Asc(paraValue) >= "0" And Asc(paraValue) <= Asc("9") Then
        For i = 1 To cind�ֵ�(�ֵ�_�շѱ�׼).Rows - 1
            If Left(cind�ֵ�(�ֵ�_�շѱ�׼).TextMatrix(i, �ֵ�_�շѱ�׼���), Len(paraValue)) = paraValue Then
                cind�ֵ�(�ֵ�_�շѱ�׼).TopRow = i
                cind�ֵ�(�ֵ�_�շѱ�׼).Select i, 0
                funcƥ���շѱ�׼ = i
                Exit Function
            End If
        Next
    End If
    
    '�� paraValue ��λΪ��ĸ,��ƥ�����Ƿ�
    If Asc(paraValue) >= Asc("A") And Asc(paraValue) <= Asc("z") Then
        For i = 1 To cind�ֵ�(�ֵ�_�շѱ�׼).Rows - 1
            If Left(cind�ֵ�(�ֵ�_�շѱ�׼).TextMatrix(i, �ֵ�_���Ƿ�), Len(paraValue)) = paraValue Then
                cind�ֵ�(�ֵ�_�շѱ�׼).TopRow = i
                cind�ֵ�(�ֵ�_�շѱ�׼).Select i, 0
                funcƥ���շѱ�׼ = i
                Exit Function
            End If
        Next
    End If
    '����ƥ������
    For i = 1 To cind�ֵ�(�ֵ�_�շѱ�׼).Rows - 1
        If Left(cind�ֵ�(�ֵ�_�շѱ�׼).TextMatrix(i, �ֵ�_�շѱ�׼����), Len(paraValue)) = paraValue Then
            cind�ֵ�(�ֵ�_�շѱ�׼).TopRow = i
            cind�ֵ�(�ֵ�_�շѱ�׼).Select i, 0
            funcƥ���շѱ�׼ = i
            Exit Function
        End If
    Next
Exit Function
errhandle:
    sfsub������ "�շѽ���", "frm�շ�", "funcƥ���շѱ�׼", Err.Number, Err.Description
End Function

Private Sub Disable()
On Error Resume Next
    Dim i As Control
    For Each i In Controls
        i.Enabled = False
    Next
    ctlb������.Buttons("�˳�(ESC)").Enabled = True
End Sub


'���ܣ�ͨ����ʹ�ͻ��ˣ���ָ���Ŀ��ҷ�����Ϣ
'���룺ParaĿ�Ŀ��ұ��,Para��Ϣ����
'ʱ�䣺2002/08/05
'���ߣ��켽��

Private Sub sub��Ϣ����(ByVal para�շѱ�� As String, ByVal Para��Ϣ���� As String)
On Error GoTo errhander
    Dim lstrĿ�Ŀ��Һ� As String        '���������¼Ŀ�Ŀ��Һ�
    Dim lstrSql As String               '���������¼SqL���
    Dim lobjRec As Object               '��������¼�����ݼ�
    Dim lstr������ As String            '���������¼������
    Dim lstr���ѵ�λ As String          '���������¼���ѵ�λ
    
    
    '����û����ұ�Ż���û����Ϣ���ݾ��˳�����
    If para�շѱ�� = "" Or Para��Ϣ���� = "" Then
        Exit Sub
    Else
        
        lstrSql = "select distinct(�շѱ��),������,���ѵ�λ����,���ܿ��ұ�� from  �շѹ���_������Ϣ�� where �շѱ��='" & para�շѱ�� & "'"
        Set lobjRec = dafuncGetData(lstrSql)
        If lobjRec.RecordCount > 0 Then
            lstrĿ�Ŀ��Һ� = IIf(IsNull(lobjRec("���ܿ��ұ��")), "", lobjRec("���ܿ��ұ��"))
            lstr������ = IIf(IsNull(lobjRec("������")), "����", lobjRec("������"))
            lstr���ѵ�λ = IIf(IsNull(lobjRec("���ѵ�λ����")), "����", lobjRec("���ѵ�λ����"))
            Para��Ϣ���� = Para��Ϣ���� & "��������" & lstr������ & "�����ѵ�λ��" & lstr���ѵ�λ & "��"
            
            '������Ϣ�����ǿ�ѡ��װ,�����ڷ���ʱҪ��������Ƿ���� �켽�� 2002/09/17
            If Not um��ʹ�ͻ��� Is Nothing Then
                um��ʹ�ͻ���.sub������Ϣ um�û��������ұ��, lstrĿ�Ŀ��Һ�, Para��Ϣ����, "�����޸�"
            End If
        End If
    End If
Exit Sub
errhander:
    'sfsub������ "�շѽ���", "frm�շ�", "sub��Ϣ����", Err.Number, Err.Description
End Sub

