VERSION 5.00
Object = "{8099FCC2-0A81-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6.ocx"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmInputTestResult 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "�����¼��"
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
   StartUpPosition =   3  '����ȱʡ
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
            Name            =   "����"
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
         Caption         =   "���ڣ�"
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
         Caption         =   "ҽʦ��"
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
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "   ��������(&F7)   "
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
      TabCaption(1)   =   "   ��������(&F8)   "
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
      Tab(1).Control(8)=   "ctxt��λ����"
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
         Caption         =   "ϵͳ���"
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   45
         Top             =   1800
         Width           =   1095
      End
      Begin VB.OptionButton coptBatchType 
         Caption         =   "��λ����"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   44
         Top             =   1320
         Width           =   1095
      End
      Begin VB.OptionButton coptBatchType 
         Caption         =   "�������"
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
         Caption         =   "������"
         Height          =   255
         Index           =   1
         Left            =   4320
         TabIndex        =   39
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CheckBox cchkUnEnd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "δ����"
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
         Caption         =   "������"
         Height          =   255
         Index           =   0
         Left            =   -72000
         TabIndex        =   37
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CheckBox cchkUnEnd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "δ����"
         Height          =   255
         Index           =   0
         Left            =   -73080
         TabIndex        =   36
         Top             =   2760
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.OptionButton coptType 
         Caption         =   "����"
         Height          =   255
         Index           =   3
         Left            =   -71040
         TabIndex        =   35
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox ctxt��λ���� 
         Height          =   375
         Left            =   1320
         TabIndex        =   34
         Top             =   1320
         Width           =   2415
      End
      Begin VB.OptionButton coptType 
         Caption         =   "��쵥��"
         Height          =   255
         Index           =   2
         Left            =   -72240
         TabIndex        =   33
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton coptType 
         Caption         =   "�Թܱ��"
         Height          =   255
         Index           =   1
         Left            =   -73560
         TabIndex        =   32
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton coptType 
         Caption         =   "ϵͳ���"
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
         Caption         =   "��ѯ(&Q)"
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
            Name            =   "����"
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
         Caption         =   "���(&C)"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2160
         Width           =   975
      End
      Begin VB.CommandButton ccmdRemove 
         Caption         =   "ȥ��(&R)"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1140
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2160
         Width           =   975
      End
      Begin VB.CommandButton ccmdAdd 
         Caption         =   "���(&A)"
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
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
         FormatString    =   "^ϵͳ���         |�Թܱ��|^����    |^�Ա�|^��λ����               |^����|^��쵥��"
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
            Name            =   "����"
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
            Name            =   "����"
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
         FormatString    =   "^ϵͳ���         |�Թܱ��|^����    |^�Ա�|^��λ����               |^����|^��쵥��"
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
         Caption         =   "ѡ������"
         Height          =   180
         Left            =   120
         TabIndex        =   25
         Top             =   480
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Թܱ�ţ�"
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
         Caption         =   "��λ���ƣ�"
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
         Caption         =   "���䣺"
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
         Caption         =   "�Ա�"
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
         Caption         =   "������"
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
         Caption         =   "ϵͳ��ţ�"
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
         Name            =   "����"
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
         Caption         =   "��ɫ��ʾ���ϸ�,����˫�������޸�"
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
         Caption         =   "�����Ŀ�����"
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
         Name            =   "����"
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
      ImageList       =   "cimg��ťͼ��"
      _Version        =   393216
      Begin VB.CheckBox cchkGrid 
         Caption         =   "����¼��"
         Height          =   255
         Left            =   8520
         TabIndex        =   46
         Top             =   200
         Width           =   1455
      End
      Begin VB.CheckBox cchkˢ���� 
         Caption         =   "ˢ����"
         Height          =   375
         Left            =   7320
         TabIndex        =   6
         Top             =   120
         Width           =   1095
      End
   End
   Begin MSComctlLib.ImageList cimg��ťͼ�� 
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
'���ߣ��

Private WithEvents mobj����ͨ�ö��� As cls����ͨ�ö���    '�ṩ��������ʼ�����ȼ�����
Attribute mobj����ͨ�ö���.VB_VarHelpID = -1

Private mobj���ҽʦ  As Object   'clsMedicalExamer    ��ȡ��ǰ���ҽʦ��������ָ�����ԣ�����/���飩�������Ŀ

Private mstr�������� As String  '��������ʱ����ǰһ������¼��ʹ�õ�����ģ�����ơ�
Private mstr��������  As String   '��Ӧ����"��������"��
Private mstr�����Ŀ���� As String
Private mstrϵͳ��Ź̶����� As String

Private mblnSys As Boolean

Private mblnInUse As Boolean      '��Ӧ����"pblnInUse"

Private mobj���� As cls�û��������� '�޸ģ�2001-12-29�����Ӹö��󣩡�

Private mintFixed As Integer
Private mcol�����Ŀ As Collection

'���ܣ�������ǰ�����Ƿ�һ���أ��Ա������������жϵ�ǰ�����Ƿ���ִ�й�Form_Load��
Public Property Get pblnInUse() As Boolean
Attribute pblnInUse.VB_Description = "'���ܣ�������ǰ�����Ƿ�һ���أ��Ա������������жϵ�ǰ�����Ƿ���ִ�й�Form_Load��"
    pblnInUse = mblnInUse
End Property

Private Sub cchkEnd_Click(Index As Integer)
    If Index = 0 Then
        ccmdSingleQuery_Click
    Else
        ccmdAdd_Click
    End If
End Sub



'2006-6-19(����¼�룩
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
    
    '��ʾָ��������ڵ�δ�½��۵������Ա������
    subShowSingleList
    
    Exit Sub
    
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "�����沿��", "FrmInputTestResult", "ccmdSingleQuery_Click", 6666, lstrError, False
End Sub




Private Sub cgrdInput_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lstr������� As String
    On Error GoTo errHandler
    If Row > 0 Then
        lstr������� = pobjҵ�����.func��ȡ�������(cgrdInput.TextMatrix(Row, 0), cgrdInput.TextMatrix(Row, 2))
        If lstr������� = "���ϸ�" Then
            '������ɫ��
            cgrdInput.Cell(flexcpBackColor, Row, 2, Row, 2) = &H8A5AFA
        Else
            '������ɫ��
            cgrdInput.Cell(flexcpBackColor, Row, 2, Row, 2) = vbWhite
        End If
    End If
    Exit Sub
errHandler:
End Sub

Private Sub cgrdInput_DblClick()
    '�޸���ɫ��
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
        '���С�
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
'2006-6-19(����¼��)
Private Sub cgrdPerson_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lstr������� As String
    On Error GoTo errHandler
    If Row > 0 Then
        lstr������� = pobjҵ�����.func��ȡ�������(cgrdPerson.TextMatrix(Row, 0), cgrdPerson.TextMatrix(Row, Col))
        If lstr������� = "���ϸ�" Then
            '������ɫ��
            cgrdPerson.Cell(flexcpBackColor, Row, Col, Row, Col) = &H8A5AFA
        Else
            '������ɫ��
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
    Dim lobj��� As Object
    Dim lstrNo As String 'ϵͳ��š�
    
    On Error GoTo errHandler
    If cgrdPerson.Row < 1 Or mblnSys Then Exit Sub
    MousePointer = 11
    
    cstbMain.Panels(1).Text = "���ڻ�ȡ�����Ա��Ƭ�����Ժ�..."
    
    '���水ʱ���ɲ�����
    ctbMain.Enabled = False
    ctabPerson.Enabled = False
    Frame1.Enabled = False
    
    '��ȡϵͳ��š�
    lstrNo = cgrdPerson.TextMatrix(cgrdPerson.Row, 0)
    
    '����������
    Set lobj��� = CreateObject("������.clsMedicalExam")
    
    lobj���.ϵͳ��� = lstrNo
        
    '��ʾ��Ƭ��
    If Not lobj���.�����Ա.��Ƭ Is Nothing Then
        cpicPhoto(1).Picture = lobj���.�����Ա.��Ƭ
    Else
        cpicPhoto(1).Picture = Nothing
    End If
    
errHandler:
    If Err <> 0 Then
        Dim lstrError As String
        lstrError = func������(Err.Number, Err.Description)
        sfsub������ "�����沿��", "FrmInputTestResult", "cgrdPerson_Click", 6666, lstrError, False
    End If
    Set lobj��� = Nothing
    '����ָ����Բ�����
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
    Dim lobj��� As Object
    Dim lstrSysNo As String 'ϵͳ��š�
    
    On Error GoTo errHandler
    If cgrdSingleList.Row < 1 Or mblnSys Then Exit Sub
    MousePointer = 11
    
    cstbMain.Panels(1).Text = "���ڻ�ȡ�����Ա��Ƭ�����Ժ�..."
    
    '���水ʱ���ɲ�����
    ctbMain.Enabled = False
    ctabPerson.Enabled = False
    Frame1.Enabled = False
    
    '��ȡϵͳ��š�
    lstrSysNo = cgrdSingleList.TextMatrix(cgrdSingleList.Row, 0)
    
    '����������
    Set lobj��� = CreateObject("������.clsMedicalExam")
    
    lobj���.ϵͳ��� = lstrSysNo
        
    '��ʾ������Ϣ��
    With lobj���.�����Ա
        clblInfo(0) = .����
        clblInfo(1) = .�Ա�
        clblInfo(2) = .����
        clblInfo(3) = .��λ����
        ctxtSingleNo.Text = lstrSysNo
        clblInfo(4) = lobj���.�Թܱ��
        
        '��ʾ��Ƭ��
        If Not .��Ƭ Is Nothing Then
            cpicPhoto(0).Picture = .��Ƭ
        Else
            cpicPhoto(0).Picture = Nothing
        End If
    End With
    
    '���������¼������
    subShowInputGrid lstrSysNo
    
    ctbMain.Buttons(1).Enabled = True
    
errHandler:
    If Err <> 0 Then
        Dim lstrError As String
        lstrError = func������(Err.Number, Err.Description)
        sfsub������ "�����沿��", "FrmInputTestResult", "cgrdSingleList_SelChange", 6666, lstrError, False
    End If
    '����ָ����Բ�����
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
            ctxt��λ����.SetFocus
        Else
            ctxtBatchNo.SetFocus
        End If
    End If
    If coptBatchType(2).Value Then
        cchkˢ����.Visible = True
    Else
        cchkˢ����.Visible = False
        cchkˢ����.Value = 0
    End If
    
End Sub

Private Sub coptType_Click(Index As Integer)
    On Error Resume Next
    Label1(3).Caption = coptType(Index).Caption
    ctxtSingleNo.Text = ""
    ctxtSingleNo.SetFocus
    If coptType(0).Value Then
        cchkˢ����.Visible = True
    Else
        cchkˢ����.Visible = False
        cchkˢ����.Value = 0
    End If
End Sub

Private Sub ctxtBatchNo_GotFocus()
    On Error Resume Next
    '������ϵͳ��ţ�����ϵͳ��ŵĹ̶����֣�����¼�롣
    If ctxtBatchNo.Text = "" And cchkˢ����.Value = 0 And cchkˢ����.Visible Then
        ctxtBatchNo.Text = mstrϵͳ��Ź̶�����
        ctxtBatchNo.SelLength = 0
        ctxtBatchNo.SelStart = Len(mstrϵͳ��Ź̶�����)
    End If
End Sub

Private Sub ctxtBatchNo_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 And Trim(ctxtBatchNo.Text) <> "" Then
        ccmdAdd_Click
        If cchkˢ����.Value = 1 Then
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
    '�����Ƚ���
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
        '���������롰'����
        KeyAscii = 0
    End If

End Sub

Private Sub Form_Load()
    Dim i As Long
    
    On Error GoTo errHandler
    
    If mblnInUse Then Exit Sub
    
    '���ô����Ѽ��ر�־��
    mblnInUse = True
    
    '����mobj����ͨ�ö��󣬳�ʼ����������
    Dim lcol������ As Collection
    Set lcol������ = New Collection
    With lcol������
        .Add "����"
        .Add "|"
        .Add "�˳�"
    End With
    Set mobj����ͨ�ö��� = New cls����ͨ�ö���
    With mobj����ͨ�ö���
        Set .Form = Me
        Set .c������ = ctbMain
        Set .c״̬�� = cstbMain
    
        .subInitialize lcol������, ""
    End With
    
    '��ʼ�������¼������
    With cgrdInput
        .Rows = 1
        .Cols = 6
        .TextMatrix(0, 0) = "����"
        .ColWidth(0) = 700
        .TextMatrix(0, 1) = "����"
        .ColWidth(1) = 1600
        .TextMatrix(0, 2) = "�����"
        .ColWidth(2) = 1000
        .ColHidden(3) = True '��Ÿ������Ŀ��ö����Դ��
        .ColWidth(4) = 1200
        .TextMatrix(0, 4) = "��׼ֵ"
        .ColWidth(5) = 600
        .TextMatrix(0, 5) = "��λ"
        .ShowComboButton = True
        
        .Editable = True
    End With
    '�����������ȫ�ֶ���mobj���ҽʦ��
    Set mobj���ҽʦ = CreateObject("������.clsMedicalExaminer")
    mobj���ҽʦ.��� = um�û����
    
    '��ȡϵͳ��Ź̶����֡�
    Dim lobj��� As Object '�����󣬻�ȡϵͳ��ŵĹ̶����֡�
    Set lobj��� = CreateObject("������.clsMedicalExam")
    mstrϵͳ��Ź̶����� = lobj���.ϵͳ��Ź̶�����
    
    ctxtSingleNo.Text = mstrϵͳ��Ź̶�����
    
    '���ҽʦ"��ʾ����ʾ��ǰ�û�����
    ctxtDoctor.Text = um�û���
    cdtpInputDate.Value = Date
    cdtpQueryDate.Value = Date
    cdtpSingleQuery.Value = Date
    
    cgrdInput.Rows = 1
    ctxtSingleNo.TabIndex = 0
    
        
    '�޸ģ�2001-11-2���������ʾ�����������ơ�
    Dim lobj����ģ�弯 As Object
    Dim lcolInfo As Collection
    Set lobj����ģ�弯 = CreateObject("������.ClsMedicalExamTemplateSet")
    Set lcolInfo = lobj����ģ�弯.Ԫ�ؼ�
    
    ccmbSheet.Clear
    '�޸ģ�2002-8-14��������ӡ�<����>��ѡ���
    If lcolInfo.Count > 0 Then
        ccmbSheet.AddItem "<����>"
    End If
    For i = 1 To lcolInfo.Count
        ccmbSheet.AddItem lcolInfo(i)
    Next i
    If ccmbSheet.ListCount > 1 Then
        ccmbSheet.ListIndex = 1
    Else
        cstbMain.Panels(1).Text = "���Ƚ��롰�������á�����������������"
    End If
    Set lcolInfo = Nothing
    Set lobj����ģ�弯 = Nothing
    
    '�޸ģ�2001-12-29����ȡ��������ֵ����
    On Error Resume Next
    Set mobj���� = New cls�û���������
    mobj����.�û���� = um�û����
    mobj����.ҵ���� = "������"
    
    '�޸ģ�2002-9-26�������ȡˢ�����������ֵ��
    If mobj����.������ֵ("¼����ʱˢ����") = "��" Then
        cchkˢ����.Value = 1
    Else
        cchkˢ����.Value = 0
    End If
    
    cchkGrid.Visible = False
    ctabPerson.Tab = 0
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "�����沿��", "FrmInputTestResult", "Form_Load", 6666, lstrError, False
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
        '����¼�롣
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
            '����¼��
            Frame3.Top = 1100
'            cgrdPerson.SelectionMode = flexSelectionFree
            cgrdPerson.Editable = True
        End If
        
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    '�޸ģ�2002-9-26����������������ֵ��
    mobj����.sub���Ǽ���ֵ "¼����ʱˢ����", IIf(cchkˢ����.Value = 1, "��", "��")
    
    '�ͷű������ȫ�ֶ���
    Set mobj����ͨ�ö��� = Nothing
    Set mobj���ҽʦ = Nothing
    Set mobj���� = Nothing
    
    '���ñ�־pblnInUse�����������Ѳ���ʹ�á�
    mblnInUse = False
    
End Sub

'���ܣ������Ա������
Private Sub ccmdAdd_Click()
    On Error GoTo errHandler
    '�޸ģ�2001-11-2������ж��Ƿ���������û�У�������ʾ��
    If ccmbSheet.ListCount = 0 Then
        sffuncMsg "���Ƚ��롰�������á������������ú�����", sf����
        Exit Sub
    End If
    If coptBatchType(1).Value And ctxt��λ���� = "" Then
        MsgBox "�����뵥λ���ƣ�", vbOKOnly + vbExclamation, "ϵͳ��ʾ"
        ctxt��λ����.SetFocus
        Exit Sub
    End If
    If coptBatchType(2).Value And ctxtBatchNo = "" Then
        MsgBox "������ϵͳ��ţ���ˢ�������ϵ����룩��", vbOKOnly + vbExclamation, "ϵͳ��ʾ"
        ctxtBatchNo.SetFocus
        Exit Sub
    End If
    
    MousePointer = 11
    '����¼�뷽ʽ��
    
    '��ʾһ�������Ա�������С�
    subShowBatchPerson
        
    
    '�жϡ����桱��ť�Ƿ���á�
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
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "�����沿��", "FrmInputTestResult", "ccmdAdd_Click", 6666, lstrError, False
    cdtpQueryDate.SetFocus
    cstbMain.Panels(1) = ""
    MousePointer = 0
    Exit Sub
    Resume
End Sub

Private Sub cchkˢ����_Click()
    On Error GoTo errHandler
    If Not cchkˢ����.Visible Then Exit Sub
    
    If cchkˢ����.Value = 1 Then
        '���ϵͳ��������
        ctxtSingleNo = ""
        ctxtBatchNo = ""
    Else
        '���������Թܱ�š�
        If ctabPerson.Tab = 0 Then
            If ctxtSingleNo = "" Then
                ctxtSingleNo = mstrϵͳ��Ź̶�����
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
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "�����沿��", "FrmInputTestResult", "cdtpQueryDate_KeyDown", 6666, lstrError, False
End Sub

Private Sub cgrdInput_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim lstrEnum As String   '��ǰ�������ö����Դ����Ӣ�Ķ��Ż����Ķ��Ÿ�������
    Dim i As Long
    
    On Error GoTo errHandler
    'ֻ��������п���¼�롣
    If Col <> 2 Then
        Cancel = True
    Else
        '������������д�ŵ�ö����Դ���õ�ǰ��Ԫ�������б�
        lstrEnum = cgrdInput.TextMatrix(cgrdInput.Row, 3)
        If lstrEnum = "" Then
            'û��ö����Դ��
            cgrdInput.ColComboList(2) = ""
            
        Else
            '����������е�¼�뷽ʽΪ����ѡ��
            cgrdInput.ColComboList(2) = "|" & lstrEnum
        End If
        
    End If
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "�����沿��", "FrmInputTestResult", "cgrdInput_BeforeEdit", 6666, lstrError, False
    Exit Sub
    Resume
End Sub


Private Sub ctxtSingleNo_GotFocus()
    On Error Resume Next
    '������ϵͳ��ţ�����ϵͳ��ŵĹ̶����֣�����¼�롣
    If ctxtSingleNo.Text = "" And cchkˢ����.Value = 0 And cchkˢ����.Visible Then
        ctxtSingleNo.Text = mstrϵͳ��Ź̶�����
        ctxtSingleNo.SelLength = 0
        ctxtSingleNo.SelStart = Len(mstrϵͳ��Ź̶�����)
        ctbMain.Buttons(1).Enabled = False
        
    End If
End Sub

Private Sub ctxtSingleNo_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errHandler
    
    If KeyCode = 13 And Trim(ctxtSingleNo.Text) <> "" Then
        '��ʾ��Ա��Ϣ��
        subShowSinglePerson
            
        If cgrdInput.Rows > 1 Then
            ctbMain.Buttons(1).Enabled = True
        End If
    End If
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "�����沿��", "FrmInputTestResult", "ctxtSingleNo_KeyDown", 6666, lstrError, False
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

'���ܣ�����ϵͳ��ţ����������¼������
'���ߣ����
Private Sub subShowInputGrid(ByVal paraSysNo As String)
    Dim lcolInfo As Collection '��ŵ�ǰϵͳ����е�ǰҽʦ��������ָ�����Ե������Ŀ��������
    Dim lobjItem As Variant    'clsFactTestItem,lcolInfo�е�Ԫ�ء�
    Dim lstrEnum As String
    Dim i As Long
    Dim j As Long
    On Error GoTo errHandler
    
    'ʹ���Ż��㷨��ȡ�����Ŀ��
    '�޸ģ�2001-4-22(�)��
    Dim lobjRec As Object
    Dim lobjTemp As Object
    
    '��ȡָ�����ԣ�����/���飩�������Ŀ��clsFactTestItem(�����Ŀ���룬�����Ŀ���ƣ�ȱʡֵ��ö����Դ�������)��
    '�޸ģ�2002-10-14�������ѡ�������������ȡ���������Ͽ�����Ŀ��
    If ctabPerson.Tab = 1 Then
        If ccmbSheet.Text = "<����>" Then
            Set lobjRec = mobj���ҽʦ.Func��ȡ�������������Ͽ����������Ŀ(mstr�����Ŀ����)
        Else
            Set lobjRec = mobj���ҽʦ.Func��ȡָ�������Ͽ����������Ŀ(ccmbSheet.Text, mstr�����Ŀ����)
        End If
    Else
        Set lobjRec = mobj���ҽʦ.Func�Ż��Ļ�ȡ���˿����������Ŀ(paraSysNo, mstr�����Ŀ����)
    End If
    
    '��ʾ�����Ŀ��cgrdInput�С�
    cgrdInput.Rows = 1
    
    Set mcol�����Ŀ = New Collection
    
    If lobjRec.recordcount > 0 Then
        cgrdInput.Rows = lobjRec.recordcount + 1
        i = 1
        Do While Not lobjRec.EOF
            cgrdInput.TextMatrix(i, 0) = lobjRec!�����Ŀ���
            cgrdInput.TextMatrix(i, 1) = lobjRec!�����Ŀ����
            cgrdInput.TextMatrix(i, 2) = IIf(IsNull(lobjRec!�����), "", lobjRec!�����)
            
            '���ݵ������������ɫ��
            Dim lstr������� As String
            If IIf(IsNull(lobjRec!�������), "", lobjRec!�������) = "" And cgrdInput.TextMatrix(i, 2) <> "" Then
                '�����µ�����ۡ�
                lstr������� = pobjҵ�����.func��ȡ�������(lobjRec!�����Ŀ���, IIf(IsNull(lobjRec!�����), "", lobjRec!�����))
            Else
                lstr������� = IIf(IsNull(lobjRec!�������), "", lobjRec!�������)
            End If
            If lstr������� = "���ϸ�" Then
                '������ɫ��
                cgrdInput.Cell(flexcpBackColor, i, 2, i, 2) = &H8A5AFA
            Else
                '������ɫ��
                cgrdInput.Cell(flexcpBackColor, i, 2, i, 2) = vbWhite
            End If
            '��ö����Դ��ת��ΪGrid����ʶ���ColComboList�����ԡ�|����������
            lstrEnum = IIf(IsNull(lobjRec!ö����Դ), "", lobjRec!ö����Դ)
            lstrEnum = gffuncStrReplace(lstrEnum, ",", "|")
            lstrEnum = gffuncStrReplace(lstrEnum, "��", "|")
            cgrdInput.TextMatrix(i, 3) = lstrEnum

            cgrdInput.TextMatrix(i, 4) = IIf(IsNull(lobjRec!��׼ֵ), "", lobjRec!��׼ֵ)
            cgrdInput.TextMatrix(i, 5) = IIf(IsNull(lobjRec!��λ), "", lobjRec!��λ)

            '2006-6-19(Ϊ�˽�������¼�룬��cgrdperson������������ʾ�����Ŀ����)��
            If ctabPerson.Tab = 1 Then
                For j = mintFixed + 1 To cgrdPerson.Cols - 1
                    If cgrdPerson.TextMatrix(0, j) = lobjRec!�����Ŀ���� Then Exit For
                Next
                If j = cgrdPerson.Cols Then
                    cgrdPerson.Cols = cgrdPerson.Cols + 1
                    cgrdPerson.TextMatrix(0, j) = lobjRec!�����Ŀ����
                    
                    cgrdPerson.ColHidden(j) = IIf(cchkGrid.Value = 0, True, False)
                    If lstrEnum = "" Then
                        cgrdPerson.ColComboList(j) = ""
                    Else
                        cgrdPerson.ColComboList(j) = "|" & lstrEnum
                    End If
                End If
                
                mcol�����Ŀ.Add lobjRec("�����Ŀ���").Value, lobjRec!�����Ŀ����
            End If
            i = i + 1
            lobjRec.movenext
        Loop
        cgrdInput.Select 1, 2, 1, 2
    Else
        cgrdInput.Rows = 1
        
        Err.Raise 6666, , "�Բ��𣬸������Ա�����ϵ�����" & mstr�����Ŀ���� & "�����Ŀ���㶼�����Բ���������������Ա��ʹ�õ���������û������" & mstr�����Ŀ���� & "��Ŀ�������ҵ�����õġ����ҽʦ���á������ɲ�������Ŀ�������롰�������á�������������õ���Ŀ��"
    End If
    
    Exit Sub
errHandler:
    sfsub������ "�����沿��", "FrmInputTestResult", "subShowInputGrid", Err.Number, Err.Description, True
    Exit Sub
    Resume
End Sub

'���ܣ����ݵ�������ʽ����ı����ʾ�����Ա��Ϣ�������Ŀ������������������ʽ����ĵ�����Ż�ȡ�����Ա��Ϣ�����������У���ʾ�����Ŀ��������
Private Sub subShowSinglePerson()
    On Error GoTo errHandler
    Dim lobj��� As Object     '������
    Dim lobj��켯 As Object   '��켯�������ڸ����Թܱ��+���ڻ�ȡϵͳ��š�
    Dim lobjRec As Object
    
    Dim lstrNo As String       'ϵͳ��Ż��Թܱ�š�
    Dim llngNoType As Long     '������ͣ�0 ϵͳ���/1 �Թܱ�š�
    Dim lstrSysNo As String    'ϵͳ��š�
    Dim i As Long
    
    
    '��ȡ�����ϵͳ��ţ����Թܱ�ţ���
    lstrNo = Trim(ctxtSingleNo.Text)
    
    If lstrNo <> "" Then
        '����������
        Set lobj��� = CreateObject("������.clsMedicalExam")
        
        '��ȡ������͡�
        If coptType(0).Value Then
            llngNoType = 0
        ElseIf coptType(1).Value Then
            '�����Թܱ�Ż�ȡϵͳ��š�
            llngNoType = 1
        ElseIf coptType(2).Value Then
            '������쵥�Ż�ȡϵͳ��š�
            llngNoType = 2
        Else
            '����������ȡϵͳ���
            llngNoType = 3
        End If
        
        If llngNoType <> 0 Then
            lstrNo = lobj���.func����������Ż�ȡϵͳ���(lstrNo, llngNoType)
        End If
        
        '�������ϵͳ��š�
        lobj���.ϵͳ��� = lstrNo
        
        lstrSysNo = lobj���.ϵͳ���
        ctxtSingleNo.Text = lstrSysNo
        
        '��ս��档
        If ctabPerson.Tab = 0 Then
            clblInfo(0) = ""
            clblInfo(1) = ""
            clblInfo(2) = ""
            clblInfo(3) = ""
            clblInfo(4) = ""
            cpicPhoto(0).Picture = Nothing
        End If
        
        '�ж��Ƿ���ڡ�
        If Not lobj���.�Ƿ��Ѵ��� Then
            Err.Raise 6666, , "�������������ŵ������Ա�����������롣"
        End If
        
        '�ж��Ƿ����������ۡ�
        If lobj���.���״̬ = P_ENDED_STATUS Then
            Err.Raise 6666, , "�������ŵ�����ѱ�ҽʦȷ���������ۣ��������޸������������۵��������" & Chr(13) & Chr(10) & "��ȷʵҪ�޸ģ����½��۵�ҽʦ���롰������¼�롱����������ȡ���½��ۣ��ٻص��˲��������޸ġ�"
        End If
        
        '��ʾ��Ա��Ϣ��������Ƭ����
        If ctabPerson.Tab = 0 Then
            '��������ʽ��
            With lobj���.�����Ա
                clblInfo(0) = .����
                clblInfo(1) = .�Ա�
                clblInfo(2) = .����
                clblInfo(3) = .��λ����
                If llngNoType = 1 Then 'ϵͳ������뷽ʽ����Ҫ��ʾ�Թܱ�š�
                    clblInfo(4) = lobj���.��쵥��
                    Label1(8).Caption = "��쵥�ţ�"
                Else
                    clblInfo(4) = lobj���.�Թܱ��
                    Label1(8).Caption = "�Թܱ�ţ�"
                End If
                
                '��ʾ��Ƭ��
                If Not .��Ƭ Is Nothing Then
                    cpicPhoto(0).Picture = .��Ƭ
                End If
            End With
            
            '���������¼������
            subShowInputGrid lstrSysNo
            
            cgrdSingleList.Row = 0
            
        Else
            '�޸ģ�2001-11-2�������Ϊֻ��ѯָ�����������¼�����Բ����ж������Ƿ���ͬ��
            
            '��������ʽ������Ա��Ϣ���뵽cgrdPerson�У�ע������������ͬ�ģ���
            If cgrdPerson.Rows = 1 Then
'                '��cgrdPerson��ԭû�м�¼������mstr�������ơ�
'                mstr�������� = lobj���.����.������
'
'                '���������¼������
                subShowInputGrid lstrSysNo
            Else
'                '�ж������Ա���������Ƿ�һ�¡�
                '�޸ģ�2002-8-14������������ѡ�����С�
                If ccmbSheet.Text <> "<����>" Then
                    If ccmbSheet.Text <> lobj���.����.������ Then
                        Err.Raise 6666, , "����������������" & lobj���.����.������ & "����ָ������һ�£���������¼��������ͬ���������"
                    End If
                End If
            End If

            '�жϸ���Ա�Ƿ����������У�����������Լ�������
            For i = 1 To cgrdPerson.Rows - 1
                If cgrdPerson.TextMatrix(i, 0) = lstrSysNo Then
                    '���������д��ڣ����ټ��롣
                    Exit Sub
                End If
            Next
            
            '����Ա��ӵ������Ա�����С�
            cgrdPerson.Rows = cgrdPerson.Rows + 1
            
            i = cgrdPerson.Rows - 1
            cgrdPerson.TextMatrix(i, 0) = lstrSysNo
            
            '�޸ģ�2002-10-11�����������ʾ�Թܱ�š�
            cgrdPerson.TextMatrix(i, 1) = lobj���.�Թܱ��
            With lobj���.�����Ա
                cgrdPerson.TextMatrix(i, 2) = .����
                cgrdPerson.TextMatrix(i, 3) = .�Ա�
                cgrdPerson.TextMatrix(i, 4) = .��λ����
                cgrdPerson.TextMatrix(i, 5) = .����
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
    sfsub������ "�����沿��", "FrmInputTestResult", "subShowSinglePerson", Err.Number, Err.Description, True
    Exit Sub
    Resume
End Sub

'���ܣ���һ����ָ����ŷ�Χ����������ڣ��������Ա�������񣬲���ʾ�����Ŀ��������
Private Sub subShowBatchPerson()
    Dim lobjRec As Object        'ͨ��ҵ������ȡ��ָ����Χ�ڿ���¼�������������¼��
    Dim llngNoType  As Integer   '��ŷ�ʽ��0ϵͳ���/1�Թܱ�š�
    Dim llngStartRow As Long     '��ǰ�����Ա����������+1��
    Dim llngRow As Long          '��ǰ��ӵ��С�
    Dim i As Long
    Dim lobjResult As Object
    
    On Error GoTo errHandler
    cstbMain.Panels(1) = "���ڻ�ȡ����¼�����Ժ�..."
    
    '��ȡ������͡�
    llngNoType = 0 'ϵͳ��š�
    
    
    '�������������ڡ�
    '�޸ģ�2001-11-2�����Ӳ�ѯ�������������ƣ���
    '�޸ģ�2002-8-14������������ѡ�����С�
    Set lobjRec = pobjҵ�����.Func��ȡ���޸ĵ�����¼(IIf(coptBatchType(2).Value, ctxtBatchNo, ""), IIf(coptBatchType(2).Value, ctxtBatchNo, ""), IIf(coptBatchType(0).Value, Str(cdtpQueryDate.Value), ""), llngNoType, IIf(ccmbSheet.Text = "<����>", "", ccmbSheet.Text), IIf(coptBatchType(1).Value, ctxt��λ����.Text, ""))

    If lobjRec.recordcount > 0 Then
   
        lobjRec.Filter = ""
        If cchkUnEnd(1).Value = 1 And cchkEnd(1).Value = 0 Then
            lobjRec.Filter = "���״̬<>2"
        ElseIf cchkUnEnd(1).Value = 0 And cchkEnd(1).Value = 1 Then
            lobjRec.Filter = "���״̬=2"
        ElseIf cchkUnEnd(1).Value = 0 And cchkEnd(1).Value = 0 Then
            lobjRec.Filter = "���״̬=-1"
        End If
        
        cgrdPerson.Redraw = False
        mblnSys = True
        mintFixed = 6
        If cgrdPerson.Rows = 1 And lobjRec.recordcount > 0 Then
            '�޸ģ�2001-11-2���������Ҫ�ж��������ơ�
            '���������¼�������С�
            subShowInputGrid lobjRec!ϵͳ���
        End If
        
        '��ʾ��Ա��Ϣ��cgrdPerson�У�ע������������ͬ�ģ���
        llngStartRow = cgrdPerson.Rows - 1
        Do While Not lobjRec.EOF
            '�޸ģ�2001-11-2���������Ҫ�ж��������ơ�
    '        If lobjRec!�������� = mstr�������� Then
                '�жϸ���Ա�Ƿ����������У�����������Լ�������
                For i = 1 To llngStartRow
                    If cgrdPerson.TextMatrix(i, 0) = lobjRec!ϵͳ��� Then
                        '���������д��ڣ����ټ��롣
                        GoTo LabelContinue
                    End If
                Next
                cgrdPerson.AddItem ""
                llngRow = cgrdPerson.Rows - 1
                With cgrdPerson
                    .TextMatrix(llngRow, 0) = lobjRec!ϵͳ���
                    
                    '�޸ģ�2002-10-11�����������ʾ�Թܱ�š�
                    .TextMatrix(llngRow, 1) = IIf(IsNull(lobjRec!�Թܱ��), "", lobjRec!�Թܱ��)
                    .TextMatrix(llngRow, 2) = IIf(IsNull(lobjRec!����), "", lobjRec!����)
                    .TextMatrix(llngRow, 3) = IIf(IsNull(lobjRec!�Ա�), "", lobjRec!�Ա�)
                    .TextMatrix(llngRow, 4) = IIf(IsNull(lobjRec!��λ����), "", lobjRec!��λ����)
                    .TextMatrix(llngRow, 5) = IIf(IsNull(lobjRec!����), "", lobjRec!����)
                    .TextMatrix(llngRow, 6) = IIf(IsNull(lobjRec!��쵥��), "", lobjRec!��쵥��)
                    
                    If lobjRec!���״̬ = 2 Then
                        .Cell(flexcpBackColor, llngRow, 0, llngRow, mintFixed) = cchkEnd(1).BackColor
                    Else
                        .Cell(flexcpBackColor, llngRow, 0, llngRow, mintFixed) = cchkUnEnd(1).BackColor
                    End If
                    
                    '2006-6-19(����¼�룩
                    'If cchkGrid.Value = 1 Then
                        '��ȡ���˵������������
                        subShowPersonResult llngRow, lobjRec!ϵͳ���
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
    sfsub������ "�����沿��", "FrmInputTestResult", "subShowBatchPerson", Err.Number, Err.Description, True
    mblnSys = False
    Exit Sub
    Resume
End Sub

Private Sub subShowPersonResult(ByVal paraRow As Long, ByVal paraϵͳ��� As String)
    Dim i As Long
    Dim lobjResult As Object
    
    
    Set lobjResult = pobjҵ�����.func��ȡ�����(paraϵͳ���)
    Do While Not lobjResult.EOF
        For i = mintFixed + 1 To cgrdPerson.Cols - 1
            If cgrdPerson.TextMatrix(0, i) = lobjResult!�����Ŀ���� Then
                cgrdPerson.TextMatrix(paraRow, i) = IIf(IIf(IsNull(lobjResult!�����), "", lobjResult!�����) = "", lobjResult!ȱʡֵ, lobjResult!�����)
                '������ɫ��
                If IIf(IsNull(lobjResult!�������), "", lobjResult!�������) = "���ϸ�" Then
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


Private Sub mobj����ͨ�ö���_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    On Error GoTo errHandler
    Dim lcolNo As Collection     'ϵͳ��ż��ϡ�
    Dim lcolResult As Collection '��������ϣ�item:[�����Ŀ�������]��
    Dim lcolItem As Collection   '���������Ŀ���������[�����Ŀ�������]��
    Dim lcolDetail As Collection
    Dim lblnNotOver As Boolean
    Dim i As Long
    Dim j As Long
    
    Select Case Operate
    Case "����"
        MousePointer = 11
        cstbMain.Panels(1) = "���ڱ��棬���Ժ�..."
        
        '��ʱ���治�ܲ�����
        ctbMain.Enabled = False
        ctabPerson.Enabled = False
        Frame1.Enabled = False
        cgrdInput.Select 1, 0, 1, 0
        
        Set lcolNo = New Collection
        Set lcolResult = New Collection
        
        '��ȡ��������ϣ�[�����Ŀ�������]��
        lblnNotOver = False
                
        If ctabPerson.Tab = 0 Then
            '���ǵ���¼�뷽ʽ����ctxtSingleNo.text��Ӧ��ϵͳ��ż��뼯��lcolNo�У�
            '�������ϵͳ��š�
            lcolNo.Add ctxtSingleNo.Text
        Else
            '��������¼�뷽ʽ����cgrdPerson�����������е�ϵͳ��ż���lcolNo�С�
            For i = 1 To cgrdPerson.Rows - 1
                lcolNo.Add cgrdPerson.TextMatrix(i, 0)
                
                '2006-6-19(����¼��)
                Set lcolDetail = New Collection
                If cchkGrid.Value = 1 Then
                   
                    For j = mintFixed + 1 To cgrdPerson.Cols - 1
                        Set lcolItem = New Collection
                        lcolItem.Add mcol�����Ŀ(cgrdPerson.TextMatrix(0, j)), "�����Ŀ"
                        lcolItem.Add cgrdPerson.TextMatrix(i, j), "�����"
                        
                        If cgrdPerson.TextMatrix(i, j) = "" Then
                            lblnNotOver = True
                            lcolItem.Add "", "�������"
                        ElseIf cgrdPerson.Cell(flexcpBackColor, i, j, i, j) = &H8A5AFA Then
                            lcolItem.Add "���ϸ�", "�������"
                        Else
                            lcolItem.Add "�ϸ�", "�������"
                        End If
                        
                        lcolDetail.Add lcolItem, lcolItem("�����Ŀ")
                    Next
                    lcolResult.Add lcolDetail, cgrdPerson.TextMatrix(i, 0)
                End If
            Next
        End If
        
        If lcolNo.Count = 0 Then
            Err.Raise 6666, , "��ѡ�������Ա����¼����������ٰ������桱��"
        End If
        
        If ctabPerson.Tab = 0 Or cchkGrid.Value = 0 Then
            
            For i = 1 To cgrdInput.Rows - 1
                Set lcolItem = New Collection
                lcolItem.Add cgrdInput.TextMatrix(i, 0), "�����Ŀ"
                lcolItem.Add cgrdInput.TextMatrix(i, 2), "�����"
                
                '��¼û��¼�ꡣ
                If cgrdInput.TextMatrix(i, 2) = "" Then
                    lblnNotOver = True
                    lcolItem.Add "", "�������"
                ElseIf cgrdInput.Cell(flexcpBackColor, i, 2, i, 2) = &H8A5AFA Then
                    lcolItem.Add "���ϸ�", "�������"
                Else
                    lcolItem.Add "�ϸ�", "�������"
                End If
                lcolResult.Add lcolItem, lcolItem("�����Ŀ")
            Next
        End If
        
        '��û��¼�꣬������ʾ��
        If lblnNotOver Then
            If Not sffuncMsg("��û��¼�����������Ŀ����������Ƿ���Ҫ���棿", sfѯ��) Then
                '�û�ѡ�񲻱��档
                GoTo errHandler
            End If
        End If
        
        'ʹ���Ż����㷨�����������
        If ctabPerson.Tab = 0 Or cchkGrid.Value = 0 Then
            pobjҵ�����.Sub�Ż�����д����� lcolNo, lcolResult, um�û����, cdtpInputDate.Value
        Else
            pobjҵ�����.Sub������д����� lcolNo, lcolResult, um�û����, cdtpInputDate.Value
        End If
        '�ָ����档
        ctbMain.Buttons(1).Enabled = False
        ctbMain.Enabled = True
        ctabPerson.Enabled = True
        Frame1.Enabled = True
        
        '��ս��档
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
        cstbMain.Panels(1) = "����ɹ���"
        Cancel = True
    End Select
    Exit Sub
    
errHandler:
    If Err.Number <> 0 Then
        Dim lstrError As String
        lstrError = func������(Err.Number, Err.Description)
        sfsub������ "�����沿��", "FrmInputTestResult", "mobj����ͨ�ö���_BeforeOperate", 6666, lstrError, False
    End If
    If Operate = "����" Then
        '�ָ�������Բ�����
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

'���ܣ���һ����ָ����ŷ�Χ����������ڣ��������Ա�������񣬲���ʾ�����Ŀ��������
Private Sub subShowSingleList()
    Dim lobjRec As Object        'ͨ��ҵ������ȡ��ָ����Χ�ڿ���¼�������������¼��
    Dim llngStartRow As Long     '��ǰ�����Ա����������+1��
    Dim llngRow As Long          '��ǰ��ӵ��С�
    Dim i As Long
    
    On Error GoTo errHandler
    cstbMain.Panels(1) = "���ڻ�ȡ����¼�����Ժ�..."
    
    cgrdSingleList.Rows = 1
    
    '�������������ڡ�
    Set lobjRec = pobjҵ�����.Func��ȡ���޸ĵ�����¼("", "", Str(cdtpSingleQuery.Value), 0, "")
    If lobjRec.recordcount = 0 Then
        cstbMain.Panels(1) = ""
        Exit Sub
    End If
    lobjRec.Filter = ""
    If cchkUnEnd(0).Value = 1 And cchkEnd(0).Value = 0 Then
        lobjRec.Filter = "���״̬<>2"
    ElseIf cchkUnEnd(0).Value = 0 And cchkEnd(0).Value = 1 Then
        lobjRec.Filter = "���״̬=2"
    ElseIf cchkUnEnd(0).Value = 0 And cchkEnd(0).Value = 0 Then
        lobjRec.Filter = "���״̬=-1"
    End If
    cgrdSingleList.Redraw = False
    mblnSys = True
    
    '��ʾ��Ա��Ϣ��cgrdSingleList�У�ע������������ͬ�ģ���
    cgrdSingleList.Rows = 1
    
    llngStartRow = cgrdSingleList.Rows - 1
    Do While Not lobjRec.EOF
        cgrdSingleList.AddItem ""
        llngRow = cgrdSingleList.Rows - 1
        With cgrdSingleList
            .TextMatrix(llngRow, 0) = lobjRec!ϵͳ���
            '�޸ģ�2002-10-11�����������ʾ�Թܱ�š�
            .TextMatrix(llngRow, 1) = IIf(IsNull(lobjRec!�Թܱ��), "", lobjRec!�Թܱ��)
            .TextMatrix(llngRow, 2) = IIf(IsNull(lobjRec!����), "", lobjRec!����)
            .TextMatrix(llngRow, 3) = IIf(IsNull(lobjRec!�Ա�), "", lobjRec!�Ա�)
            .TextMatrix(llngRow, 4) = IIf(IsNull(lobjRec!��λ����), "", lobjRec!��λ����)
            .TextMatrix(llngRow, 5) = IIf(IsNull(lobjRec!����), "", lobjRec!����)
            .TextMatrix(llngRow, 6) = IIf(IsNull(lobjRec!��쵥��), "", lobjRec!��쵥��)
            If lobjRec!���״̬ = 2 Then
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
    sfsub������ "�����沿��", "FrmInputTestResult", "subShowSingleList", Err.Number, Err.Description, True
    Exit Sub
    Resume
End Sub


