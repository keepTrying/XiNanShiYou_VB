VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{2327075A-FED2-4AEF-82B8-DD0C2B1AC8E1}#2.0#0"; "dyCatchPhoto.ocx"
Object = "{F48D4DEC-1198-11D5-91BE-0050BA06B70C}#5.9#0"; "¼��ؼ�.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmRegister 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ְҵ�������-���Ǽ�"
   ClientHeight    =   8505
   ClientLeft      =   240
   ClientTop       =   375
   ClientWidth     =   11520
   ClipControls    =   0   'False
   Icon            =   "FrmRegister.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8938.523
   ScaleMode       =   0  'User
   ScaleWidth      =   11520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   900
      Left            =   5640
      Top             =   600
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7335
      Left            =   240
      TabIndex        =   22
      Top             =   840
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   12938
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "������Ϣ¼��            "
      TabPicture(0)   =   "FrmRegister.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2(3)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2(2)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label2(6)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label8"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label2(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label3"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label12"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label30"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label33"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label34"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "clblHintCheck"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "clblHistory"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "ctlInputDictGrid1"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cdtpDate"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "ccmb������"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "ccmb�������"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "ccmbTemplate"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "ctxtName"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "ctxtAge"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "ctxt���֤��"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "clblsysno"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "frmPhoto"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "cdtp����"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "ccmb���������"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Ccmb��������"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Picture2"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "cgrdHistory"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "ccmbSex"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).ControlCount=   32
      TabCaption(1)   =   "������Ϣ¼��           "
      TabPicture(1)   =   "FrmRegister.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(2)=   "Frame1"
      Tab(1).ControlCount=   3
      Begin VB.ComboBox ccmbSex 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   300
         Left            =   360
         TabIndex        =   107
         Text            =   "Combo1"
         Top             =   3720
         Width           =   855
      End
      Begin VSFlex8Ctl.VSFlexGrid cgrdHistory 
         Height          =   2535
         Left            =   360
         TabIndex        =   105
         Top             =   4560
         Visible         =   0   'False
         Width           =   5415
         _cx             =   2088772943
         _cy             =   2088767863
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
      Begin VB.PictureBox Picture2 
         Height          =   1935
         Left            =   4200
         ScaleHeight     =   1875
         ScaleWidth      =   1515
         TabIndex        =   100
         Top             =   2040
         Width           =   1575
      End
      Begin VB.ComboBox Ccmb�������� 
         Height          =   300
         Left            =   4800
         TabIndex        =   96
         Top             =   1320
         Width           =   1935
      End
      Begin VB.ComboBox ccmb��������� 
         Height          =   300
         Left            =   2880
         TabIndex        =   95
         Top             =   1320
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker cdtp���� 
         Height          =   300
         Left            =   2400
         TabIndex        =   92
         Top             =   3720
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   59572224
         CurrentDate     =   40960
      End
      Begin VB.Frame Frame3 
         Caption         =   "  �ֵ�λ��Ϣ¼��   "
         ForeColor       =   &H000080FF&
         Height          =   2415
         Left            =   -74520
         TabIndex        =   76
         Top             =   4800
         Width           =   9735
         Begin VB.ComboBox Ccmb��ҵ��� 
            Height          =   300
            Left            =   6720
            TabIndex        =   90
            Text            =   "Ccmb��ҵ���"
            Top             =   1200
            Width           =   1815
         End
         Begin VB.ComboBox ccmb�������� 
            Height          =   300
            Left            =   4560
            TabIndex        =   88
            Text            =   "ccmb��������"
            Top             =   1200
            Width           =   1695
         End
         Begin VB.TextBox ctxt��λ��ַ 
            Height          =   300
            Left            =   480
            TabIndex        =   86
            Top             =   1920
            Width           =   5775
         End
         Begin VB.TextBox ctxt��ϵ�绰 
            Height          =   300
            Left            =   2520
            TabIndex        =   84
            Top             =   1200
            Width           =   1575
         End
         Begin VB.TextBox ctxt������ 
            Height          =   300
            Left            =   480
            TabIndex        =   82
            Top             =   1200
            Width           =   1695
         End
         Begin VB.CheckBox cchk¼�뵥λ���� 
            Caption         =   "¼�뵥λ����"
            Height          =   255
            Left            =   6600
            TabIndex        =   79
            Top             =   240
            Value           =   1  'Checked
            Width           =   1575
         End
         Begin VB.CommandButton ccmd��λ��λ 
            Caption         =   "��λ(&T)"
            Height          =   375
            Left            =   4080
            Style           =   1  'Graphical
            TabIndex        =   78
            TabStop         =   0   'False
            Top             =   360
            Width           =   945
         End
         Begin VB.ComboBox ccmbUnit 
            Height          =   300
            Left            =   480
            TabIndex        =   77
            Top             =   480
            Width           =   3480
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "��ҵ���"
            Height          =   180
            Left            =   6720
            TabIndex        =   89
            Top             =   960
            Width           =   900
         End
         Begin VB.Label Label28 
            Caption         =   "�������ʣ�"
            Height          =   255
            Left            =   4560
            TabIndex        =   87
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "��λ��ַ��"
            Height          =   180
            Left            =   480
            TabIndex        =   85
            Top             =   1680
            Width           =   900
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "��ϵ�绰��"
            Height          =   180
            Left            =   2520
            TabIndex        =   83
            Top             =   960
            Width           =   900
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "�����ˣ�"
            Height          =   180
            Left            =   480
            TabIndex        =   81
            Top             =   960
            Width           =   720
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��λ���ƣ�"
            Height          =   180
            Index           =   5
            Left            =   480
            TabIndex        =   80
            Top             =   240
            Width           =   900
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "  Σ��������Ϣ¼��    "
         ForeColor       =   &H000080FF&
         Height          =   1815
         Left            =   -74520
         TabIndex        =   61
         Top             =   2880
         Width           =   9735
         Begin VB.ComboBox ccmbΣ������ 
            Height          =   300
            Left            =   4560
            TabIndex        =   68
            Top             =   480
            Width           =   1815
         End
         Begin VB.ComboBox ccmbְҵ��� 
            Height          =   300
            Left            =   2640
            TabIndex        =   67
            Top             =   480
            Width           =   1575
         End
         Begin VB.ComboBox ccmb����Դ 
            Height          =   300
            Left            =   480
            TabIndex        =   66
            Top             =   480
            Width           =   1815
         End
         Begin VB.ComboBox ccmb�ֹ��� 
            Height          =   300
            Left            =   480
            TabIndex        =   65
            Top             =   1200
            Width           =   1815
         End
         Begin VB.ComboBox ccmbְ�� 
            Height          =   300
            Left            =   2640
            TabIndex        =   64
            Top             =   1200
            Width           =   1575
         End
         Begin VB.TextBox ctxt������� 
            Height          =   270
            Left            =   6720
            TabIndex        =   63
            Top             =   1200
            Width           =   1815
         End
         Begin VB.TextBox ctxtΣ������ 
            Height          =   270
            Left            =   4560
            TabIndex        =   62
            Top             =   1200
            Width           =   1815
         End
         Begin VB.Label Label17 
            Caption         =   "Σ�����أ�"
            Height          =   255
            Left            =   4560
            TabIndex        =   75
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "ְҵ���"
            Height          =   180
            Left            =   2640
            TabIndex        =   74
            Top             =   240
            Width           =   900
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "����Դ��"
            Height          =   180
            Left            =   480
            TabIndex        =   73
            Top             =   240
            Width           =   720
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "�ֹ��֣�"
            Height          =   180
            Left            =   480
            TabIndex        =   72
            Top             =   960
            Width           =   720
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "ְҵ/ְ�ƣ�"
            Height          =   180
            Left            =   2640
            TabIndex        =   71
            Top             =   960
            Width           =   990
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "���������"
            Height          =   180
            Left            =   6720
            TabIndex        =   70
            Top             =   960
            Width           =   900
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "ְҵΣ�����䣺"
            Height          =   180
            Left            =   4560
            TabIndex        =   69
            Top             =   960
            Width           =   1260
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "  ���˸�����Ϣ      "
         ForeColor       =   &H000080FF&
         Height          =   2295
         Left            =   -74520
         TabIndex        =   46
         Top             =   480
         Width           =   9735
         Begin VB.TextBox ctxt������ 
            Height          =   300
            Left            =   1200
            TabIndex        =   97
            Top             =   1750
            Width           =   4455
         End
         Begin VB.ComboBox ccmb���� 
            Height          =   300
            Left            =   3840
            TabIndex        =   94
            Text            =   "ccmb����"
            Top             =   1320
            Width           =   1215
         End
         Begin VB.ComboBox Ccmb��� 
            Height          =   300
            Left            =   2640
            TabIndex        =   56
            Text            =   "Ccmb���"
            Top             =   1320
            Width           =   855
         End
         Begin VB.TextBox ctxt�绰 
            Height          =   300
            Left            =   5400
            TabIndex        =   55
            Top             =   1320
            Width           =   1815
         End
         Begin VB.TextBox ctxt���� 
            Height          =   270
            Left            =   7920
            TabIndex        =   54
            Top             =   1320
            Width           =   975
         End
         Begin VB.ComboBox ccmb�Ļ��̶� 
            Height          =   300
            Left            =   480
            TabIndex        =   53
            Text            =   "ccmb�Ļ��̶�"
            Top             =   1320
            Width           =   1815
         End
         Begin VB.TextBox ctxtסַ 
            Height          =   300
            Left            =   4560
            TabIndex        =   50
            Top             =   600
            Width           =   4215
         End
         Begin VB.TextBox ctxt�ʱ� 
            Height          =   300
            Left            =   2640
            TabIndex        =   49
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox ctxt���� 
            Height          =   300
            Left            =   480
            TabIndex        =   47
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "�����أ�"
            Height          =   180
            Left            =   480
            TabIndex        =   98
            Top             =   1800
            Width           =   720
         End
         Begin VB.Label label80 
            AutoSize        =   -1  'True
            Caption         =   "���壺"
            Height          =   180
            Left            =   3840
            TabIndex        =   93
            Top             =   1080
            Width           =   540
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "���"
            Height          =   180
            Left            =   2640
            TabIndex        =   60
            Top             =   1080
            Width           =   540
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "�Ļ��̶ȣ�"
            Height          =   180
            Left            =   480
            TabIndex        =   59
            Top             =   1080
            Width           =   900
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "�绰���룺"
            Height          =   180
            Left            =   5400
            TabIndex        =   58
            Top             =   1080
            Width           =   900
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "���䣺"
            Height          =   180
            Left            =   7920
            TabIndex        =   57
            Top             =   1080
            Width           =   540
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "סַ��"
            Height          =   180
            Left            =   4560
            TabIndex        =   52
            Top             =   360
            Width           =   540
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "������ţ�"
            Height          =   180
            Left            =   2640
            TabIndex        =   51
            Top             =   360
            Width           =   900
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "���᣺"
            Height          =   180
            Left            =   480
            TabIndex        =   48
            Top             =   360
            Width           =   540
         End
      End
      Begin VB.Frame frmPhoto 
         Caption         =   "����"
         ClipControls    =   0   'False
         ForeColor       =   &H00800000&
         Height          =   4275
         Left            =   5880
         TabIndex        =   42
         Top             =   1800
         Width           =   4905
         Begin VB.CommandButton ccmdTakePhotoAgain 
            Caption         =   "����ȡ��"
            Height          =   495
            Left            =   3000
            TabIndex        =   103
            Top             =   2880
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00E0E0E0&
            Height          =   1935
            Left            =   1320
            ScaleHeight     =   1875
            ScaleWidth      =   1515
            TabIndex        =   43
            Top             =   480
            Width           =   1575
         End
         Begin dyCatchPhoto.ctlCatchPhoto cctlCatchPhoto 
            Height          =   3570
            Left            =   120
            TabIndex        =   44
            Top             =   360
            Visible         =   0   'False
            Width           =   4485
            _ExtentX        =   8017
            _ExtentY        =   6297
            BackColor       =   0
            FontSize        =   9.75
            OriginalSize    =   -1  'True
         End
      End
      Begin VB.TextBox clblsysno 
         Height          =   270
         Left            =   360
         TabIndex        =   38
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox ctxt���֤�� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   360
         TabIndex        =   36
         Top             =   3000
         Width           =   2130
      End
      Begin VB.TextBox ctxtAge 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1440
         TabIndex        =   34
         Top             =   3720
         Width           =   495
      End
      Begin VB.TextBox ctxtName 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   300
         Left            =   2640
         TabIndex        =   31
         Top             =   3000
         Width           =   1410
      End
      Begin VB.ComboBox ccmbTemplate 
         Height          =   300
         ItemData        =   "FrmRegister.frx":047A
         Left            =   360
         List            =   "FrmRegister.frx":047C
         TabIndex        =   27
         Top             =   2040
         Width           =   3480
      End
      Begin VB.ComboBox ccmb������� 
         Height          =   300
         Left            =   7080
         TabIndex        =   25
         Top             =   1320
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.ComboBox ccmb������ 
         Height          =   300
         ItemData        =   "FrmRegister.frx":047E
         Left            =   6000
         List            =   "FrmRegister.frx":0485
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   960
         Visible         =   0   'False
         Width           =   2175
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
         TabIndex        =   29
         Top             =   960
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
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
         Format          =   59572224
         CurrentDate     =   36950
         MaxDate         =   73050
         MinDate         =   17899
      End
      Begin ¼��ؼ�.ctlInputDictGrid ctlInputDictGrid1 
         Height          =   2535
         Left            =   8280
         TabIndex        =   41
         Top             =   6480
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   4471
         Cols            =   10
         Count           =   0
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
      Begin VB.Label clblHistory 
         Caption         =   "˫���У�������������Ϣ�͸�����Ϣ��"
         Height          =   255
         Left            =   360
         TabIndex        =   106
         Top             =   4320
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.Label clblHintCheck 
         Caption         =   "ע�⣺У��֮��ֻ�������࣬�������ݼ�ʹ�޸ģ�Ҳ���ᱣ�档"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3480
         TabIndex        =   104
         Top             =   720
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "�뽫�������֤���ڶ������ϣ�"
         Height          =   180
         Left            =   360
         TabIndex        =   102
         Top             =   2520
         Width           =   2520
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "ע����ˢ���룬��ˢ���֤"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   7680
         TabIndex        =   99
         Top             =   6240
         Visible         =   0   'False
         Width           =   2160
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "�������ڣ�"
         Height          =   180
         Left            =   2400
         TabIndex        =   91
         Top             =   3480
         Width           =   900
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C000&
         Caption         =   "�ǿ���¼��ʱ��ɫΪ��¼�����¼��ʱֻ��ˢ�������֤"
         Height          =   180
         Left            =   360
         TabIndex        =   45
         Top             =   480
         Visible         =   0   'False
         Width           =   4680
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "ע��ˢ����ǰ��ȷ���ı���������Ϊ��"
         Height          =   180
         Left            =   360
         TabIndex        =   40
         Top             =   720
         Visible         =   0   'False
         Width           =   3060
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ϵͳ��ţ�"
         Height          =   180
         Index           =   0
         Left            =   360
         TabIndex        =   39
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "���֤�ţ�"
         Height          =   180
         Left            =   360
         TabIndex        =   37
         Top             =   2760
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���䣺"
         Height          =   180
         Index           =   6
         Left            =   1440
         TabIndex        =   35
         Top             =   3480
         Width           =   540
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         Height          =   180
         Left            =   360
         TabIndex        =   33
         Top             =   3480
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   2640
         TabIndex        =   32
         Top             =   2760
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������ڣ�"
         Height          =   180
         Index           =   2
         Left            =   8760
         TabIndex        =   30
         Top             =   720
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   360
         TabIndex        =   28
         Top             =   1800
         Width           =   720
      End
      Begin VB.Label Label4 
         Caption         =   "�����Ա���ͣ�"
         Height          =   255
         Left            =   2880
         TabIndex        =   26
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Index           =   3
         Left            =   4800
         TabIndex        =   24
         Top             =   1080
         Visible         =   0   'False
         Width           =   900
      End
   End
   Begin VB.CheckBox Check���֤ 
      Caption         =   "ˢ�������֤"
      Height          =   255
      Left            =   8520
      TabIndex        =   17
      Top             =   480
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.CheckBox cchkClear 
      Caption         =   "��������"
      Height          =   345
      Left            =   8520
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   6600
      Top             =   360
   End
   Begin VB.Frame cfram������Ϣ 
      Caption         =   "�Ǽǻ�����Ϣ���ǿ���¼��ʱ��ɫΪ��¼�����¼��ʱֻ������):"
      ClipControls    =   0   'False
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   2520
      TabIndex        =   5
      Top             =   9360
      Width           =   6300
      Begin VB.TextBox ctxt���� 
         Height          =   300
         Left            =   4800
         TabIndex        =   21
         Top             =   2040
         Width           =   1695
      End
      Begin VB.ComboBox ccmb���ʱ�� 
         Height          =   300
         Left            =   8160
         TabIndex        =   19
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox ctxt���� 
         Height          =   270
         Left            =   4440
         TabIndex        =   15
         Text            =   "1"
         Top             =   2880
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox ctxt��쵥�� 
         Height          =   315
         Left            =   4200
         MaxLength       =   20
         TabIndex        =   1
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox ctxtTubeNo 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6000
         TabIndex        =   0
         Top             =   1320
         Width           =   1575
      End
      Begin VB.VScrollBar cvscLetter 
         Height          =   345
         Left            =   6360
         TabIndex        =   3
         Top             =   1320
         Width           =   345
      End
      Begin ¼��ؼ�.ctlInputDictGrid c�ֵ�� 
         Height          =   3255
         Left            =   360
         TabIndex        =   9
         Top             =   240
         Visible         =   0   'False
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   5741
         Cols            =   10
         Count           =   0
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
      Begin ¼��ؼ�.ctlInputFrame ciptBase 
         Height          =   975
         Left            =   6120
         TabIndex        =   2
         Top             =   2280
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   1720
         BackColor       =   15791081
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackStyle       =   1
         Caption         =   ""
         Rows            =   1
         Cols            =   27
         DistanceofRow   =   0
         AutoSize        =   0   'False
         FormatString    =   "���֤��,1,0,12"
         Count           =   1
         titleInputBox0001=   "���֤��"
         statusinfoInputBox0001=   ""
         lengthInputBox0001=   12
         orderInputBox0001=   1
         valueInputBox0001=   ""
         datatypeInputBox0001=   3
         colInputBox0001 =   0
         rowInputBox0001 =   1
         PassWordCharInputBox0001=   0   'False
         ����InputBox0001=   0   'False
         ����������ֵInputBox0001=   0   'False
         ���������СֵInputBox0001=   0   'False
         �ֵ�����InputBox0001=   ""
         ��ʾ�ֵ��ֶ�InputBox0001=   ""
         �����ֵ��ֶ�InputBox0001=   ""
         ����InputBox0001=   "����� 1"
         ȱʡֵInputBox0001=   ""
         ����ȱʡֵInputBox0001=   ""
         ����InputBox0001=   0
         MaxInputBox0001 =   ""
         MinInputBox0001 =   ""
         VisibleInputBox0001=   -1  'True
         PermitNullInputBox0001=   -1  'True
         TriggerstrInputBox0001=   ""
         �����ѡInputBox0001=   0   'False
         ErrColor        =   12648447
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "�뽫�������֤���ڶ������ϣ�"
         Height          =   180
         Left            =   0
         TabIndex        =   101
         Top             =   0
         Width           =   2520
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "���壺"
         Height          =   180
         Left            =   4800
         TabIndex        =   20
         Top             =   1800
         Width           =   540
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "���ʱ�ڣ�"
         Height          =   180
         Left            =   8160
         TabIndex        =   18
         Top             =   480
         Width           =   900
      End
      Begin VB.Label clbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   3600
         TabIndex        =   14
         Top             =   2880
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��쵥�ţ�"
         Height          =   180
         Index           =   7
         Left            =   4200
         TabIndex        =   12
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label clbl��������� 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   8760
         TabIndex        =   11
         Top             =   2760
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ϴ�������ڣ�"
         Height          =   180
         Index           =   4
         Left            =   8640
         TabIndex        =   10
         Top             =   2520
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Label clblTubeNo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "������뿴״̬��"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   6600
         TabIndex        =   8
         Top             =   1320
         Width           =   915
      End
      Begin VB.Label clblLetter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   6000
         TabIndex        =   7
         Top             =   1320
         Width           =   345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Թܱ�ţ�"
         Height          =   180
         Index           =   1
         Left            =   6000
         TabIndex        =   6
         Top             =   1080
         Width           =   900
      End
   End
   Begin MSComctlLib.Toolbar ctbMain 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   11520
      _ExtentX        =   20320
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "cimg��ťͼ��"
      _Version        =   393216
      Begin MSComctlLib.ImageList cimg��ťͼ�� 
         Left            =   0
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   393216
      End
   End
   Begin MSComctlLib.StatusBar cstbMain 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   8130
      Width           =   11520
      _ExtentX        =   20320
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   20267
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
   Begin MSComDlg.CommonDialog ccmdFile 
      Left            =   0
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************************
'���ܣ�ְҵ�����Ǽǽ�����ƣ�
'      �������֤���ݶ�ȡ
'      �����Ͽ�ѡ��
'����: ������
'ʱ�䣺2012-03
'**********************************************************************

Public pstrϵͳ��� As String
'2012-08-18 �ڵ�� ��
'���Ӹ�����ر���
Public pstr����ϵͳ��� As String
'2012-08-18 �ڵ�� ��

'2012-06-15 �ڵ�� ��
'��ӹ�����������¼��ǰ�����Ա״̬
Public mintState As Integer '0��ʾ��δ���棻1��ʾУ��ͨ����2��ʾ�޸ĺ��ѱ���
'2012-06-15 �ڵ�� ��

Private mobj����� As Object                   '�����Ա������ε���졣
Private mobj��� As Object                     '��ְҵ�������ṩ��ȡϵͳ��ź��Թܱ�ţ�����Ǽ���Ϣ�ķ�����
Private mobj��켯 As Object                   '��켯��������λ��Ҫ���������Ա��Ϣ��
Private mobj����ģ�� As Object               '����ģ�壬��ȡ���еķǸ�������ģ�����ơ�
Private WithEvents mobjGUI As cls����ͨ�ö���  '����ͨ�ö���������ʼ��������������¼���ؼ���
Attribute mobjGUI.VB_VarHelpID = -1
 Public pblnOk As Boolean
Public selectedDeptName As Collection
'ҵ�����á�
Private mblnTakePhoto As Boolean               'ҵ�����á��Ƿ����࡯��
Private mbln����¼�� As Boolean

Private mcolTubeNo As New Collection           '��ǰ�����ѡ���Թ���ĸ��

Private mstr��λ������ As String             '��λ��λ�������š�
Private mblnInUse As Boolean

'��ѡ��������Ŀ���շ���Ŀ
Private mcol�����Ŀ As New Collection
Private mcol�շ���Ŀ As New Collection               'item:���,key����š�

Public pstrϵͳ������� As String

Private mobj����  As cls�û���������
Private mstrĬ������ As String


'���ܣ����ص�ǰ�����Ƿ��Ѽ��أ��Ա������������жϵ�ǰ�����Ƿ���ִ�й�Form_Load��
Public Property Get pblnInUse() As Boolean
    pblnInUse = mblnInUse
End Property

Private Sub cchkClear_Click()
    On Error Resume Next
    ctxtName.SetFocus
End Sub

Private Sub cchk¼�뵥λ����_Click()
    Dim lblnVisible As Boolean
    On Error Resume Next
    If cchk¼�뵥λ����.Value = 1 Then
        lblnVisible = True
    Else
        lblnVisible = False
    End If
    ccmbUnit.Visible = lblnVisible
    ccmd��λ��λ.Visible = lblnVisible
    Label2(5).Visible = lblnVisible
    ctxtName.SetFocus
End Sub

Private Sub ccmbSex_GotFocus()
    On Error Resume Next
    If ccmbSex = "" And ccmbSex.ListCount > 0 Then
        ccmbSex.ListIndex = 0
    End If
End Sub
Private Sub ccmbSex_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ctxtAge.SetFocus
    End If
End Sub

Private Sub ccmb���������_Click()
    Dim lobj������� As Object
    On Error GoTo errHandler
    
    Set lobj������� = CreateObject("ְҵ������.clsmedicalexam")
    lobj�������.������� = ccmb���������.ItemData(ccmb���������.ListIndex)
    
    '2012-06-14 �ڵ�� ��
    '���ݲ�ͬ�����Ա���ͣ����ɲ�ͬ��ϵͳ���
    '2015-3-13��ΰ ȡ�����ݲ�ͬ�����Ա���ͣ����ɲ�ͬ��ϵͳ��ţ����ֱ�Ų���
   ' If Len(clblsysno.Text) = 0 Then
    '    clblsysno.Text = lobj�������.Func����ְҵ�����ϵͳ��� & (ccmb���������.ListIndex + 1)
   ' Else
   '     clblsysno.Text = Left(clblsysno.Text, Len(clblsysno.Text) - 1) & (ccmb���������.ListIndex + 1)
   ' End If
    mobj���.ϵͳ��� = Trim(clblsysno.Text)
    mobj���.�����Ա.ϵͳ��� = Trim(clblsysno.Text)
    '2012-06-14 �ڵ�� ��
    '2012-12-18 ������  ��
    'BUG�ţ�0000092
    If InStr(ccmb���������.Text, "����") > 0 Then
'modify by lanchao 2015-03-12 if else ע��ȡ��
        Ccmb��������.Text = "�ڸ��ڼ�"
    Else
       Ccmb��������.ListIndex = 0
    End If
    Call Ccmb��������_Click
    '2012-12-18 ������  ��
    Exit Sub
errHandler:
    sfsub������ "ְҵ������", "frmregister", "Private Sub ccmb���������_Click", Err.Number, Err.Description, True
End Sub

Private Sub ccmb����Դ_click()
    ccmbְҵ���.Visible = True
    Call funcְҵ���
    Exit Sub
End Sub

'2012-07-11 �ڵ��
'�ж��Ƿ���������
Private Sub ccmdTakePhotoAgain_Click()
    'Check���֤.Value = 1
    If Check���֤.Value = 1 Then
        Check���֤.Value = 0
    Else
        Check���֤_Click
    End If
    ccmdTakePhotoAgain.Visible = False
End Sub

'2012-07-16 �ڵ��
'˫��ĳ�����뵱ʱ����������Ϣ
Private Sub cgrdHistory_DblClick()
    Dim lstrSysNo As String
    lstrSysNo = clblsysno.Text
    clblsysno.Text = cgrdHistory.TextMatrix(cgrdHistory.Row, 0)
    clblsysno_LostFocus
    clblsysno.Text = lstrSysNo        'ֻ��ץȡ��ǰ����Ϣ�������±�ţ�
End Sub

Private Sub ctxt������_Change()
    If Len(Trim(ctxt������.Text)) > 50 Then
        ctxt������.Text = Left(Trim(ctxt������.Text), 50)
    End If
End Sub

Private Sub ctxt�绰_Change()
    If Len(Trim(ctxt�绰.Text)) > 11 Then
        ctxt�绰.Text = Left(Trim(ctxt�绰.Text), 11)
    End If
End Sub

Private Sub ctxt����_Change()
    If Len(Trim(ctxt����.Text)) > 2 Then
        ctxt����.Text = Left(Trim(ctxt����.Text), 2)
    End If
End Sub

'Private Sub ctxt����_Change()
'    If Len(Trim(ctxt����.Text)) > 2 Then
'        ctxt����.Text = Left(Trim(ctxt����.Text), 2)
'    End If
'End Sub

Private Sub ctxt��ϵ�绰_Change()
    If Len(Trim(ctxt��ϵ�绰.Text)) > 11 Then
        ctxt��ϵ�绰.Text = Left(Trim(ctxt��ϵ�绰.Text), 11)
    End If
End Sub

Private Sub ctxt���֤��_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ctxtName.SetFocus
        sub�鿴��ʷ��Ϣ (ctxt���֤��.Text)
    End If
End Sub
Private Sub ccmbTemplate_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        If ctxtTubeNo.Visible Then
            ctxtTubeNo.SetFocus
        Else
            ctxt��쵥��.SetFocus
        End If
    End If
End Sub

'���ܣ����Ʋ��������������ƣ�ֻ��ѡ��
Private Sub ccmbTemplate_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub ccmbUnit_Click()
    On Error GoTo errHandler
    Dim i As Integer
    
    '�ж�¼��ĵ�λ�Ƿ����б��д��ڣ�������������б�
    i = gffuncItemIsInComboBox(ccmbUnit, ccmbUnit.Text)
    If i = -1 Then
        '���뵽�б����
        ccmbUnit.AddItem ccmbUnit.Text
        
        '���ص��������䲾�ļ���
        pobjҵ�����.���չ������䲾.sub���ӵ�λ���� mstr��λ������ & "|" & ccmbUnit.Text
    Else
        '�޸ģ�2001-8-23��
        On Error Resume Next
        mstr��λ������ = pobjҵ�����.���չ������䲾.��λ���(ccmbUnit.Text)
        sub��ʾ��λ���� ciptBase, mstr��λ������, mobjGUI
    End If
    Exit Sub
errHandler:
    sfsub������ "ְҵ������", "frmregister", "Sub ccmbUnit_Click", Err.Number, Err.Description, True

End Sub

Private Sub ccmbUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        If ctxt����.Visible Then
            ctxt����.SetFocus
        Else
            If ciptBase.Visible Then
                ciptBase.SetFocus
            End If
        End If
    Else
        mstr��λ������ = ""
    End If
        
End Sub

Private Sub Ccmb��������_Click()
    Dim lobj����ģ�弯 As Object
    Dim lobj������ As Object
    Dim lcolInfo As New Collection
    Dim lcol������ As Collection
    Dim i As Integer
    On Error GoTo errHandler
     '�������������Ͽ���
    'Set lobj������ = CreateObject("ְҵ������.clsmedicalexam")
    'lobj������.������ = ccmb��������.ItemData(ccmb��������.ListIndex)
    'lobj������.������� = 1
    'Set lcol��� = lobj������.������
    'ccmb��������.AddItem ""
    'For i = 1 To lcol���.recordCount
    '    ccmb��������.AddItem lcol���("���")
    '    ccmb��������.ItemData(ccmb��������.NewIndex) = lcol���("���")
    '    lcol���.movenext
    'Next
    'ccmb��������.ListIndex = 0
    'Set lobj������ = Nothing
   
    
    '�����еķǸ�������ģ����뵽���������б���С��ټ�����������
    ccmbTemplate.Clear
    Set lobj����ģ�弯 = CreateObject("ְҵ������.ClsMedicalExamTemplateSet")
    lobj����ģ�弯.�������� = Trim(ccmb���������.Text)
    'lobj����ģ�弯.������� = ccmb��������.ItemData(ccmb��������.ListIndex)
    lobj����ģ�弯.������� = Trim(Ccmb��������.Text)
    Set lcolInfo = lobj����ģ�弯.Ԫ�ؼ�
    Set lcol������ = lobj����ģ�弯.������Ԫ�ؼ�
    'ccmbTemplate.ListIndex = 0
    If lcolInfo.Count = 0 Then Exit Sub
    For i = 1 To lcolInfo.Count
        ccmbTemplate.AddItem lcolInfo(i)
        ccmbTemplate.ItemData(ccmbTemplate.NewIndex) = lcol������(i)
    Next
    ccmbTemplate.Text = ccmbTemplate.List(0)
    
    Set lobj����ģ�弯 = Nothing
    Call ccmbTemplate_Click
    Exit Sub
errHandler:
    sfsub������ "ְҵ������", "frmregister", "ccmb��������_click", Err.Number, Err.Description, True
End Sub

Private Sub ccmb���������_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error Resume Next
    If KeyCode = 13 Then
        ctxt����.SetFocus
    End If
End Sub

Private Sub cdtpDate_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ctxtName.SetFocus
    End If
End Sub

Private Sub Check���֤_Click()
    On Error GoTo errHandler
    '2012-06-20 �ڵ�� ��
    'ֻ��У��ͨ���󣬲Ż��������Ȩ�ޡ�
    'δͨ��ʱ��ֻ��ˢ���֤��
    If Check���֤.Value = 0 Then   '����
        Timer2.Enabled = False
        If mintState = 1 Then  'mintstate=1��ʾͨ��У�ˣ���������
            Picture1.Visible = False
            cctlCatchPhoto.Visible = True
            cctlCatchPhoto.funcInitVideo
            cctlCatchPhoto.Enabled = True
            mblnTakePhoto = True
            '2012-04-14 �ڵ�� ��
            '����ˢ�������֤ʱ������������Ƭ��Ҳ����������ͷ����
            ctbMain.Buttons(4).Enabled = True
            '2012-04-14 �ڵ�� ��
        End If
        
        '2012-07-11 �ڵ�� ��
        'ˢ�������֤���������Ա����䣬�������ڣ����֤��  enabled=false ,����true
        'У��֮��(mintstate=1)���������޸Ļ�����Ϣ��
        '�������Ͽ��ƣ�ֻ��Ϊ�˲鿴ʱ������Ϊ�����޸ġ�ʵ���ϣ�ֻ�д�ʱ���յ���Ƭ���Ա��档��
        ctxt���֤��.Enabled = (mintState <> 1) 'True
        ctxtName.Enabled = (mintState <> 1)  'True
        ccmbSex.Enabled = (mintState <> 1)  'True
        ctxtAge.Enabled = (mintState <> 1)  'True
        cdtp����.Enabled = (mintState <> 1)  'True
        '2012-07-11 �ڵ�� ��
        
        Label31.Visible = False
    Else                            'ˢ���֤
        Picture1.Visible = True
        cctlCatchPhoto.Visible = False
        ctxt���֤��.Enabled = False
        ctxtName.Enabled = False
        ccmbSex.Enabled = False
        ctxtAge.Enabled = False
        cdtp����.Enabled = False
        Label31.Visible = True
        If mblnTakePhoto Then
            cctlCatchPhoto.subDisconnect
            mblnTakePhoto = False
        End If
        '2012-04-14 �ڵ�� ��
        '��ˢ�������֤ʱ������������Ƭ
        ctbMain.Buttons(4).Enabled = False
        '2012-04-14 �ڵ�� ��
    End If
    
    Exit Sub
errHandler:
    sfsub������ "ְҵ������", "frmregister", "Sub Check���֤_Click", Err.Number, Err.Description, True
End Sub

'�˹����ݲ���
'Private Sub ciptBase_LastLostFocus()
'    Dim blnCancel As Boolean
'    On Error Resume Next
    '�Զ����档
 '   If ctbMain.Buttons(6).Enabled Then
 '       ctxtName.SetFocus
 '       SendKeys "{F2}"
 '   End If
'End Sub

'Private Sub ciptBase_LostFocus()
'    On Error Resume Next
'    If ActiveControl.Name <> "c�ֵ��" Then
 '       c�ֵ��.Visible = False
 '   End If
'End Sub


Private Sub clblsysno_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ccmb���������.SetFocus
    End If
End Sub

Private Sub clblsysno_LostFocus()
    Dim lobjRec As Object
    Dim strSQL As String
    'Dim lobjϵͳ��� As Object
    Dim strTmp As String
    Dim str��λ������ As String
    '2012-06-13 �ڵ�� ��
    '��ȡ���֤��Ƭ��������ȡ�ֳ���Ƭ����
    Dim lobj���֤��Ƭ As Object
    Dim lobj�ֳ���Ƭ As Object
    '2012-06-13 �ڵ�� ��
    
    On Error GoTo errHandler
    strTmp = Trim(clblsysno.Text)
    
    '2012-07-11 �ڵ�� ��
    'ϵͳ��Ź̶��ˣ��Ͳ����ٸ����ˡ�
    clblsysno.Enabled = False
    '2012-07-11 �ڵ�� ��
    
    '2012-06-14 �ڵ�� ��
    '��Ϊ������¼����Ϣ����������ŵ������������ж�ȡ��
    If Len(clblsysno.Text) = 0 Then
        'MsgBox "ϵͳ��Ŵ������飡", vbInformation, "ϵͳ��ʾ"
        Exit Sub
    End If
    '2012-06-14 �ڵ�� ��
    
    Set lobjRec = dafuncGetData("select * from ְҵ�����_���������ݿ� where ϵͳ���='" & strTmp & "'")
    If lobjRec.RecordCount = 0 Then
        mobj���.ϵͳ��� = Trim(clblsysno.Text)
        'mobj���. = Trim(clblSysNo.Text)
    ElseIf lobjRec.RecordCount = 1 Then
        '2012-07-11 �ڵ�� ��
        '������Ϣʱ�����������ļ�����ʱȱ�ٲ�����Ϣ����ʱ�����Դ��������һ�д��롣
        On Error Resume Next
        '2012-07-11 �ڵ�� ��
        mobj���.ϵͳ��� = Trim(clblsysno.Text)
        ccmb���������.Text = IIf(IsNull(lobjRec!�������), "", lobjRec!�������)
        Ccmb��������.Text = IIf(IsNull(lobjRec!������), "", lobjRec!������)
        ccmbTemplate.Text = IIf(IsNull(lobjRec!������), "", lobjRec!������)
        
        '�����ڸ�Ϊ��ǰϵͳ������ 2015-10-16
        cdtpDate.Value = Now
'        cdtpDate.Value = IIf(IsNull(lobjRec!��������), "", lobjRec!��������)

        ctxt���֤�� = IIf(IsNull(lobjRec!������ݺ���), "", lobjRec!������ݺ���)
        ctxtName = IIf(IsNull(lobjRec!����), "", lobjRec!����)
        ccmbSex = IIf(IsNull(lobjRec!�Ա�), "", lobjRec!�Ա�)
        ctxtAge = IIf(IsNull(lobjRec!����), "", lobjRec!����)
        cdtp���� = IIf(IsNull(lobjRec!��������), "", lobjRec!��������)
        ctxt���� = IIf(IsNull(lobjRec!����), "", lobjRec!����)
        ctxt�ʱ� = IIf(IsNull(lobjRec!�ʱ�), "", lobjRec!�ʱ�)
        ctxtסַ = IIf(IsNull(lobjRec!סַ), "", lobjRec!סַ)
        ccmb�Ļ��̶� = IIf(IsNull(lobjRec!�Ļ��̶�), "", lobjRec!�Ļ��̶�)
        Ccmb��� = IIf(IsNull(lobjRec!���), "", lobjRec!���)
        ccmb���� = IIf(IsNull(lobjRec!����), "", lobjRec!����)
        ctxt�绰 = IIf(IsNull(lobjRec!�绰����), "", lobjRec!�绰����)
        ctxt���� = IIf(IsNull(lobjRec!����), "", lobjRec!����)
        ctxt������ = IIf(IsNull(lobjRec!������), "", lobjRec!������)
        ccmb����Դ = IIf(IsNull(lobjRec!����Դ), "", lobjRec!����Դ)
        ccmbְҵ��� = IIf(IsNull(lobjRec!ְҵ����), "", lobjRec!ְҵ����)
        ccmbΣ������ = IIf(IsNull(lobjRec!Σ������), "", lobjRec!Σ������)
        ccmb�ֹ��� = IIf(IsNull(lobjRec!�ֹ���), "", lobjRec!�ֹ���)
        ccmbְ�� = IIf(IsNull(lobjRec!ְ���ְ��), "", lobjRec!ְ���ְ��)
        ctxtΣ������ = IIf(IsNull(lobjRec!ְҵΣ������), "", lobjRec!ְҵΣ������)
        ctxt������� = IIf(IsNull(lobjRec!�������), "", lobjRec!�������)
        str��λ������ = IIf(IsNull(lobjRec!��λ������), "", lobjRec!��λ������)
        
       ' ��ȡ��Ƭ
        Set lobjRec = CreateObject("ְҵ������.clspersonexamed")
        lobjRec.ϵͳ��� = Trim(clblsysno.Text)
        If lobjRec.��Ƭ <> 0 Then
            Picture1.Picture = lobjRec.��Ƭ
            Picture1.Visible = True
            cctlCatchPhoto.Visible = False
            If mblnTakePhoto Then
                cctlCatchPhoto.subDisconnect
                mblnTakePhoto = False
            End If
        End If
        
'''        '2012-07-11 �ڵ�� ��
'''        '������ȡ�ֳ���Ƭ
'''        Set lobj�ֳ���Ƭ = lobjRec.func��ȡ�ֳ���Ƭ(Trim(clblsysno.Text), "ְҵ�����")
'''        If Not lobj�ֳ���Ƭ Is Nothing Then Picture1.Picture = lobj�ֳ���Ƭ
'''        Picture1.Visible = True
'''        '2012-07-11 �ڵ�� ��
        
        
        '2012-06-13 �ڵ�� ��
        '��ȡ�������Ա���֤��Ƭ
        Set lobj���֤��Ƭ = lobjRec.func�������֤��Ƭ(Trim(clblsysno.Text) & "IDcard", "ְҵ�����")
        If Not lobj���֤��Ƭ Is Nothing Then
            Picture2.Picture = lobj���֤��Ƭ
            '
            If mintState = 1 And mblnTakePhoto = False Then ccmdTakePhotoAgain.Visible = True
            
            '2015-3-2
'            Else
'            End If
        End If
        Picture2.Visible = True
        '2012-06-13 �ڵ�� ��

        If FrmRegister.pstr����ϵͳ��� <> "" Then
            ccmdTakePhotoAgain.Visible = False
            Me.ctbMain.Buttons(7).Visible = False
        End If

        On Error GoTo errHandler
        If Not IsNull(str��λ������) Then
            func��ȡ��λ��Ϣ str��λ������
        End If
    Else
        MsgBox "ϵͳ��Ų�Ψһ�����飡", vbInformation, "ϵͳ��ʾ"
        Exit Sub
    End If
    
    Set lobjRec = Nothing
    Exit Sub
errHandler:
    sfsub������ "ְҵ������", "frmregister", "Sub clblsysno_LostFocus", Err.Number, Err.Description, True
End Sub

Private Sub ctxtAge_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ctxt��쵥��.SetFocus
    End If
End Sub

Private Sub ctxtName_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ccmbSex.SetFocus
    End If
End Sub

Private Sub ctxtTubeNo_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        cdtpDate.SetFocus

    End If
End Sub
'����
Private Sub ctxt����_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        '��¼���û��¼����Ŀ����ֱ�ӱ��档
        If ciptBase.Visible Then
            ciptBase.SetFocus
            ciptBase.ItemSetFocus 0
        End If
    End If
End Sub
'¼�����֤�ź󣬻�ȡ���䣬�Ա𣬳�������
Private Sub ctxt���֤��_lostfocus()
    Dim ldatBirth As String
    Dim lstrSex As String
    On Error GoTo errHandler
    If Trim(ctxt���֤��.Text) <> "" Then
            '��ȷʱ�����֤���л�ȡ�������ڡ�
            sub���ݹ�����ݺ����ȡ���պ��Ա� ctxt���֤��.Text, ldatBirth, lstrSex
            If Not IsDate(ldatBirth) Then
                MsgBox ("���֤�Ų��Ϸ���")
                Exit Sub
            End If
            
            '�����Ƿ���Ҫ¼��������ڣ���Ҫʱ�Զ��������֤����д��������
            On Error Resume Next
            If IsDate(ldatBirth) Then
                cdtp����.Value = ldatBirth
'                ctxtAge.Text = DateDiff("yyyy", ldatBirth, Date)
                ctxtAge.Text = Year(Date) - Year(ldatBirth)
'����� 2012-12-11 ��
'˵�����������ж����֤�������Ƿ���˵�ǰ���ڣ��������һ�ꡣ
'             If Month(Date) > Month(ldatBirth) Then
'                 ctxtAge.Text = Year(Date) - Year(ldatBirth) + 1
'             ElseIf Month(Date) = Month(ldatBirth) Then
'                If Day(Date) >= Day(ldatBirth) Then
'                    ctxtAge.Text = Year(Date) - Year(ldatBirth) + 1
'                Else
'                    ctxtAge.Text = Year(Date) - Year(ldatBirth)
'                End If
'             Else
'                 ctxtAge.Text = Year(Date) - Year(ldatBirth)
'             End If
             
    '����� 2012-12-11 ��
            End If
            ccmbSex.Text = lstrSex
    End If
    Exit Sub
errHandler:
    sfsub������ "ְҵ������", "frmregister", "Sub ctxt���֤��_lostfocus", Err.Number, Err.Description, True
End Sub

Private Sub ctxt��쵥��_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        If ccmbUnit.Visible Then
            ccmbUnit.SetFocus
        Else
            If ctxt����.Visible Then
                ctxt����.SetFocus
            Else
                If ciptBase.Visible Then
                    ciptBase.SetFocus
                End If
            End If
        End If
    End If
End Sub

Private Sub ctxt�ʱ�_Change()
    If Len(Trim(ctxt�ʱ�.Text)) > 6 Then
        ctxt�ʱ�.Text = Left(Trim(ctxt�ʱ�.Text), 6)
    End If
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If mblnTakePhoto Then
        '���³�ʼ������ؼ���
        cctlCatchPhoto.funcInitVideo
    End If
    
    
 

    'ctxtName.SetFocus
End Sub

Private Sub Form_Deactivate()
    On Error Resume Next
    gfsubHideComboList ccmbUnit
End Sub

'��ʼ������
Private Sub Form_Load()
    On Error GoTo errHandler
   
    '��������Ѿ���ʼ�������ٽ��г�ʼ����
    If mblnInUse Then Exit Sub
    
    '���ô�������ʹ�õı�־��
    mblnInUse = True
    
    MousePointer = 11
    
    '���治�ɲ�����
'    cfram������Ϣ.Enabled = False
    'ctbMain.Enabled = False
    'clblsysno.Visible = False
    Set mcol�շ���Ŀ = New Collection
    Set mcol�����Ŀ = New Collection
    
    Set mobj����� = CreateObject("ְҵ������.clsMedicalExam")
    
    Set mobj��� = CreateObject("ְҵ������.clsMedicalExam")
    '�޸ģ�2002-10-10������ϵͳ������ƣ���
    If pstrϵͳ������� <> "" Then
        mobj���.ϵͳ������� = pstrϵͳ�������
    End If
    
    Set mobj��켯 = CreateObject("ְҵ������.clsMedicalExamSet")
    Set mobj����ģ�� = CreateObject("ְҵ������.ClsMedicalExamTemplate")
    
    '��ʼ������ͨ�ö���ÿ�������Ӧһ������ͨ�ö���ʵ�������ɻ��ã��м��мǣ���
    Set mobjGUI = New cls����ͨ�ö���
    mobjGUI.pbln�Զ������ֵ�߶� = False
    
    '���ù�����������Ҫ�ĸ��ְ�ť��
    Dim lcol��������ť As New Collection           '�������ϵİ�ť��ʼ�����ϡ�
    With lcol��������ť
        .Add "���(&Cl)110"
        .Add "|"
        .Add "�����Ŀ(&T)102"
        .Add "������Ƭ(&E)103"
        .Add "|"
        .Add "����"
        '2012-06-15 �ڵ�� ��
        '��֤�Ƿ�ͨ��У�ˣ�ȡ���޸İ�ť
        .Add "У��ͨ��(&J)106"
        '.Add "�޸�"
        '2012-06-15 �ڵ�� ��
        .Add "|"
        .Add "�˳�"
    End With
    
    'Ϊ��Ҫͨ������ͨ�ö�����Ƶĸ��ֿؼ�����ʼֵ��
    With mobjGUI
        Set .Form = Me
        Set .c������ = ctbMain
        Set .c¼��� = ciptBase
        Set .c�ֵ�� = c�ֵ��
        Set .c״̬�� = cstbMain
        
        '���ý���ͨ�ö����ṩ�ķ������Խ���ؼ����г�ʼ����
        .subInitialize lcol��������ť, ""
    End With
    
    ctbMain.Buttons(1).Visible = False
    ctbMain.Buttons(2).Visible = False
'    ctbMain.Buttons(3).Visible = False
    ctbMain.Buttons(4).Visible = False
'    ctbMain.Buttons(5).Visible = False
    ctbMain.Buttons(7).Visible = False
    
    If ���ʱ�־ = 2 Then
        ccmb���������.Enabled = False
        Ccmb��������.Enabled = False
        ccmbTemplate.Enabled = False
        ctbMain.Buttons(1).Enabled = False
    End If
    
    '2012-06-15 �ڵ�� ��
    '��Ӹý����Ȩ��
    Dim lobjTmp As Object
    Set lobjTmp = CreateObject("ְҵ������.clsPermissionConfigure")
    If lobjTmp.func���Ҳ���Ȩ��(um�û����, "ְҵ�����_���Ǽ�_����Ǽ�_У��ͨ��") = False Then
        ctbMain.Buttons(7).Visible = False
    End If
    '2012-06-15 �ڵ�� ��
    
    '���� SSTab �ؼ��ĵ�ǰѡ�
    SSTab1.Tab = 0
    
    '���
    subClear
    cdtpDate.Value = Now
    
    '�·���ϵͳ���
    'clblsysno.Caption = mobj���.Func����ϵͳ���
    'clblSysNo.Text = ""
    mobj���.ϵͳ��� = Trim(clblsysno.Text)
    'pstrϵͳ��� = Trim(clblSysNo.Text)
    If Check���֤.Value = 0 Then
        Check���֤.Value = 1
    Else
        Check���֤_Click
    End If
    'cctlCatchPhoto.Visible = False
    'cctlCatchPhoto.Visible = True
    
'    If pobjҵ�����.ҵ������("�Թܱ���Զ�����") = "��" Then
'        ctxtTubeNo.Visible = True
'        ctxtTubeNo.TabIndex = 1
'        clblTubeNo.Visible = False
'        clblLetter.Visible = False
'        cvscLetter.Visible = False
'    Else
'        ctxtTubeNo.Visible = False
'        clblTubeNo.Visible = True
'        clblLetter.Visible = True
'        cvscLetter.Visible = True
'    End If
    
    cgrdHistory.Visible = False
    DoEvents
    ccmb���������.Visible = True
    Label2(3).Visible = True
    clbl���������.Visible = True

    'Ϊ�˼ӿ촰������ٶȣ����³�ʼ���������ڶ�ʱ������ɡ�
    Timer1.Enabled = True
    Timer2.Enabled = True

    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����沿��", "FrmRegisterAnnual", "Form_Load", 6666, lstrError, False
    '�ָ����������á�
    ctbMain.Enabled = True
    MousePointer = 0
    cstbMain.Panels(1) = lstrError
End Sub





'���ܣ����form_load���µĳ�ʼ��������
Private Sub Timer1_Timer()
    Dim lobj����ģ�弯 As Object  '����ģ�弯����ȡ���еķǸ�������ģ�����ơ�
    Dim lcolInfo As Collection
    Dim lcol��� As Object
    Dim lcol���� As Object
    Dim i As Integer
    Dim lobj������ As Object
    Dim lobj������� As Object
    Dim lobjRec As Object
    On Error GoTo errHandler
    
    '��ʱ�����������á�
    Timer1.Enabled = False
    
    '�ӵ��չ������Ѳ��л�ȡ����¼����ĵ�λ���ơ�
    Set lcolInfo = pobjҵ�����.���չ������䲾.��λ���Ƽ�
    For i = 1 To lcolInfo.Count
        ccmbUnit.AddItem lcolInfo(i)
    Next
    
    '�������������Ͽ���
    Set lobjRec = pobjDict.FetchEx("��������ֵ�")
    Ccmb��������.Clear
    'Ccmb��������.AddItem ""
    For i = 1 To lobjRec.RecordCount
        Ccmb��������.AddItem lobjRec("����")
        Ccmb��������.ItemData(Ccmb��������.NewIndex) = lobjRec("���")
        lobjRec.MoveNext
    Next
'    Ccmb��������.ListIndex = 0
    '2012-06-15 �ڵ�� ��
    '�޸������Ա��Ϣ����������ʱ����ֹ����ų���
    If clblsysno.Text = "" Then
        Ccmb��������.ListIndex = 0
    Else
        If Right(clblsysno.Text, 1) < "0" Or Right(clblsysno.Text, 1) > "9" Then
            Ccmb��������.ListIndex = CInt(Left(Right(clblsysno.Text, 2), 1) - 1)
        Else
            Ccmb��������.ListIndex = CInt(Right(clblsysno.Text, 1) - 1)
'            Ccmb��������.ListIndex = 0
        End If
    End If
    '2012-06-15 �ڵ�� ��
   
    Set lobjRec = pobjDict.FetchEx("���������ֵ�")
    ccmb���������.Clear
    'ccmb���������.AddItem ""
    For i = 1 To lobjRec.RecordCount
        ccmb���������.AddItem lobjRec("����")
        ccmb���������.ItemData(ccmb���������.NewIndex) = lobjRec("���")
        lobjRec.MoveNext
    Next
    '2012-06-15 �ڵ�� ��
    '�޸������Ա��Ϣ����������ʱ����ֹ����ų���
    If clblsysno.Text = "" Then
        ccmb���������.ListIndex = 0
    End If
    '2012-06-15 �ڵ�� ��
    
    '�����еķǸ�������ģ����뵽���������б���С�
    'Set lobj����ģ�弯 = CreateObject("ְҵ������.ClsMedicalExamTemplateSet")
    'lobj����ģ�弯.�������� = 3
    
    'lobj����ģ�弯.������� = ccmb��������.ItemData(ccmb��������.ListIndex)
    'Set lcolInfo = lobj����ģ�弯.Ԫ�ؼ�
    'For i = 1 To lcolInfo.Count
    '    ccmbTemplate.AddItem lcolInfo(i)
    'Next
    'ccmbTemplate.Text = ccmbTemplate.List(0)
    'Set lobj����ģ�弯 = Nothing
    
    '2012-06-15 �ڵ�� ��
    '�ù����ڸú����������У���ע�͵�
'''    '����ҵ�������ж��Ƿ�����,�ж�ʱ���Ͻ����ϵ��Ƿ�ˢ�������֤��
'''    If pobjҵ�����.ҵ������("�Ƿ�����") = "��" Then
'''        mblnTakePhoto = True
'''    Else
'''        mblnTakePhoto = False
'''    End If
    '2012-06-15 �ڵ�� ��
    
    If pobjҵ�����.ҵ������("�Ƿ���ٵǼ�") = "��" Then
        mbln����¼�� = True
    Else
        mbln����¼�� = False
    End If
    
    'ֻ�г��죬���ҿ��ٵǼǲſ��������Ǽǡ�
    If Not mbln����¼�� Or pstrϵͳ��� <> "" Then
        clbl����.Visible = False
        ctxt����.Visible = False
    End If
    
    'ccmb���������.ListIndex = 0
    
    If ccmbTemplate.ListCount > 0 Then
        'ccmbTemplate.ListIndex = 0
'        ccmbTemplate.Text = ccmbTemplate.List(0)
'        subChangeTemplate
        
    End If
    
'    If pstrϵͳ��� <> "" Then
'        '�����Ǽǡ�
'        '��ʾ�����Ա������Ϣ��
'        SubGetPersonInfo pstrϵͳ���
'    End If
    
    On Error Resume Next
    Set mobj���� = New cls�û���������
    mobj����.�û���� = "*"
    mobj����.ҵ���� = "������"
    mstrĬ������ = mobj����.������ֵ("�������")
'    If mstrĬ������ <> "" And ctxtAge = "" Then
'        ctxtAge = mstrĬ������
'    End If
    
    If mobj����.������ֵ("���Ǽ�ʱ¼�뵥λ����") = "" Or mobj����.������ֵ("���Ǽ�ʱ¼�뵥λ����") = "��" Then
        cchk¼�뵥λ����.Value = 1
    Else
        cchk¼�뵥λ����.Value = 0
    End If
    cfram������Ϣ.Enabled = True
    ctbMain.Enabled = True
    
    '2012-06-15 �ڵ�� ��
    'ʡ������Ҫ�󣬳�ʼ�Ǽ�ֻˢ���֤��У��ͨ���������ࡣ
'''    '��Ҫ����ʱ��ʼ������ؼ���
'''    If mblnTakePhoto And Check���֤ = False Then
'''        '��ʼ���ؼ���
'''        cctlCatchPhoto.funcInitVideo
'''        '����ؼ���visible=false��visible=true��ˢ��һ�Σ����� ��ȡ�� ��ť����������ʾ
'''        Picture1.Visible = False
'''        cctlCatchPhoto.Visible = False
'''        cctlCatchPhoto.Visible = True
'''        cctlCatchPhoto.Enabled = True
'''    Else
'''        cctlCatchPhoto.Visible = False
'''        Picture1.Visible = True
'''        mblnTakePhoto = False
'''    End If
'''    If mblnTakePhoto Then
'''        cctlCatchPhoto.subDisconnect
'''        mblnTakePhoto = False
'''    End If
    Dim lstrTmp As String
    Set lobjRec = CreateObject("ְҵ������.clsManageMedicalExam")
    lstrTmp = lobjRec.func��ȡ���˵�ǰ���״̬(Trim(clblsysno.Text))
    If lstrTmp = "δУ��" Or lstrTmp = "" Or (���ʱ�־ = 1 And lstrTmp <> "δ���嵥") Then   'ֻ�е��Ǽ�δУ��ʱ����ˢ���֤
        If Check���֤.Value = 0 Then
            Check���֤.Value = 1
        Else
            Check���֤_Click
        End If
        mintState = 0
        clblHintCheck.Visible = False
        sub��������ʼ��
       func���֤��֤
       
    ElseIf lstrTmp = "δ���嵥" Then
        mintState = 1
'        clblHintCheck.Visible = True
        'If ���ʱ�־ = 0 Then Check���֤.Value = 0:Check���֤_Click
        If Check���֤.Value = 1 Then
            Check���֤.Value = 0
        Else
            Check���֤_Click
        End If
    ElseIf lstrTmp = "������" Then  '�����������ҵ���δ���嵥���ʼ��������ˢ���֤���ܡ�
        mintState = 1
'        clblHintCheck.Visible = True
        If Check���֤.Value = 1 Then
            Check���֤.Value = 0
        Else
            Check���֤_Click
        End If
    End If
    '2012-06-15 �ڵ�� ��
    
    MousePointer = 0

    '��ȡ�Ļ��̶�
    Set lobjRec = pobjDict.FetchEx("�Ļ��̶��ֵ�")
    ccmb�Ļ��̶�.Clear
    ccmb�Ļ��̶�.AddItem ""
    For i = 1 To lobjRec.RecordCount
        ccmb�Ļ��̶�.AddItem lobjRec("����")
        ccmb�Ļ��̶�.ItemData(ccmb�Ļ��̶�.NewIndex) = lobjRec("���")
        lobjRec.MoveNext
    Next
    ccmb�Ļ��̶�.ListIndex = 0
    
    '��ȡ���
    Set lobjRec = pobjDict.FetchEx("�����ֵ�")
    Ccmb���.Clear
    Ccmb���.AddItem ""
    For i = 1 To lobjRec.RecordCount
        Ccmb���.AddItem lobjRec("����")
        Ccmb���.ItemData(Ccmb���.NewIndex) = lobjRec("���")
        lobjRec.MoveNext
    Next
    Ccmb���.ListIndex = 0
    
     '��ȡ����
    Set lobjRec = pobjDict.FetchEx("�����ֵ�")
    ccmb����.Clear
    ccmb����.AddItem ""
    For i = 1 To lobjRec.RecordCount
        ccmb����.AddItem lobjRec("����")
        ccmb����.ItemData(ccmb����.NewIndex) = lobjRec("���")
        lobjRec.MoveNext
    Next
    ccmb����.ListIndex = 0
    
     '��ȡ��������
    Set lobjRec = pobjDict.FetchEx("���������ֵ�")
    ccmb��������.Clear
    ccmb��������.AddItem ""
    For i = 1 To lobjRec.RecordCount
        ccmb��������.AddItem lobjRec("����")
        ccmb��������.ItemData(ccmb��������.NewIndex) = lobjRec("���")
        lobjRec.MoveNext
    Next
    ccmb��������.ListIndex = 0
    
    '��ȡΣ������
    Set lobjRec = pobjDict.FetchEx("Σ�������ֵ�")
    ccmbΣ������.Clear
    ccmbΣ������.AddItem ""
    For i = 1 To lobjRec.RecordCount
        ccmbΣ������.AddItem lobjRec("����")
       ccmbΣ������.ItemData(ccmbΣ������.NewIndex) = lobjRec("���")
        lobjRec.MoveNext
    Next
    ccmbΣ������.ListIndex = 0
    
    '��ȡְҵ��ְ��
    Set lobjRec = pobjDict.FetchEx("ְҵ��ְ���ֵ�")
    ccmbְ��.Clear
    ccmbְ��.AddItem ""
    For i = 1 To lobjRec.RecordCount
        ccmbְ��.AddItem lobjRec("����")
        ccmbְ��.ItemData(ccmbְ��.NewIndex) = lobjRec("���")
        lobjRec.MoveNext
    Next
    ccmbְ��.ListIndex = 0
    
    '��ȡ�����ֵ�
    Set lobjRec = pobjDict.FetchEx("�����ֵ�")
    ccmb�ֹ���.Clear
    ccmb�ֹ���.AddItem ""
    For i = 1 To lobjRec.RecordCount
        ccmb�ֹ���.AddItem lobjRec("����")
        ccmb�ֹ���.ItemData(ccmb�ֹ���.NewIndex) = lobjRec("���")
        lobjRec.MoveNext
    Next
    ccmb�ֹ���.ListIndex = 0
    Call func��ȡ��ҵ����ֵ�
    Call func����Դ
    clblsysno.Visible = True
    clblsysno.SetFocus
     '2012-07-11 �ڵ�� ��
    'ϵͳ�������֮�󣬲����Ըı䡣ת��focus֮ǰ���������趨setfocus��������֤������Ϣ���Լ��ؽ�ȥ��
    '�Ҽ�����󣬿ؼ�enabled=false
    'If ���ʱ�־ = 1 Then
        ctxt����.SetFocus
    'End If
    '���ʱ�־ = 0       '��֪��ʱ���ֵĴ��롣���ڱ���������������ƵǼǵĲ��ֱ�����̡�
    '2012-07-11 �ڵ�� ��
    
    '2012-08-18 �ڵ�� ��
    '������˻�����Ϣ
    If pstr����ϵͳ��� <> "" Then
        ctbMain.Buttons(4).Enabled = False
    End If
    
    Form_Activate
    clblsysno_LostFocus
    
    '2012-08-18 �ڵ�� ��
    
'''    '2012-08-19 �ڵ�� ��
'''    '����ʱ�����������Ϊ���գ�����Ϊ���콨������
'''    If ���ʱ�־ = 2 And pstr����ϵͳ��� <> "" Then cdtpDate.Value = Now
'''    '2012-08-19 �ڵ�� ��


    Timer2.Enabled = True  '2016-2-24 by Ĳ��
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����沿��", "FrmRegisterAnnual", "Timer1_Timer", 6666, lstrError, False
    '�ָ�����ɲ�����
    cfram������Ϣ.Enabled = True
    ctbMain.Enabled = True
    MousePointer = 0
End Sub

Private Sub ccmbTemplate_Click()
    On Error GoTo errHandler
    'ѡ������
    Dim lcolInfo As Collection
    Dim lstrTubeNo As String
    Dim lstrTemp As String
    Dim i As Integer, j As Integer
    MousePointer = 11
    subChangeTemplate
    ctbMain.Buttons(6).Enabled = True
'    If ctxtTubeNo.Visible Then
'        ctxtTubeNo.SetFocus
'    Else
'        ctxtName.SetFocus
'    End If
    MousePointer = 0
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����沿��", "FrmRegisterAnnual", "ccmbTemplate_Click", 6666, lstrError, False
    
    Exit Sub
    Resume
End Sub

Private Sub subChangeTemplate()
    On Error GoTo errHandler
    'ѡ������
    Dim lcolInfo As Collection
    Dim lstrTubeNo As String
    Dim lstrTemp As String
    Dim i As Integer, j As Integer
    
    '��ȡ���Թܱ�š�
    If mobj���.����.������ <> ccmbTemplate.Text Then
'        mobj���.����.������ = ccmbTemplate.Text

        '��������ģ���ȡ���������п��õ���ĸ��
        mobj����ģ��.������ = ccmbTemplate.Text

'        If pobjҵ�����.ҵ������("�Թܱ���Զ�����") = "��" Then
'            '�Թܱ����ĸΪ��ʱcvscLetter����
'            If mobj���.����.�Թܱ����ĸ = "" Then
'                '����ĸ�����ŷֿ�������mcoltubeNo��
'                lstrTubeNo = mobj����ģ��.�Թ���ĸ���
'                If Right(lstrTubeNo, 1) <> "," Then lstrTubeNo = lstrTubeNo & ","
'                lstrTemp = ""
'                Set mcolTubeNo = New Collection
'                For i = 1 To Len(lstrTubeNo)
'                    lstrTemp = lstrTemp & Mid(lstrTubeNo, i, 1)
'                    If Mid(lstrTubeNo, i, 1) = "," Then
'                        If Left(lstrTemp, Len(lstrTemp) - 1) <> "" Then
'                            mcolTubeNo.Add Left(lstrTemp, Len(lstrTemp) - 1)
'                        End If
'                        lstrTemp = ""
'                    End If
'                Next i
'                If mcolTubeNo.Count > 0 Then
'                    '�Թ���ĸ�ı��ˣ�������ʾ��
'                    If clblLetter.Caption <> "" And clblLetter.Caption <> mcolTubeNo(1) Then
'                        sffuncMsg "��ע�⣬������ѡ�������ʹ�õ��Թ���ĸ��ǰһ����" & clblLetter.Caption & "����ͬ�ˡ�"
'                    End If
'
'                    '��ֵ��clblLetter
'                    clblLetter.Caption = mcolTubeNo(1)
'                    cvscLetter.Enabled = True
'                    cvscLetter.Min = 1
'                    cvscLetter.Max = mcolTubeNo.Count
'                    cvscLetter.Value = 1
'                Else
'                    ctbMain.Buttons(6).Enabled = False
'                    '��ʾ�������޿��õ���ĸ��
'                    Err.Raise 6666, , "�������޿����Թ���ĸ��ţ��������������Ӧ���Թ���ĸ���"
'                End If
'            Else
                '����ĸ������ѡ����ĸ��
'                clblLetter.Caption = mobj���.����.�Թܱ����ĸ
'                cvscLetter.Enabled = False
'            End If
'        Else
'            clblLetter.Caption = mobj����ģ��.�Թ���ĸ���
'        End If
        
        '��ʼ��������Ϣ��
        On Error Resume Next
        mobjGUI.sub��ʼ��¼��� ccmbTemplate.Text
        
        '�޸ģ�2001-8-23����ʾ��λ���ԣ���
        If mstr��λ������ <> "" Then
            sub��ʾ��λ���� ciptBase, mstr��λ������, mobjGUI
        End If

        '������д������Ϣֵ��
'        If mobj����ģ��.����������Ŀ��.Count > 0 Then
'            Set lcolInfo = mobj�����.����.������Ϣ
'            If lcolInfo.Count > 0 Then
'                sub��¼���ֵ ciptBase, mobjGUI, lcolInfo
'            End If
'        End If

        '�޸ģ�2002-7-26��������ݡ��Ƿ�����ѡ���������͡�
        'If mobj����ģ��.�Ƿ����� Then
        '    ccmb���������.ListIndex = 1
        'Else
        '    ccmb���������.ListIndex = 0
        'End If

        '�޸ģ�2002-10-10������ζ����ƣ���ʾ����
        On Error Resume Next
        ciptBase.Box1("�����").Text = mobj����ģ��.�շѱ�׼���
    End If
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����沿��", "FrmRegisterAnnual", "subChangeTemplate", 6666, lstrError, True
    
    Exit Sub
    Resume
End Sub

'�Զ������б��
Private Sub ccmbUnit_GotFocus()
    On Error GoTo errHandler
'    gfsubShowComboList ccmbUnit
    Exit Sub
errHandler:
    'sfsub������ "ְҵ�����沿��", "FrmRegisterAnnual", "ccmbUnit_GotFocus", Err.Number, Err.Description, False
End Sub

Private Sub ccmbUnit_LostFocus()
    On Error GoTo errHandler
    Dim i As Integer
    If Trim(ccmbUnit.Text) = "" Then Exit Sub
    
    '�ж�¼��ĵ�λ�Ƿ����б��д��ڣ�������������б�
    i = gffuncItemIsInComboBox(ccmbUnit, ccmbUnit.Text)
    If i = -1 Then
        '���뵽�б����
        ccmbUnit.AddItem ccmbUnit.Text
        
        '���ص��������䲾�ļ���
        pobjҵ�����.���չ������䲾.sub���ӵ�λ���� mstr��λ������ & "|" & ccmbUnit.Text
    Else
        '�޸ģ�2001-8-26������λ�����Ų�ͬ���޸Ĺ������䲾����
        If mstr��λ������ <> pobjҵ�����.���չ������䲾.��λ���(ccmbUnit.Text) And mstr��λ������ <> "" Then
            pobjҵ�����.���չ������䲾.sub���ӵ�λ���� mstr��λ������ & "|" & ccmbUnit.Text
        End If
    End If
    Exit Sub
errHandler:
    sfsub������ "ְҵ������", "frmregister", "Sub ccmbUnit_LostFocus", Err.Number, Err.Description, True
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    Select Case KeyCode
    Case vbKeyF8
        If mblnTakePhoto Then
            If cctlCatchPhoto.VideoIsOk Then
                cctlCatchPhoto.subת��״̬
            End If
        End If
    
    End Select
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 39 Then
        KeyAscii = 0
    End If

End Sub

'���ܣ������ʼ����

'���õ�λ��λ
Private Sub ccmd��λ��λ_Click()
    On Error GoTo errHandler
    Dim lobjRec As Object  '��λ��λ���صĽ����¼��
    Dim lobj��λ As Object
    Dim lobj��λ��Ϣ As Object
    '������λ��λ���档
    Set lobjRec = pobjҵ�����.func��λ��λ
    '��ȡ��λ�ĵ�λ����ʾ�ڡ���λ���ơ�¼����С�
    If Not lobjRec Is Nothing Then
        If lobjRec.RecordCount > 0 Then
            ccmbUnit.Text = IIf(IsNull(lobjRec("��λ����")), "", lobjRec("��λ����"))
            mstr��λ������ = lobjRec!������
            'Set lobj��λ = CreateObject("ְҵ������.class1")
            'lobj��λ.��λ��Ϣ���� = lobjRec!������
            'Set lobj��λ��Ϣ���� = lobj��λ.��λ��Ϣ
            
            
            
            If mstr��λ������ <> "" Then
                '�޸ģ�2001-8-23����ʾ��λ���ԣ���
                On Error Resume Next
                'sub��ʾ��λ���� ciptBase, mstr��λ������, mobjGUI
                func��ȡ��λ��Ϣ lobjRec!������
            End If
        End If
    End If
    
    '�ѽ���ص���λ¼��򡣱����ܱ����µ�λ��λ��Ϣ��
    ccmbUnit.SetFocus
    SendKeys vbTab
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����沿��", "FrmRegisterAnnual", "ccmd��λ��λ_Click", 6666, lstrError, False
End Sub

Private Sub mobjGUI_ItemLostFocus(ByVal Index As Integer, ByVal ���� As String, ByVal ���� As String, ByVal �������� As String, ByVal IsError As Boolean)
    On Error GoTo errHandler
    Dim lstrIDCard As String
    Dim i As Integer
    Dim ldatBirth As String
    Dim lstrSex As String
    

    ldatBirth = ""
    Select Case ����
    Case "���֤��"
        lstrIDCard = ciptBase.ItemText(Index)
        If lstrIDCard <> "" Then
            '��ȷʱ�����֤���л�ȡ�������ڡ�
            sub���ݹ�����ݺ����ȡ���պ��Ա� lstrIDCard, ldatBirth, lstrSex
            If Not IsDate(ldatBirth) Then
                Err.Raise 6666, , "���֤�Ų��Ϸ���"
            End If
            
            '�����Ƿ���Ҫ¼��������ڣ���Ҫʱ�Զ��������֤����д��������
            On Error Resume Next
            If IsDate(ldatBirth) Then
                ciptBase.Box1("��������").Text = ldatBirth
                ctxtAge.Text = DateDiff("yyyy", ldatBirth, Date)
            End If
        End If
    Case "��������"
        Dim lstrItemText  As String
        '������ҵ���¼�����ֵ䡣
        For i = 1 To ciptBase.InfoCollection.Count
            If ciptBase.InfoCollection(i).Title = "��ҵ���" Then
                If Not ciptBase.InfoCollection(Index + 1).DictRecordSet Is Nothing Then
                    If ciptBase.InfoCollection(Index + 1).DictRecordSet.EOF Then
                    Else
                        mobjGUI.sub��ʼ���ֵ�� i, "Parent=" & ciptBase.InfoCollection(Index + 1).DictRecordSet("InnerId")
                    End If
                End If
                ciptBase.pblnTemp = True
                lstrItemText = ciptBase.Box1(i - 1).Text
                ciptBase.Box1(i - 1).Text = ""
                ciptBase.Box1(i - 1).Text = lstrItemText
                ciptBase.pblnTemp = False
                
                Exit For
            End If
        Next
    Case "����"
        '��Ч���жϡ�
        If ���� <> "" Then
            If Val(����) > 100 Then
                Err.Raise 6666, , "���䲻�ܴ���100��"
            End If
            If Val(����) >= Val(ctxtAge.Text) Then
                Err.Raise 6666, , "����>=���䣬���ǷǷ������ݣ�"
            End If
        End If
    End Select
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����沿��", "FrmRegisterAnnual", "mobjGUI_ItemLostFocus", 6666, lstrError, False
    
    ciptBase.ItemBox(Index).Text = ""
    ciptBase.ItemSetFocus Index
End Sub

Private Sub cvscLetter_Change()
    On Error Resume Next
    '����������������Ӧ����ĸ��
    If mcolTubeNo.Count > 0 Then
        clblLetter.Caption = mcolTubeNo.Item(cvscLetter.Value)
    End If
End Sub

'���ܣ���ս��档
Private Sub subClear()
    
    On Error Resume Next
    '¼����������пؼ��ָ�formload
'    ccmb���������.Text = ccmb���������.List(0)
'    Ccmb��������.Text = Ccmb��������.List(0)
    '2012-06-14 �ڵ�� ��
    '���������������text��ֵ�󣬱��뽫ListIndexͬʱ����
'    ccmb���������.ListIndex = 0
'    Ccmb��������.ListIndex = 0
    '2012-06-14 �ڵ�� ��
    
    'ccmbTemplate.Clear
       Dim lobjFile As Object
       
    ctxt���֤��.Text = ""
    'cdtpDate.Value = now
    ctxtName.Text = ""
    'ctxtAge = mstrĬ������
    clblsysno.Text = ""
    ccmbSex.Text = ""
    ctxtAge = ""
    ccmbUnit.Text = ""
    ctxtTubeNo = ""
    ctxt��쵥�� = ""
    ctxt����.Text = ""
    ctxt�ʱ�.Text = ""
    ctxtסַ.Text = ""
    ccmb�Ļ��̶�.Text = ccmb�Ļ��̶�.List(0)
    Ccmb���.Text = Ccmb���.List(0)
    ccmb����.Text = ccmb����.List(0)
    ctxt�绰.Text = ""
    ctxt����.Text = ""
    ctxt������.Text = ""
    ccmb����Դ.Text = ccmb����Դ.List(0)
    ccmbְҵ���.Text = ccmb����Դ.List(0)
    ccmbΣ������.Text = ccmbΣ������.List(0)
    ccmb�ֹ���.Text = ccmb�ֹ���.List(0)
    ccmbְ��.Text = ccmbְ��.List(0)
    ctxtΣ������.Text = ""
    ctxt�������.Text = ""
    
    ccmbUnit.Text = ""
    mstr��λ������ = ""
    ctxt������.Text = ""
    ctxt��ϵ�绰.Text = ""
    ccmb��������.Text = ""
    Ccmb��ҵ���.Text = ""
    ctxt��λ��ַ.Text = ""
    '�޸ģ�2002-10-10������ζ����ƣ�������ա�
    Dim ldbl����� As Double
'    ldbl����� = ciptBase.Box1("�����").Text
'    ciptBase.ClearContent
'    ciptBase.Box1("�����").Text = ldbl�����
    clblHistory.Visible = False
    cgrdHistory.rows = 1
    cgrdHistory.Visible = False
    '2012-06-20 �ڵ�� ��
    '�ָ�����Ǽ�ʱ�����ˢ���֤��״̬
    mintState = 0
    If mintState = 0 Then
        If Check���֤.Value = 0 Then
            Check���֤.Value = 1
        Else
            Check���֤_Click
        End If
    End If
    '2012-06-20 �ڵ�� ��
    Picture1.Picture = LoadPicture()
    Picture2.Picture = LoadPicture()
    
    clbl���������.Caption = ""
    Label2(4).Visible = False
    clbl���������.Visible = False
    Set cctlCatchPhoto.Photo = Nothing
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    '2012-07-11 �ڵ�� ��
'''    '�Ѿ���unload֮ǰ��mobjGUI���˳������жϹ���
'''    '����������¼û�б��棬�˻�ϵͳ��š�
'''    If Not mobj��� Is Nothing Then
'''        If mobj���.ϵͳ��� <> "" And Not mobj���.�Ƿ��Ѵ��� Then
'''            '�˻�ϵͳ��š�
'''            mobj���.sub�˻�ϵͳ��� mobj���.ϵͳ���
'''        End If
'''    End If
    '2012-07-11 �ڵ�� ��
    
    mobj����.sub���Ǽ���ֵ "���Ǽ�ʱ¼�뵥λ����", IIf(cchk¼�뵥λ����.Value = 1, "��", "��")
     
    Set mobj��� = Nothing
    Set mobj��켯 = Nothing
    Set mobj����ģ�� = Nothing
    '�ر������
    If mblnTakePhoto Then
        cctlCatchPhoto.subDisconnect
    End If
    mblnInUse = False
    pstrϵͳ������� = ""
frmRegisterManage.subȫ����ʾ
    
    
'    CVR_CloseComm
End Sub



'���ܣ����������ϰ�ť��
Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    Dim i As Integer
    Dim lstr��ˮ�� As String
    Dim lstrϵͳ��� As String
    Dim lcolԭ�����Ŀ As Collection
    Dim lobjrec���� As Object
    Dim lobj������ As Object
    Dim lobjRec As Object
    Dim lstrError As String
    
    '2012-06-13 �ڵ�� ��
    '�洢���֤��Ƭ������ϵͳ����˻ر��������֤�����Ϣ
    Dim lobjRec���֤��Ƭ As Object
    Dim lobjRecϵͳ����˻� As Object
    Dim paraSysNo As String
    Dim lstrSex As String
    Dim lstrBirth As String
    Dim lstrSysNo As String
    '2012-06-13 �ڵ�� ��
    
    On Error GoTo errHandler
    
    Select Case Operate
    
    Case "���"
        
        lstrSysNo = Trim(clblsysno.Text)
        subClear
        clblsysno.Text = lstrSysNo
        '��ս���������ţ���ʾ��¼��������Ա��
        mobj���.�����Ա.����������� = ""
        
        Cancel = True
    
    Case "����"
    
    
        '2012-12-18 ������
        'Bug No:0000025
        '�������Ա����䲻��Ϊ��
        If Trim(ctxtName.Text) = "" Then
            MsgBox "��������Ϊ�գ�", vbExclamation, "ϵͳ��ʾ"
            Exit Sub
        End If
        If Trim(ccmbSex.Text) = "" Then
            MsgBox "�Ա���Ϊ�գ�", vbExclamation, "ϵͳ��ʾ"
            Exit Sub
        End If
        If Trim(ctxtAge.Text) = "" Then
            MsgBox "���䲻��Ϊ�գ�", vbExclamation, "ϵͳ��ʾ"
            Exit Sub
        End If
        '2012-12-18 ������
        '��ʾ���ȡ�
        frmProcess.proPercent.max = 8
        frmProcess.Label1.Caption = "���ڱ��棬��ȴ�..."
        frmProcess.proPercent.Value = 1
        frmProcess.Show 0, Me
        DoEvents
        '2012-08-18 �ڵ�� ��
        '���鱣���������Ϣ����ʡ���жϣ���ʡȥ�жϣ����Ᵽ�档
        '�Ҹ������ʹ���µ�ϵͳ��ţ��ʱ��뽫ԭ���븴����Ϣ�ֿ����档
        If ���ʱ�־ = 2 Then
            frmProcess.proPercent.Value = 2
            frmProcess.proPercent.Value = 3
            frmProcess.proPercent.Value = 4
            frmProcess.proPercent.Value = 5
            frmProcess.proPercent.Value = 6
            frmProcess.proPercent.Value = 7
            frmProcess.proPercent.Value = 8
            Cancel = True
            sub���鱣��
            Unload frmProcess
            MsgBox "����ɹ���"
            Unload Me
            
            Exit Sub
        End If
        frmProcess.proPercent.Value = 2
        DoEvents
        '2012-08-18 �ڵ�� ��
        
        '2012-07-11 �ڵ�� ��
        '���У��ͨ����ֻ�ܴ洢�ֳ���Ƭ
        Cancel = True
        MousePointer = 11
        If mintState = 1 And ���ʱ�־ = 1 Then
            If mblnTakePhoto Then
                Dim strSQL As String
                Dim lobjPhoto As StdPicture
                '����Ƭ��Ϊ�գ��򱣴浽��ӦĿ¼�����õ��� ͨ�ö���.clsͼƬ����.cls
                If Not cctlCatchPhoto.Photo Is Nothing Then
                    Set lobjPhoto = cctlCatchPhoto.Photo
                    pmsub����ͼƬ lobjPhoto, Trim(clblsysno.Text), "ְҵ�����"
                    
                    '2015-10-16
                    strSQL = "update ְҵ�����_��������Ϣ�� set �������='" & Now & "' where ϵͳ���='" & clblsysno & "'"
                    dafuncGetData strSQL
                    
                End If
            End If
            subClear
            ���ʱ�־ = 0
            clblsysno.Text = mobj���.Func����ְҵ�����ϵͳ��� & (ccmb���������.ListIndex + 1)
            Set mcol�����Ŀ = New Collection
           
            mobj���.����.������ = ccmbTemplate.Text
                   
            ctxt����.SetFocus
            frmRegisterManage.sub��ѯ����ʾ
            Cancel = True
            cgrdHistory.rows = 1
            cgrdHistory.Visible = False
            clblHistory.Visible = False
            Check���֤_Click

            GoTo END_SELECT
        End If
        Cancel = False
        '2012-07-11 �ڵ�� ��
        
        '���ϵͳ��ų���С��5�������ж��ǲ���ʧ��
        If Len(clblsysno.Text) < 5 Then
            MousePointer = 0
            MsgBox "ϵͳ��Ŵ������飡", vbInformation, "ϵͳ��ʾ"
            '�޸��ˣ����� 2012.12.20  ����
            '�޸�˵�����˳�����ʱ�رս�������
            Unload frmProcess
            '�޸��ˣ����� 2012.12.20  ����
            Exit Sub
        End If
        
        '2012-07-11 �ڵ�� ��
        'δ�������֤�Ż����֤�Ŵ��󣬲������浱ǰ��Ϣ
        If Len(ctxt���֤��.Text) = 0 Then
            MousePointer = 0
            MsgBox ("δ�������֤�ţ��������浱ǰ���ݣ�")
            '�޸��ˣ����� 2012.12.20  ����
            '�޸�˵�����˳�����ʱ�رս�������
            Unload frmProcess
            '�޸��ˣ����� 2012.12.20  ����
            Exit Sub
        End If
        sub���ݹ�����ݺ����ȡ���պ��Ա� ctxt���֤��.Text, lstrBirth, lstrSex
        If ccmbSex.Text = "" Or lstrSex <> ccmbSex.Text Or Format(lstrBirth, "yyyy-mm-dd") <> Format(cdtp����.Value, "yyyy-mm-dd") Then
            MsgBox ("���֤�뵱ǰ������Ϣ�������������浱ǰ���ݣ�")
            '�޸��ˣ����� 2012.12.20  ����
            '�޸�˵�����˳�����ʱ�رս�������
            Unload frmProcess
            '�޸��ˣ����� 2012.12.20  ����
            Exit Sub
        End If
        '2012-07-11 �ڵ�� ��
        
        '2012-12-18 �����֡�
        'ͬһ���˵����֤���첻������¼��
        'Bug No:0000089
        If dafuncGetData("select ������� from ְҵ�����_���������ݿ� where ������ݺ��� = '" & Trim(ctxt���֤��.Text) & "' and convert(varchar(10),�������,111) = convert(varchar(10),getdate(),111)  and ���״̬ <> 0 ").RecordCount > 0 Then
            MsgBox "ͬһ�����֤�ŵ��첻�ܶ��¼�룡����"
            '2012-06-15 �ڵ�� ��
             'δУ��ʱֻ��ˢ�������֤
             If mintState = 0 Then
                 '�������֤��������ʼ���������֤
                 sub��������ʼ��
                 func���֤��֤
                 mblnTakePhoto = False
                 pobjҵ�����.funcд�뵥�˵�ǰ���״̬ clblsysno.Text, mintState
             ElseIf mintState = 1 Then
                 mobjGUI_BeforeOperate "У��ͨ��", False
             End If
             mintState = 2
             ���ʱ�־ = 1
             '2012-06-15 �ڵ�� ��
             
             '�ָ����ࡣ
             If mblnTakePhoto Then
                 If cctlCatchPhoto.Status = "�ָ�" Then
                     cctlCatchPhoto.subת��״̬
                 End If
             End If
            
            
            lstrSysNo = Trim(clblsysno.Text)
            subClear
            ���ʱ�־ = 0
             clblsysno.Text = lstrSysNo
             Set mcol�����Ŀ = New Collection
             'Set mcol�շ���Ŀ = New Collection
            
             mobj���.����.������ = ccmbTemplate.Text
                    
             
             ctxt����.SetFocus
             'modify by lanchao 2015-03-12 ȡ������2��ע��
             subClear
             frmRegisterManage.sub��ѯ����ʾ
             Unload frmProcess
             Cancel = True
             MousePointer = 0
             If Check���֤.Value = 1 Then
                Timer2.Enabled = True
             End If
             Exit Sub
        End If
        
        '2012-12-18 ������ ��
        
        '���� ϵͳ��� �ı��� ����
        ctxt����.SetFocus
        '�ж��Ƿ���Ҫ���ࡣ
        If mblnTakePhoto = True Then
            '�ж��Ƿ�����
            If cctlCatchPhoto.Photo Is Nothing Then
                Err.Raise 6666, , "���ڡ�ҵ�����á���������Ҫ���񣬵���������û�����࣬�޷����档����취��" & Chr(13) & Chr(10) & "��1�� �밴��ȡ�񡱰�ť����󱣴棡" & Chr(13) & Chr(10) & "��2�����㲻׼�����࣬���Ƚ��롰ҵ�����á����ò����ࡣ"
                Unload frmProcess
            End If
        End If
        
        '2012-06-13 �ڵ�� ��
        'ʡ������Ҫ��������Ƭ�����֤��Ƭ�ֿ��洢���ֿ���ʾ
        '���ﵥ���洢���֤��Ƭ��������Ƭ��ԭ�������洢��
        Set lobjRec���֤��Ƭ = CreateObject("ְҵ������.clsPersonExamed")
        lobjRec���֤��Ƭ.func�������֤��Ƭ Picture2.Image, clblsysno.Text & "IDcard", "ְҵ�����"
        Set lobjRec���֤��Ƭ = Nothing
        '2012-06-13 �ڵ�� ��
        frmProcess.proPercent.Value = 3
        DoEvents
        '�����ǿ���¼�룬���¼���Ƿ��д���
'        If mobj���.����.������Ϣ.Count > 0 Then
'            '�޸ģ�2001-9-12�������
'            On Error Resume Next
'            ciptBase.Box1(ciptBase.ActiveInputBoxIndex).LostFocus
'            On Error GoTo errHandler
'
'            If ciptBase.ItemsError.Count > 0 And Not mbln����¼�� Then
'                Err.Raise 6666, , "�������ɫ¼������ݣ�"
'            End If
'        End If
        MousePointer = 11
        
        Set lobj������ = CreateObject("ְҵ������.clsmedicalexamsheet")
        'lobj������.������ = ccmbTemplate.ItemData(ccmbTemplate.ListIndex(1))
        lobj������.������ = ccmbTemplate.Text
        'Set lobjrec���� = CreateObject("ְҵ������.clsmedicalexam")
        'lobjrec1.������� = ccmb���������.ItemData(ccmb���������.ListIndex)
        'lobjrec1.������ = ccmb��������.ItemData(ccmb��������.ListIndex)
        
        '�����Թܱ�Ų�����
        With mobj���
            '2012-06-14 �ڵ�� ��
            'ϵͳ��ű������������¸�ֵ�������һ�λ�ʹ��form_loadʱ��ϵͳ���
            .ϵͳ��� = Trim(clblsysno.Text)
            '2012-06-14 �ڵ�� ��
            
            If .����.������ <> ccmbTemplate.Text Then
                .����.������ = ccmbTemplate.Text
            End If
            '�޸ģ�2004-1-9���Թܱ�ſ������룩
'            If pobjҵ�����.ҵ������("�Թܱ���Զ�����") = "��" Then
'                If .����.�Թܱ����ĸ <> clblLetter.Caption Then
'                    .����.�Թܱ����ĸ = clblLetter.Caption
'                End If
'            Else
''                .����.�Թܱ����ĸ = clblLetter.Caption
'                .�Թܱ�� = ctxtTubeNo.Text
'            End If

              '����Ա������ 2015-12-25 by Ĳ��
'            mobj���.�����Ա.�Ա� = Trim(ccmbSex.Text)
            mobj���.�����Ա.���� = Trim(ctxtAge.Text)

            .�����Ա.ϵͳ��� = Trim(clblsysno.Text)
            .�����Ա.���� = ctxtName
            .�����Ա.�Ա� = ccmbSex.Text
            .�����Ա.��λ���� = ccmbUnit.Text
            .�����Ա.Σ������ = ccmbΣ������.Text
            .�����Ա.����Դ = Right(ccmb����Դ.Text, 2)
            .�����Ա.ְҵ���� = ccmbְҵ���.Text
            .�����Ա.�ֹ��� = ccmb�ֹ���.Text
            .�����Ա.ְ���ְ�� = ccmbְ��.Text
            .�����Ա.ְҵΣ������ = ctxtΣ������.Text
            .�����Ա.������� = Trim(ctxt�������.Text)
            .�����Ա.���� = Trim(ctxt����.Text)
            .�����Ա.�ʱ� = Trim(ctxt�ʱ�.Text)
            .�����Ա.סַ = Trim(ctxtסַ.Text)
            .�����Ա.��� = Ccmb���.Text
            .�����Ա.�绰���� = Trim(ctxt�绰.Text)
            .�����Ա.���� = Trim(ctxt����.Text)
            .�����Ա.������ = Trim(ctxt������.Text)
            .�����Ա.������ = ctxt������.Text
            .�����Ա.��ϵ�绰 = ctxt��ϵ�绰.Text
            .�����Ա.�������� = ccmb��������.Text
            .�����Ա.��ҵ��� = Ccmb��ҵ���.Text
            .�����Ա.��λ��ַ = ctxt��λ��ַ.Text
            If mblnTakePhoto Then
                .�����Ա.��Ƭ = cctlCatchPhoto.Photo
'                .�����Ա.��Ƭѹ�� = cctlCatchPhoto.Photo
           ' ElseIf Not Picture1.Picture Is Nothing Then
             '   .�����Ա.��Ƭ = Picture1.Picture
            '2015-3-13
            ElseIf Not Picture2.Picture Is Nothing Then
                .�����Ա.��Ƭ = Picture2.Picture
        
            End If
            If Val(ctxtAge.Text) > 0 Then
'                If Val(ctxtAge.Text) > 200 Then
'                    Err.Raise 6666, , "���䳬��ϵͳ������������200��"
'                End If
                .�����Ա.�������� = DateAdd("yyyy", -Val(ctxtAge.Text), Date)
            Else
                '��������ַ������������䡣
                mobj����.sub���Ǽ���ֵ "�������", ctxtAge.Text
                mstrĬ������ = ctxtAge.Text
            End If
            .�����Ա.���� = ctxtAge.Text
            
            On Error Resume Next
            .�����Ա.������ݺ��� = ctxt���֤��.Text
            '.�����Ա.�������� = ciptBase.Box1("��������").TrueText
            '.�����Ա.��ҵ��� = ciptBase.Box1("��ҵ���").TrueText
            '.�����Ա.Ƭ�� = ciptBase.Box1("Ƭ��").TrueText
            .�����Ա.�Ļ��̶� = Left(ccmb�Ļ��̶�.Text, 2)
            
            .�����Ա.���� = ccmb����.Text
            If ccmbUnit.Text = "" Then
                .�����Ա.��λ������ = ""
            Else
                If .�����Ա.��λ������ <> mstr��λ������ Then
                    '����λ������¸�ֵ���������»�ȡ���������ࡢ��ҵ���Ƭ����
                    .�����Ա.��λ������ = mstr��λ������
                End If
            End If
            
            '���渽����Ϣ
            'For i = 1 To ciptBase.ItemCount
                'If ciptBase.Box1(i - 1).TrueText <> ciptBase.Box1(i - 1).Text And ciptBase.Box1(i - 1).Text <> "" Then
             '   If ciptBase.InfoCollection(i).�ֵ����� <> "" And ciptBase.Box1(i - 1).TrueText <> "" Then
             '       .����.Sub�����Ϣֵ ciptBase.InfoCollection(i).����, ciptBase.Box1(i - 1).TrueText & " " & ciptBase.Box1(i - 1).Text
             '   Else
             '       .����.Sub�����Ϣֵ ciptBase.InfoCollection.Item(i).Title, ciptBase.ItemText(i - 1)
             '   End If
           ' Next i
            
            '����Ϊ������
            'If ccmb���������.Text = "����" Then
            '    .������ = P_EXAM_FIRST
            'Else
            '    .������ = P_EXAM_ANNUAL
            'End If
            .������� = cdtpDate.Value ' ,Format(cdtpDate.Value, "yyyy-mm-dd hh:mm:ss")
            
            '�޸ģ�2004-1-9��������쵥�ţ�
            '.��쵥�� = ctxt��쵥��.Text
            .��������� = ccmb���������.Text
            .�������� = Ccmb��������.Text
            '.��������� = ccmb���������.ItemData(ccmb���������.ListIndex)
            '.�������� = Ccmb��������.ItemData(Ccmb��������.ListIndex)
            
            On Error GoTo errHandler
            Set lobj������ = CreateObject("ְҵ������.clsmedicalexamsheet")
            lobj������.������ = Trim(ccmbTemplate.Text)
            If mcol�����Ŀ.Count = 0 Then
                Set mcol�����Ŀ = lobj������.������Ŀ��("")
            End If
            Set .col�����Ŀ = mcol�����Ŀ
        
        End With
        
        '���ܣ����������Ŀ
        'ʱ�䣺2012-06-04
        '���ߣ�����
        save�Ż��������Ŀ mcol�����Ŀ, Trim(clblsysno.Text)
        frmProcess.proPercent.Value = 4
        DoEvents

        
      If mcol�շ���Ŀ.Count > 0 Then
            pobjҵ�����.Sub���Ǽ� mobj���, , , mcol�շ���Ŀ, Val(ctxt����)
        Else
            pobjҵ�����.Sub���Ǽ� mobj���, , , , Val(ctxt����)
        End If
        frmProcess.proPercent.Value = 5
        DoEvents
        Set lobjRec = CreateObject("ְҵ������.clsMoney")
        lobjRec.mstrϵͳ��� = Trim(clblsysno.Text)
        lobjRec.mstr�����Ա���� = ctxtName.Text
        Set lobjRec.col�����Ŀ = mcol�����Ŀ
        Dim lstr�շ����� As String
        lstrError = lobjRec.func�շ�(lstr�շ�����)
        mobj���.�շ����� = lstr�շ�����
        If lstrError <> "" And lstrError <> "Cancel" Then
            MsgBox lstrError, vbOKOnly + vbExclamation, "ϵͳ��ʾ"
        End If
    
        cstbMain.Panels(1) = "�ϴα�������ϵͳ��ţ�" & mobj���.ϵͳ��� '& " ���Թܱ�ţ�" & mobj���.�Թܱ��
        If mobj���.�շ����� <> "" Then
            cstbMain.Panels(1) = cstbMain.Panels(1) & "���շ����ţ�" & mobj���.�շ�����
        End If
        frmProcess.proPercent.Value = 6
        DoEvents
        '2012-06-25 �ڵ�� ��
        '��ʼ����������Ϣ���С��������״̬���ֶ�
        subInit�������״̬ mcol�����Ŀ, Trim(clblsysno.Text)
        '2012-06-25 �ڵ�� ��
        frmProcess.proPercent.Value = 7
        DoEvents
        '2012-06-15 �ڵ�� ��
        'δУ��ʱֻ��ˢ�������֤
        If mintState = 0 Then
            '�������֤��������ʼ���������֤
'            sub��������ʼ��
'            func���֤��֤
            mblnTakePhoto = False
            pobjҵ�����.funcд�뵥�˵�ǰ���״̬ clblsysno.Text, mintState
        ElseIf mintState = 1 Then
            mobjGUI_BeforeOperate "У��ͨ��", False
        End If
        mintState = 2
        ���ʱ�־ = 1
        '2012-06-15 �ڵ�� ��
        
        '�ָ����ࡣ
        If mblnTakePhoto Then
            If cctlCatchPhoto.Status = "�ָ�" Then
                cctlCatchPhoto.subת��״̬
            End If
        End If
        '2012-11-08 ������
        '�Ǽ����Ҫˢ��
'        If cchkClear = 1 Then
            subClear
            ���ʱ�־ = 0
            clblsysno.Text = mobj���.Func����ְҵ�����ϵͳ��� & (ccmb���������.ListIndex + 1)
'        End If
        
        Set mcol�����Ŀ = New Collection
        'Set mcol�շ���Ŀ = New Collection
       
        mobj���.����.������ = ccmbTemplate.Text
               
        '�Թ���ĸ������ѡ��
'        cvscLetter.Enabled = False
        ctxt����.SetFocus
        frmRegisterManage.sub��ѯ����ʾ
        Cancel = True
        frmProcess.proPercent.Value = 8
        cgrdHistory.rows = 1
        cgrdHistory.Visible = False
        clblHistory.Visible = False
        
        '��Ӵ�ӡ��ǩ����
        Dim strsql1 As String
        strsql1 = "select distinct left(�����Ŀ,2) as ��Ŀ  from ְҵ�����_����ģ�������Ŀ�� where ��������='" & ccmbTemplate.Text & "'"
        Dim objds1 As Object
        Set objds1 = dafuncGetData(strsql1)
        Dim lobjFile As Object
        Set lobjFile = CreateObject("ְҵ������.cls����")
        Dim csysno As Collection
        Set csysno = New Collection
        
        csysno.Add (mobj���.ϵͳ���)
        
        lobjFile.func��ӡ����嵥 csysno ''��ͣ��ӡ��ǩ���� 2015-9-1 by lanchao
        '  Set lobjFile = Nothing
        '  '���ĵ�ǰ���״̬����ӡ�嵥֮�󣬾ͽ������״̬��
        ''  pobjҵ�����.funcд�뵥�˵�ǰ���״̬ mobj���.ϵͳ���, 2
        Dim c As Integer
           c = objds1.RecordCount
        objds1.MoveFirst
        For i = 0 To c - 1
        If objds1("��Ŀ") = "01" Then  '01 ��ٿ�
'           lobjFile.func��ӡ�Թܱ�ǩ mobj���.ϵͳ���, mobj���.�����Ա.����, "����.Ѫ��"
            lobjFile.func��ӡ�Թܱ�ǩ mobj���.ϵͳ���, mobj���.�����Ա.����, mobj���.�����Ա.�Ա�, mobj���.�����Ա.����, "��ٿ�"
        End If
        If objds1("��Ŀ") = "02" Then  '02 �ڿ�
'           lobjFile.func��ӡ�Թܱ�ǩ mobj���.ϵͳ���, mobj���.�����Ա.����, "����.Ѫ��"
            lobjFile.func��ӡ�Թܱ�ǩ mobj���.ϵͳ���, mobj���.�����Ա.����, mobj���.�����Ա.�Ա�, mobj���.�����Ա.����, "�ڿ�"
        End If
        If objds1("��Ŀ") = "03" Then  '03 ���
'           lobjFile.func��ӡ�Թܱ�ǩ mobj���.ϵͳ���, mobj���.�����Ա.����, "����.Ѫ��"
            lobjFile.func��ӡ�Թܱ�ǩ mobj���.ϵͳ���, mobj���.�����Ա.����, mobj���.�����Ա.�Ա�, mobj���.�����Ա.����, "���"
        End If
        If objds1("��Ŀ") = "08" Then  '08 �����
'           lobjFile.func��ӡ�Թܱ�ǩ mobj���.ϵͳ���, mobj���.�����Ա.����, "����.Ѫ��"
            lobjFile.func��ӡ�Թܱ�ǩ mobj���.ϵͳ���, mobj���.�����Ա.����, mobj���.�����Ա.�Ա�, mobj���.�����Ա.����, "�����"
        End If
        If objds1("��Ŀ") = "09" Then  '09 X��
'           lobjFile.func��ӡ�Թܱ�ǩ mobj���.ϵͳ���, mobj���.�����Ա.����, "����.Ѫ��"
            lobjFile.func��ӡ�Թܱ�ǩ mobj���.ϵͳ���, mobj���.�����Ա.����, mobj���.�����Ա.�Ա�, mobj���.�����Ա.����, "X����"
        End If
        If objds1("��Ŀ") = "10" Then  '10 �ĵ�
'           lobjFile.func��ӡ�Թܱ�ǩ mobj���.ϵͳ���, mobj���.�����Ա.����, "����.Ѫ��"
            lobjFile.func��ӡ�Թܱ�ǩ mobj���.ϵͳ���, mobj���.�����Ա.����, mobj���.�����Ա.�Ա�, mobj���.�����Ա.����, "�ĵ�"
        End If
        If objds1("��Ŀ") = "11" Then  '11 B��
'           lobjFile.func��ӡ�Թܱ�ǩ mobj���.ϵͳ���, mobj���.�����Ա.����, "����.Ѫ��"
            lobjFile.func��ӡ�Թܱ�ǩ mobj���.ϵͳ���, mobj���.�����Ա.����, mobj���.�����Ա.�Ա�, mobj���.�����Ա.����, "B��Ӱ���"
        End If
        If objds1("��Ŀ") = "12" Then  '12 �ι���
'           lobjFile.func��ӡ�Թܱ�ǩ mobj���.ϵͳ���, mobj���.�����Ա.����, "����.Ѫ��"
            lobjFile.func��ӡ�Թܱ�ǩ mobj���.ϵͳ���, mobj���.�����Ա.����, mobj���.�����Ա.�Ա�, mobj���.�����Ա.����, "�ι���"
        End If
        If objds1("��Ŀ") = "05" Then  '05 ����Ѫ��
'           lobjFile.func��ӡ�Թܱ�ǩ mobj���.ϵͳ���, mobj���.�����Ա.����, "����.Ѫ��"
            lobjFile.func��ӡ�Թܱ�ǩ mobj���.ϵͳ���, mobj���.�����Ա.����, mobj���.�����Ա.�Ա�, mobj���.�����Ա.����, "����.Ѫ��"
        End If
        If objds1("��Ŀ") = "06" Then  '06 �򳣹�
'           lobjFile.func��ӡ�Թܱ�ǩ mobj���.ϵͳ���, mobj���.�����Ա.����, "�򳣹�"
            lobjFile.func��ӡ�Թܱ�ǩ mobj���.ϵͳ���, mobj���.�����Ա.����, mobj���.�����Ա.�Ա�, mobj���.�����Ա.����, "�򳣹�"
        End If
        If objds1("��Ŀ") = "07" Then  '07 Ⱦɫ��
'           lobjFile.func��ӡ�Թܱ�ǩ mobj���.ϵͳ���, mobj���.�����Ա.����, "Ⱦɫ��"
            lobjFile.func��ӡ�Թܱ�ǩ mobj���.ϵͳ���, mobj���.�����Ա.����, mobj���.�����Ա.�Ա�, mobj���.�����Ա.����, "Ⱦɫ��"
        End If
        If objds1("��Ŀ") = "04" Then  '04 Ѫ����
'           lobjFile.func��ӡ�Թܱ�ǩ mobj���.ϵͳ���, mobj���.�����Ա.����, "Ѫ����.����Ѫ"
            lobjFile.func��ӡ�Թܱ�ǩ mobj���.ϵͳ���, mobj���.�����Ա.����, mobj���.�����Ա.�Ա�, mobj���.�����Ա.����, "Ѫ����.����Ѫ"
        End If
        If objds1("��Ŀ") = "17" Then  '17 ����
            '����������ֱ�Ӵ�ӡ���������Ǵ�ӡ������Ŀ  2015-12-10 by Ĳ��   ��
            Dim lobject As Object
'            Set lobject = dafuncGetData("select distinct right(�����Ŀ,2) as ��Ŀ  from ְҵ�����_����ģ�������Ŀ�� where ��������='" & ccmbTemplate.Text & "'and �����Ŀ like '1702%'and �����Ŀ<>'17020'")
            Set lobject = dafuncGetData("select distinct right(�����Ŀ,2) as ��Ŀ���,���� as ��Ŀ  from ְҵ�����_����ģ�������Ŀ�� a,ְҵ�����_�����Ŀ���ñ� b where a.��������='" & ccmbTemplate.Text & "'and a.�����Ŀ like '1702%'and a.�����Ŀ=b.���� and a.�����Ŀ<>'17020' and b.����='����'")
            If lobject.RecordCount > 0 Then
'                If lobject("��Ŀ") = "21" Then
'                lobjFile.func��ӡ�Թܱ�ǩ mobj���.ϵͳ���, mobj���.�����Ա.����, "�ι�1,����,���԰�,GLU,Ѫ֬,ACP"
'                ElseIf lobject("��Ŀ") = "22" Then
'                lobjFile.func��ӡ�Թܱ�ǩ mobj���.ϵͳ���, mobj���.�����Ա.����, "�ι�2,����,GLU,Ѫ֬,ACP"
'                Else
'                lobjFile.func��ӡ�Թܱ�ǩ mobj���.ϵͳ���, mobj���.�����Ա.����, "����"
'                End If
                Dim xiangmu As String
                Dim zhongjian As String
                Dim X As Integer
                lobject.MoveFirst
                For X = 0 To lobject.RecordCount - 1
                zhongjian = zhongjian + "," + lobject("��Ŀ")
                lobject.MoveNext
                Next X
                xiangmu = Right(zhongjian, Len(zhongjian) - 1)
'                lobjFile.func��ӡ�Թܱ�ǩ mobj���.ϵͳ���, mobj���.�����Ա.����, xiangmu
                 lobjFile.func��ӡ�Թܱ�ǩ mobj���.ϵͳ���, mobj���.�����Ա.����, mobj���.�����Ա.�Ա�, mobj���.�����Ա.����, xiangmu
            End If
            '2015-12-10 by  Ĳ��  ��
'           lobjFile.func��ӡ�Թܱ�ǩ mobj���.ϵͳ���, mobj���.�����Ա.����, "����"
        End If
        objds1.MoveNext
        Next i
        Unload frmProcess
        MousePointer = 0
        'Update ���״̬
'        Dim strSQL As String
        strSQL = "update ְҵ�����_��������Ϣ�� set ���״̬= '2'  where ϵͳ���='" & mobj���.ϵͳ��� & "'"
        dafuncGetData (strSQL)

        'sub��������ʼ��
        'Timer2.Enabled = True
        '2015-3-15ֹͣ���� ��ΰ
        Dim ret
        ret = CloseComm()
        
        frmRegisterManage.sub��ѯ����ʾ
        Unload Me
'        Exit Sub   '2016-2-24 by Ĳ��
    Case "�����Ŀ"
        Dim lobj���ģ�� As Object
        
        '��ȡ���������е������Ŀ��
        Set lcolԭ�����Ŀ = mobj���.����.�����Ŀ��("")
        
        '����ѡ����Ŀ��������ԡ�
        frmSelectItem.pstr�������� = ccmbTemplate.Text
        Set frmSelectItem.pcol������Ŀ = lcolԭ�����Ŀ
        Set frmSelectItem.pcol�շ���Ŀ = mcol�շ���Ŀ
        
        '����ѡ����Ŀ���档
        frmSelectItem.Show 1
        If frmSelectItem.pblnOk Then
            '��ȡѡ�еĸ�����Ŀ��
            Set mcol�����Ŀ = frmSelectItem.pcol������Ŀ
            '��ȡ���õ��շ���Ŀ��
            'Set mcol�շ���Ŀ = frmSelectItem.pcol�շ���Ŀ
            
            '��ʾ�շѽ�
            Dim ldblTotal As Double
            'For i = 1 To mcol�շ���Ŀ.Count
            '    ldblTotal = Format(ldblTotal + mcol�շ���Ŀ(i)("����"), "0.00")
            'Next
            On Error Resume Next
            If sffunc�жϼ��ϼ�ֵ�Ƿ����(mobj���.����.������Ϣ, "�����") Then
                ciptBase.Box1("�����").Text = ldblTotal
                mobj���.����.Sub�����Ϣֵ "�����", ldblTotal
            End If
            
        End If
    Case "������Ƭ"
        Dim lstrFile As String
        ccmdFile.Filter = "BMP|*.bmp|JPG|*.jpg"
        If Dir(App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "��Ƭ", vbDirectory) <> "" Then
            ccmdFile.InitDir = App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "��Ƭ"
        End If
        ccmdFile.FileName = Trim(clblsysno.Text)
        ccmdFile.ShowOpen
        lstrFile = ccmdFile.FileName
        If lstrFile <> "" Then
            If InStr(lstrFile, ".") > 0 Then
                Set cctlCatchPhoto.Photo = LoadPicture(lstrFile)
                mblnTakePhoto = True
            End If
        End If
    Case "�޸�"
        
        '��ȡ�������ĺš�
        If Val(Right(Trim(clblsysno.Text), Len(Trim(clblsysno.Text)) - Len(mobj���.ϵͳ��Ź̶�����))) > 1 Then
            FrmEditRegister.ϵͳ��� = mobj���.ϵͳ��Ź̶����� & Format(Val(Right(Trim(clblsysno.Text), Len(Trim(clblsysno.Text)) - Len(mobj���.ϵͳ��Ź̶�����))) - 1, String(Len(Trim(clblsysno.Text)) - Len(mobj���.ϵͳ��Ź̶�����), "0"))
        Else
            FrmEditRegister.ϵͳ��� = ""
        End If
        FrmEditRegister.Show 1, Me
    '2012-06-20 �ڵ�� ��
    '���У��ͨ�����ж�
    Case "У��ͨ��"
        If Trim(ctxtName.Text) = "" Then
            MsgBox "��������Ϊ�գ�", vbExclamation, "ϵͳ��ʾ"
            Exit Sub
        End If
        If Trim(ccmbSex.Text) = "" Then
            MsgBox "�ձ���Ϊ�գ�", vbExclamation, "ϵͳ��ʾ"
            Exit Sub
        End If
        If Trim(ctxtAge.Text) = "" Then
            MsgBox "���䲻��Ϊ�գ�", vbExclamation, "ϵͳ��ʾ"
            Exit Sub
        End If
        '2012-12-18 ������
        'Bug No:0000026
        sub���ݹ�����ݺ����ȡ���պ��Ա� ctxt���֤��.Text, lstrBirth, lstrSex
        If ccmbSex.Text = "" Or lstrSex <> ccmbSex.Text Or Format(lstrBirth, "yyyy-mm-dd") <> Format(cdtp����.Value, "yyyy-mm-dd") Then
            MsgBox ("���֤�뵱ǰ������Ϣ�������������浱ǰ���ݣ�")
            Exit Sub
        End If
        '2012-12-18 ������
            mintState = 1
            If Check���֤.Value = 1 Then
                Check���֤.Value = 0
            Else
                Check���֤_Click
            End If
            If mblnTakePhoto Then
                If cctlCatchPhoto.Status = "�ָ�" Then
                    cctlCatchPhoto.subת��״̬
                End If
            End If
            If ���ʱ�־ = 2 And pstr����ϵͳ��� <> "" Then
                pobjҵ�����.funcд�뵥�˵�ǰ���״̬ pstr����ϵͳ���, mintState
                pobjҵ�����.funcд��У������Ϣ pstr����ϵͳ���, um�û����
            Else
                pobjҵ�����.funcд�뵥�˵�ǰ���״̬ clblsysno.Text, mintState
                pobjҵ�����.funcд��У������Ϣ clblsysno.Text, um�û����
            End If
            
    '2012-06-20 �ڵ�� ��
    Case "�˳�"
        '2012-06-14 �ڵ�� ��
        '�˳�ʱ���ϵͳ��ţ����Ƿ��˻ء�
        paraSysNo = Left(clblsysno.Text, Len(clblsysno.Text) - 1)
        Set lobjRecϵͳ����˻� = dafuncGetData("select * from ְҵ�����_���������ݿ� where left(ϵͳ���,len(ϵͳ���)-1)='" & paraSysNo & "'")
        '2012-07-11 �ڵ�� ��
        '�˳�ʱ��ѯ����Ϣ�Ƿ��ѱ��棨д�ýϼ򵥣�
        i = -vbYes
        If Len(ctxt���֤��.Text) > 0 And Len(ctxtName.Text) > 0 And mintState <> 2 Then i = MsgBox("���浱ǰ��Ϣ��", vbYesNo) 'i = vbYes
        If i = vbYes Then
'            mintState = 0
            mobjGUI_BeforeOperate "����", True
        Else
        '2012-07-11 �ڵ�� ��
            If lobjRecϵͳ����˻�.RecordCount = 0 Then
                Set lobjRecϵͳ����˻� = CreateObject("ְҵ������.clsMedicalExam")
                lobjRecϵͳ����˻�.Func�˻�ְҵ�����ϵͳ��� paraSysNo
            End If
        End If
        Set lobjRecϵͳ����˻� = Nothing
        Unload Me
        '2012-06-14 �ڵ�� ��
    End Select
    Exit Sub   '2016-2-24 by Ĳ��
END_SELECT:
    Set lobjrec���� = Nothing
    Set lobj������ = Nothing
    '2012-11-08 ������
    '�����ֱ���˳�
'    paraSysNo = Left(clblSysNo.Text, Len(clblSysNo.Text) - 1)
'    Set lobjRecϵͳ����˻� = dafuncGetData("select * from ְҵ�����_���������ݿ� where left(ϵͳ���,len(ϵͳ���)-1)='" & paraSysNo & "'")
'    i = -vbYes
'    If Len(ctxt���֤��.Text) > 0 And Len(ctxtName.Text) > 0 And mintState <> 2 Then i = vbYes 'i = MsgBox("���浱ǰ�����", vbYesNo)
'    If i = vbYes Then
'        mobjGUI_BeforeOperate "����", True
'    Else
'
'        If lobjRecϵͳ����˻�.RecordCount = 0 Then
'            Set lobjRecϵͳ����˻� = CreateObject("ְҵ������.clsMedicalExam")
'            lobjRecϵͳ����˻�.Func�˻�ְҵ�����ϵͳ��� paraSysNo
'        End If
'    End If
'    Set lobjRecϵͳ����˻� = Nothing
'    Unload Me
    MousePointer = 0
    Unload frmProcess
    Exit Sub
errHandler:
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����沿��", "FrmRegisterAnnual", "mobjGUI_BeforeOperate", 6666, lstrError, False
    
    MousePointer = 0
    cstbMain.Panels(1) = lstrError
    Cancel = True
    Exit Sub
    Resume
    Exit Sub
End Sub


'���ܣ���ʾָ��ϵͳ��ŵ������Ա����Ϣ�ڽ����ϡ�
Private Sub SubGetPersonInfo(ByVal paraϵͳ��� As String)
    Dim lcolInfo As New Collection
    Dim i As Integer
    Dim j As Integer
    Dim lstrTemp As String
    Dim lstrTubeNo As String
    Dim lstrSysNo As String
    
    
    On Error GoTo errHandler
    MousePointer = 11
    
    '������ʱ���ɲ�����
    ctbMain.Enabled = False
    
    '���˻ؾ�ϵͳ��š�
    If Not mobj���.�Ƿ��Ѵ��� And mobj���.ϵͳ��� <> "" Then
        mobj���.sub�˻�ϵͳ��� mobj���.ϵͳ���
    End If
    
    '������ְҵ������
    Set mobj����� = CreateObject("ְҵ������.clsMedicalExam")
    mobj�����.ϵͳ��� = paraϵͳ���
    
    '����ϴ���������
    If ccmbTemplate.Text <> mobj�����.����.������ Then
        ccmbTemplate.Text = mobj�����.����.������
    
        '���³�ʼ��¼��塣
        On Error Resume Next
        mobjGUI.sub��ʼ��¼��� mobj�����.����.������
        On Error GoTo errHandler
    End If
    
    '��ȡ������¼�ĸ�����Ϣ��
    Set lcolInfo = mobj�����.����.������Ϣ
    
    '��д������Ϣֵ
    sub��¼���ֵ ciptBase, mobjGUI, lcolInfo
    
    '��ʾ������Ϣ��
    With mobj�����.�����Ա
        ctxtName.Text = .����
        ccmbSex.Text = .�Ա�
        ctxtAge.Text = .����
        ccmbUnit.Text = .��λ����
        ccmbUnit_LostFocus
        
        '��Ƭ
        '��ò���ʾ��Ƭ��
        If Not .��Ƭ Is Nothing Then
            Set cctlCatchPhoto.Photo = .��Ƭ
        Else
            cctlCatchPhoto.subClear
        End If
        
        '�޸ģ�2001-8-23��
        On Error Resume Next
        mstr��λ������ = .��λ������
        
        On Error GoTo errHandler
    End With
    
    '�޸ģ�2001-12-30����ʾ�ϴ�������ڣ���
    Label2(4).Visible = True
    clbl���������.Visible = True
    clbl���������.Caption = mobj�����.�������
    
    '�޸ģ�2002-1-6����ʱ��������18���£��Զ�����Ϊ���죩��
    'If IsDate(clbl���������.Caption) Then
    '    If DateDiff("m", clbl���������.Caption, Now) >= 18 Then
    '        ccmb���������.ListIndex = 0
    '    Else
            '����18���£��Զ�����Ϊ��졣
    '        ccmb���������.ListIndex = 1
    '    End If
    'End If
    '�����µ�ϵͳ���
    lstrSysNo = mobj���.Func����ϵͳ���
    mobj���.ϵͳ��� = lstrSysNo
    clblsysno.Text = lstrSysNo
    
    '�����������䡣
    mobj���.�����Ա.����������� = mobj�����.�����Ա.�����������
    
    
    '�����������������Ӷ���ȡ���Թܱ�š�
    mobj���.����.������ = ccmbTemplate.Text
    
    If pobjҵ�����.ҵ������("�Թܱ���Զ�����") = "��" Then
        '��ȡ������ĵ�����ʹ�õ��Թܱ����ĸ��
        clblLetter.Caption = mobj���.����.�Թܱ����ĸ
        If clblLetter.Caption = "" Then
            
            '�ô����Ǽ��ǵ���ĵ�һ����������ģ������л�ȡ���п�ѡ����Ļ��
            mobj����ģ��.������ = ccmbTemplate.Text
            lstrTubeNo = mobj����ģ��.�Թ���ĸ���
            
            '����ĸ�����ŷֿ�������mcoltubeNo�С�
            If Right(lstrTubeNo, 1) <> "," Then lstrTubeNo = lstrTubeNo & ","
            lstrTemp = ""
            Set mcolTubeNo = New Collection
            For i = 1 To Len(lstrTubeNo)
                lstrTemp = lstrTemp & Mid(lstrTubeNo, i, 1)
                If Mid(lstrTubeNo, i, 1) = "," Then
                    If Left(lstrTemp, Len(lstrTemp) - 1) <> "" Then
                        mcolTubeNo.Add Left(lstrTemp, Len(lstrTemp) - 1)
                    End If
                    lstrTemp = ""
                End If
            Next i
            If mcolTubeNo.Count > 0 Then
                '�Թ���ĸ�ı��ˣ�������ʾ��
                If clblLetter.Caption <> "" And clblLetter.Caption <> mcolTubeNo(1) Then
                    sffuncMsg "��ע�⣬������ѡ�������ʹ�õ��Թ���ĸ��ǰһ����" & clblLetter.Caption & "����ͬ�ˡ�"
                End If
            
                '��ֵ��clblLetter��
                clblLetter.Caption = mcolTubeNo(1)
                '��ĸ����ѡ��
                cvscLetter.Enabled = True
                cvscLetter.Min = 1
                cvscLetter.max = mcolTubeNo.Count
                cvscLetter.Value = 1
            Else
                ctbMain.Buttons(6).Enabled = False
                '��ʾ�������޿��õ���ĸ��
                Err.Raise 6666, , "�������޿����Թ���ĸ��ţ��������������Ӧ���Թ���ĸ���"
            End If
        Else
            '����ĸ������ѡ����ĸ��
            cvscLetter.Enabled = False
        End If
    Else
        ctxtTubeNo = mobj���.�Թܱ��
    End If
    '���水ť���á�
    ctbMain.Buttons(6).Enabled = True
    Err.Clear
    
errHandler:
    '�ָ�����ɲ�����
    ctbMain.Enabled = True
    MousePointer = 0
    If Err <> 0 Then
        sfsub������ "ְҵ�����沿��", "FrmRegisterAnnual", "SubGetPersonInfo", Err.Number, Err.Description, True
    End If
    
    Exit Sub
    Resume
End Sub


'��ȡ����Դ�ֵ�
Private Sub func����Դ()
    Dim lobj����Դ As Object
    Dim lcol����Դ As Object
    Dim i As Integer
    On Error GoTo errHandler
    ccmb����Դ.Visible = True
    Set lobj����Դ = CreateObject("ְҵ������.clsmedicalexamtemplateset")
    Set lcol����Դ = lobj����Դ.����Դ
    ccmb����Դ.Clear
    ccmb����Դ.AddItem ""
    For i = 1 To lcol����Դ.RecordCount
        ccmb����Դ.AddItem lcol����Դ("����")
        ccmb����Դ.ItemData(ccmb����Դ.NewIndex) = lcol����Դ("���")
        lcol����Դ.MoveNext
    Next
    ccmb����Դ.Text = ccmb����Դ.List(0)
    Set lobj����Դ = Nothing
    Set lcol����Դ = Nothing
    Exit Sub
errHandler:
    sfsub������ "ְҵ�����󲿼�", "clsMedicalExam", "sub��ȡ�Ļ��̶��ֵ�", Err.Number, Err.Description, True
End Sub

'��ȡְҵ����ֵ�
Private Sub funcְҵ���()
    Dim lobjְҵ��� As Object
    Dim lcolְҵ��� As Object
    Dim i As Integer
    On Error GoTo errHandler
    ccmbְҵ���.Visible = True
    Set lobjְҵ��� = CreateObject("ְҵ������.clsmedicalexamtemplateset")
    lobjְҵ���.int����Դ = ccmb����Դ.ItemData(ccmb����Դ.ListIndex)
    Set lcolְҵ��� = lobjְҵ���.ְҵ���
    ccmbְҵ���.Clear
    ccmbְҵ���.AddItem ""
    For i = 1 To lcolְҵ���.RecordCount
        ccmbְҵ���.AddItem lcolְҵ���("����")
        ccmbְҵ���.ItemData(ccmbְҵ���.NewIndex) = lcolְҵ���("���")
        lcolְҵ���.MoveNext
    Next
    ccmbְҵ���.Text = ccmbְҵ���.List(0)
    Set lobjְҵ��� = Nothing
    Set lcolְҵ��� = Nothing
    Exit Sub
errHandler:
    sfsub������ "ְҵ�����󲿼�", "clsMedicalExam", "sub��ȡ�Ļ��̶��ֵ�", Err.Number, Err.Description, True
End Sub

'��ȡ��ҵ����ֵ�
Private Sub func��ȡ��ҵ����ֵ�()
    Dim lobjRec As Object
    Dim lobjDetl As Object
    Dim i As Integer
    On Error GoTo errHandler
    Set lobjRec = CreateObject("ְҵ������.clsmedicalexamtemplateset")
    Set lobjDetl = lobjRec.��ҵ���
    Ccmb��ҵ���.Clear
    Ccmb��ҵ���.AddItem ""
    For i = 1 To lobjDetl.RecordCount
        Ccmb��ҵ���.AddItem lobjDetl("����")
        Ccmb��ҵ���.ItemData(Ccmb��ҵ���.NewIndex) = lobjDetl("���")
        lobjDetl.MoveNext
    Next
    'Ccmb��ҵ���.ListIndex = 0
    
    Set lobjRec = Nothing
    Set lobjDetl = Nothing
    Exit Sub
errHandler:
    sfsub������ "ְҵ�����󲿼�", "clsMedicalExam", "sub func��ȡ��ҵ����ֵ�", Err.Number, Err.Description, False
End Sub


'��ȡ��λ������Ϣ
Private Function func��ȡ��λ��Ϣ(��λ��� As String)
    Dim lobj��λ As Object
    On Error GoTo errHandler
    Set lobj��λ = dafuncGetData("select * from ��λ����_��λ������Ϣ�� where ������='" & ��λ��� & "'")
    If Not lobj��λ.RecordCount = 0 Then
        ccmbUnit.Text = IIf(IsNull(lobj��λ("��λ����")), "", lobj��λ("��λ����"))
        mstr��λ������ = ��λ���
        ctxt������.Text = IIf(IsNull(lobj��λ("������")), "", lobj��λ("������"))
        ctxt��ϵ�绰.Text = IIf(IsNull(lobj��λ("�绰")), "", lobj��λ("�绰"))
        ccmb��������.Text = IIf(IsNull(lobj��λ("��������")), "", lobj��λ("��������"))
        Ccmb��ҵ���.Text = IIf(IsNull(lobj��λ("��ҵ���")), "", lobj��λ("��ҵ���"))
        ctxt��λ��ַ.Text = IIf(IsNull(lobj��λ("��ַ")), "", lobj��λ("��ַ"))
    End If
    Exit Function
errHandler:
    sfsub������ "ְҵ������", "frmregister", "func��ȡ��λ��Ϣ", Err.Number, Err.Description, True
End Function

'�������֤����������ʼ����PC���ն˵�����
Private Sub sub��������ʼ��()
    'CVR_InitComm
    On Error GoTo errHandler
   Dim n, ret, nLen
    Comm = False
    
    For n = 1001 To 1016 Step 1     '���μ��USB�˿�1001-1016
 
      If (InitComm(n)) Then
            Comm = True
       
            'StateLabel.Caption = "�ɹ��򿪶˿ڣ�"
            'ret = MsgBox("�ɹ��򿪶˿ڣ��뽫�������Ķ����ϡ�", vbOKOnly + vbInformation, "��ʾ")
        
            Exit For
                    
        End If
       
    Next n
    If (Comm = False) Then
     For n = 1 To 4 Step 1     '���μ�鴮��1-16
    
        If (InitComm(n)) Then
            Comm = True
       
            'StateLabel.Caption = "�ɹ��򿪶˿ڣ�"
            'ret = MsgBox("�ɹ��򿪶˿ڣ��뽫�������Ķ����ϡ�", vbOKOnly + vbInformation, "��ʾ")
    
           Exit For
                    
        End If
       
       Next n
    End If
    
   
  
    If (Comm = False) Then
    
            ret = MsgBox("�򿪶˿ڲ��ɹ��������豸���ӡ�", vbOKOnly + vbCritical, "����")
            
            Exit Sub
    
    End If
    Exit Sub
errHandler:
    sfsub������ "ְҵ�����沿��", "frmregister", "func��������ʼ��", Err.Number, Err.Description, True
End Sub

'ֻ�о������֤��֤�󣬲��ܶ�ȡ��Ϣ����Ƭ��
Private Function func���֤��֤() As Integer
    'CVR_Authenticate
    'On Error GoTo errHandler
    'Dim temp As Integer
    'List1.AddItem "�����֤��֤��"
    'List1.AddItem " ���� " & CVR_Authenticate()
   ' func���֤��֤ = CVR_Authenticate()
    Exit Function
errHandler:
    sfsub������ "ְҵ�����沿��", "frmregister", "func���֤��֤", Err.Number, Err.Description, True
End Function

'��timer2 ���м���Ƿ������֤���ڶ������ϣ�����Ϊ350ms
'��timer2ʱ���Ϊ900ms  2016-1-6 by Ĳ��
Private Sub Timer2_Timer()
    '2012-07-11 �ڵ�� ��
    '��֪Ϊ�Σ����ǵ����Ҳ���termb.dll�ļ������ǣ������޸Ĵ�����
    'On Error GoTo errHandler
    On Error Resume Next
       Dim sΣ������ As String
    Dim s������� As String
    Dim s�������� As String
    Dim strs As String
    Dim strb  As String
    Dim strx As String
    Dim rs As Object
    Dim s������Ϣ As String
    Dim ccmbi
    Dim ccmbj
    Dim ccmbk


''    �Ӳ�ѯҳ�洫һ�������������ǲ�ѯҳ����ת�����ġ�����ֻ�ж�ctxt���֤��.Text
'     If (ctxt���֤��.Text <> "") Then
'    '�������ѯ�����Ա������Ϣ�����������ͺ��������Լ�Σ�����ء���ɵĽṹ�硰ְҵ���-�ڸ��ڼ�-�۳���modify by liuwei
'
'
'    strs = "select Σ������,�������,�������� from ְҵ�����_�����Ա������Ϣ�� where ϵͳ���='" & clblsysno.Text & "'"
'    Set rs = dafuncGetData(strs)
'    sΣ������ = rs("Σ������")
'    s������� = rs("�������")
'    s�������� = rs("��������")
'
'    s������Ϣ = s�������� + "-" + s������� + "-" + sΣ������
'
'
'    For ccmbi = 0 To ccmb���������.ListCount - 1
'    If ccmb���������.List(ccmbi) = s�������� Then
'    ' ccmb���������.Text = ccmb���������.List(ccmbi)
'     ccmb���������.ListIndex = ccmb���������.ItemData(ccmbi) - 1
'
'    End If
'    Next ccmbi
'    For ccmbj = 0 To Ccmb��������.ListCount - 1
'    If Ccmb��������.List(ccmbj) = s������� Then
'    ' Ccmb��������.Text = Ccmb��������.List(ccmbj)
'     Ccmb��������.ListIndex = Ccmb��������.ItemData(ccmbj) - 1
'
'    End If
'    Next ccmbj
'
'    '����ط������⣬��Ҫ�޸� lable by lanchao 2015-03-15
'    For ccmbk = 0 To ccmbTemplate.ListCount - 1
'     If ccmbTemplate.List(ccmbk) = s������Ϣ Then
'     'ccmbTemplate.Text = ccmbTemplate.List(ccmbk)
'
'     ccmbTemplate.ListIndex = ccmbTemplate.ItemData(ccmbk) - 1
'     ccmbTemplate.ListIndex = ccmbk
'    End If
'    Next ccmbk
'
'Else
     
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
    ChDir (App.Path)                '�ı䵱ǰĬ��·��ΪӦ�ó�������·��
    ret = Authenticate()
    If (ret) Then
       ret = ReadBaseInfos(iname, isex, folk, birthday, code, addr, agency, startdate, enddate)
       '�����жϣ�ֱ�Ӷ�ȡ���֤��Ϣ  modify by lanchao 2015-9-10
'       If Trim(ctxtName.Text) <> Trim(Split(iname, "")(0)) Then
'       Dim msgs
'       msgs = MsgBox("���֤��Ϣ�������Ϣ��ƥ�䡣", vbOKOnly + vbInformation, "��ʾ")
'       Exit Sub
'       End If'
       ctxtName.Text = Trim(Split(iname, "")(0))
        Timer2.Enabled = False

        '����⵽�����֤����ȡ���ݳɹ��󣬹ر�timer2
        'Call sub��ȡ��Ϣ
        ctxt���֤��.Enabled = True
        ctxt���֤��.SetFocus      '��setfocus��lostfocusĿ���ǣ���ȡ�����֤�ź����  ctxt���֤��.lostfocus �¼��������Լ�����Ա����䣬��������
        'Call sub��ȡ֤��
        ctxt���֤��.Text = Trim(code)
        ctxt���֤��_KeyDown 13, 1
        ctxt���֤��.Enabled = False
        ctxtName.Enabled = True
        ctxtName.SetFocus
        'Call sub��ȡ����
        ctxtName.Text = Trim(Split(iname, "")(0))
        ctxtName.Enabled = False
        '2012-06-13 �ڵ�� ��
        'ʡ������Ҫ������ͼƬ�����֤ͼƬ���洢
        '��ʾ���֤ͼƬ
       
        '2012-06-13 �ڵ�� ��
        'Call sub��ȡסַ
        ctxtסַ.Text = Trim(addr)
        'Call sub��ȡ����
        ccmb����.Text = Trim(folk)
        '2012-07-12 �ڵ�� ��
        'ʵ�ֹر����֤�Ķ����������ظ���ʱ����
        'ret = CloseComm()
        '2012-07-12 �ڵ�� ��

'        '2012-07-15 �ڵ�� ��
'        'ÿ�εõ��µ����֤ʱ������������������е����ʱ���¼
'        sub�鿴��ʷ��Ϣ (Trim(ctxt���֤��.Text))    '�����Ѿ���ѯ����ʷ��Ϣ�ˣ��˴�����  2016-1-15 by Ĳ��
'        '2012-07-15 �ڵ�� ��
        

       '���ʱ��2015-2-25
'�������ѯ�����Ա������Ϣ�����������ͺ��������Լ�Σ�����ء���ɵĽṹ�硰ְҵ���-�ڸ��ڼ�-�۳���modify by liuwei
strs = "select Σ������,�������,�������� from ְҵ�����_�����Ա������Ϣ�� where ϵͳ���='" & clblsysno.Text & "'"
Set rs = dafuncGetData(strs)
sΣ������ = rs("Σ������")
s������� = rs("�������")
s�������� = rs("��������")

s������Ϣ = s�������� + "-" + s������� + "-" + sΣ������

For ccmbi = 0 To ccmb���������.ListCount - 1
  If ccmb���������.List(ccmbi) = s�������� Then
    ' ccmb���������.Text = ccmb���������.List(ccmbi)
     ccmb���������.ListIndex = ccmb���������.ItemData(ccmbi) - 1
     
  End If
Next ccmbi

For ccmbj = 0 To Ccmb��������.ListCount - 1
  If Ccmb��������.List(ccmbj) = s������� Then
    ' Ccmb��������.Text = Ccmb��������.List(ccmbj)
     Ccmb��������.ListIndex = Ccmb��������.ItemData(ccmbj) - 1
     
  End If
Next ccmbj

'����ط������⣬��Ҫ�޸� lable by lanchao 2015-03-15
For ccmbk = 0 To ccmbTemplate.ListCount - 1
  If ccmbTemplate.List(ccmbk) = s������Ϣ Then
     'ccmbTemplate.Text = ccmbTemplate.List(ccmbk)
     
     ccmbTemplate.ListIndex = ccmbTemplate.ItemData(ccmbk) - 1
     ccmbTemplate.ListIndex = ccmbk
  End If
Next ccmbk
ccmbTemplate.Text = s������Ϣ     '2015-12-10 by Ĳ��
End If
'End If
   Picture2.Picture = LoadPicture(App.Path & "\photo.bmp")
'   Timer2.Enabled = True  '2016-2-24 by Ĳ��
    Exit Sub
errHandler:
    sfsub������ "ְҵ�����沿��", "frmregister", "timer2_timer", Err.Number, Err.Description, True
End Sub


'���֤��֤�Ժ󣬲Ŷ�ȡ��Ϣ��ֻ�ж�ȡ��Ϣ�Ժ󣬲Ż��ڵ�ǰĿ¼�������֤����Ƭ�ļ�zp.mbp
'Private Sub sub��ȡ��Ϣ()
'    'CVR_Read_Content
'    Dim mode As Integer
'    On Error GoTo errHandler
'    'modeȡֵ��
'    '1: ��������wz.txt?��Ƭ����xp.wlt����Ƭzp.bmp (����)
'    '2: ��������wz.txt����Ƭ����xp.wlt
'    '4: ����wz.txt(����)����Ƭzp.bmp(����)
'    '6: �������豸ģ�����������.txt�ļ�(����)����Ƭ.bmp�ļ�(����)
'    mode = 4
'    CVR_Read_Content (mode)
'    Exit Sub
'errHandler:
'    sfsub������ "ְҵ�����沿��", "frmregister", "sub��ȡ��Ϣ", Err.Number, Err.Description, True
'End Sub
'
''��ȡ���֤��
'Private Sub sub��ȡ֤��()
'    Dim strTemp As String
'    Dim nReturnLen As Integer
'    Dim nReturn As Integer
'    strTemp = Space(255)
'    nReturn = GetPeopleIDCode(strTemp, nReturnLen)
'    ctxt���֤��.Text = Trim(strTemp)
'End Sub
'
'Private Sub sub��ȡ����()
'    Dim strTemp As String
'    Dim nReturnLen As Integer
'    Dim nReturn As Integer
'    strTemp = Space(255)
'    nReturn = GetPeopleName(strTemp, nReturnLen)
'    ctxtName.Text = Trim(strTemp)
'End Sub
'
'Private Sub sub��ȡסַ()
'    Dim strTemp As String
'    Dim nReturnLen As Integer
'    Dim nReturn As Integer
'    strTemp = Space(255)
'    nReturn = GetPeopleAddress(strTemp, nReturnLen)
'    ctxtסַ.Text = Trim(strTemp)
'End Sub
'
'Private Sub sub��ȡ����()
'    Dim strTemp As String
'    Dim nReturnLen As Integer
'    Dim nReturn As Integer
'    strTemp = Space(10)
'    nReturn = GetPeopleNation(strTemp, nReturnLen)
'    ccmb����.Text = Trim(strTemp)
'End Sub

'���ܣ�����ְҵ���Ǽ���ѡ��������Ŀ
'���ߣ�����
'ʱ�䣺2012-06-04
'˵��������Ҫ�鿴���ݿ������Ƿ�����ͬ�������Ŀ��Ȼ���ٽ������ӻ����޸�

Public Sub save�Ż��������Ŀ(ByRef para�����Ŀ As Collection, ByVal paraϵͳ��� As String)
    Dim lstrSql As String
    Dim MedicProjt As String
    Dim rs As Object
    Dim i As Integer
    Dim col�����Ŀ As Collection
    On Error GoTo errHandler
    
    Set rs = dafuncGetData("select ���� from ϵͳ����_�ֵ�_�ֵ����ݱ� where ID = (select ID from ϵͳ����_�ֵ�_�ֵ���б� where ����='ְҵ���������ֵ�') and ���� like '%��'")
    
    For i = 1 To rs.RecordCount
        
        lstrSql = "delete ְҵ�����_�����Ϣ_" & rs("����") & " where ϵͳ���='" & paraϵͳ��� & "'"
        dafuncGetData lstrSql
        rs.MoveNext
    Next i
    
    Set col�����Ŀ = para�����Ŀ
    
    For i = 1 To col�����Ŀ.Count
        MedicProjt = Left(Trim(col�����Ŀ(i).Item(1)), 2)
        
        lstrSql = "select ���� from ϵͳ����_�ֵ�_�ֵ����ݱ� where ID = (select ID from ϵͳ����_�ֵ�_�ֵ���б� where ����='ְҵ���������ֵ�') and ���= '" & MedicProjt & "'"
        Set rs = dafuncGetData(lstrSql)
    
        lstrSql = "insert into ְҵ�����_�����Ϣ_" & rs("����") & "(ϵͳ���,�����Ŀ) values('" & paraϵͳ��� & "','" & col�����Ŀ(i).Item(1) & "')"
        dafuncGetData lstrSql
    Next i
    
    Exit Sub
errHandler:
   sfsub������ "ְҵ��ʷ¼��", "clscareerhstregt", "public sub save�����Ŀ", Err.Number, Err.Description, False
End Sub

'2012-06-25 �ڵ��
'��ӳ�ʼ���������״̬������
'�����ж�ÿ�������Ա���������(�����)���ҵ����״̬��
'0������Ҫ����Ŀ��ң�1������Ҫ����Ŀ��ң�2����ÿ����Ѿ������ꣻ
'3����ÿ������������۲��������޸ġ�(���У�2��3״̬�����������ս���)
'״̬��һ������Ϊ13���ַ���(6-25ʱ��13����д����Ŀ��ң��ַ�������Ϊ18)
Sub subInit�������״̬(paraCol As Collection, paraSysNo As String)
    Dim i As Integer
    Dim paraDeptNo As Integer
    Dim paraState, strSQL As String
    
    
    For i = 1 To 17: paraState = paraState & "0": Next
    'paraState = paraState & "1"
    
    For i = 1 To paraCol.Count
        paraDeptNo = CInt(Left(paraCol.Item(i).Item(1), 2))
        paraState = Left(paraState, paraDeptNo - 1) & "1" & Right(paraState, Len(paraState) - (paraDeptNo))
    Next
    
    strSQL = "update ְҵ�����_��������Ϣ�� set �������״̬='" & paraState & "' where ϵͳ���='" & paraSysNo & "'"
    dafuncGetData strSQL
End Sub

'2012-07-16 �ڵ��
'�鿴�̶����֤�ŵ������Ա��ʱ��Ϣ��������ѯ�������cgrdHistory��
'when ���״̬=0 then 'δУ��'
'when ���״̬=1 then 'δ���嵥'
'when ���״̬=2 then 'δ¼���ܼ��߸�����Ϣ'
'when ���״̬=3 then '�����'
'when ���״̬=4 then 'δ�½���'
'when ���״̬=5 then '���½���'
'when ���״̬=6 then '�Ѹ���'
'when ���״̬=7 then '�ѷ�����'
'when ���״̬=8 then '������'
Sub sub�鿴��ʷ��Ϣ(ByVal paraIDCard As String)
    Dim strSQL As String
    Dim lobjRec As Object
    Dim initState(0 To 8) As String
    Dim i, j As Integer
    '2012-12-18 ������
    'bug No:0000087,0000084
'    strSQL = "select ϵͳ���,�������� ����ʱ��,������ ����,���״̬ from ְҵ�����_���������ݿ� where ������ݺ���='" & paraIDCard & "' and ��������>='" & Format(DateAdd("yyyy", -5, Now), "yyyy-mm-dd") & "'"
    strSQL = "select ϵͳ���,�������� ����ʱ��,������ ����,���״̬ from ְҵ�����_���������ݿ� where ������ݺ���='" & paraIDCard & "' and ��������<'" & Format(Now, "yyyy-mm-dd") & "'"
    ''2012-12-18 ������
    Set lobjRec = dafuncGetData(strSQL)
    If lobjRec.RecordCount > 0 Then
        '���ƽ���ؼ��Ƿ���ʾ����ʾ��ʽ
        clblHistory.Visible = True
        With cgrdHistory
            .Visible = True
            Set .DataSource = lobjRec
            .Col = 0
            .Sort = flexSortGenericDescending
            .DataMode = flexDMFree
            .AllowSelection = False
            .AllowBigSelection = False
            .SelectionMode = flexSelectionListBox
        End With
        
        ReDim lcolIndex(0 To lobjRec.Fields.Count - 1) As String
        For i = 0 To lobjRec.Fields.Count - 1
            lcolIndex(i) = lobjRec.Fields.Item(i).name
        Next
        '��ʾʱ���������״̬��ʾ���֣����滻Ϊ���֡�ԭʼ���մ洢���̡�ְҵ�����_����������ѯ��
        '��ʼ��״̬����
        initState(0) = "δУ��"
        initState(1) = "δ���嵥"
        initState(2) = "δ¼���ܼ��߸�����Ϣ"
        initState(3) = "�����"
        initState(4) = "δ�½���"
'        initState(5) = "���½���"
        initState(5) = "������"    '2015-12-15 by Ĳ��
        initState(6) = "�Ѹ���"
        initState(7) = "�ѷ�����"
        initState(8) = "������"
        
        lobjRec.MoveFirst
        For i = 1 To lobjRec.RecordCount
            For j = 0 To lobjRec.Fields.Count - 1
                If lcolIndex(j) = "���״̬" Then cgrdHistory.TextMatrix(i, j) = initState(CInt(cgrdHistory.TextMatrix(i, j)))
            Next
            lobjRec.MoveNext
        Next
        
        cgrdHistory.AutoSize 0, cgrdHistory.cols - 1, 0, 0
    End If
End Sub

'2012-08-18 �ڵ��
'����������
Private Sub sub���鱣��()
        Dim i As Integer
        Dim lstr��ˮ�� As String
        Dim lstrϵͳ��� As String
        Dim lcolԭ�����Ŀ As Collection
        Dim lobjrec���� As Object
        Dim lobj������ As Object
        Dim lobjRec As Object
        Dim lstrError As String
        
        '2012-06-13 �ڵ�� ��
        '�洢���֤��Ƭ������ϵͳ����˻ر��������֤�����Ϣ
        Dim lobjRec���֤��Ƭ As Object
        Dim lobjRecϵͳ����˻� As Object
        Dim paraSysNo As String
        Dim lstrSex As String
        Dim lstrBirth As String
        '2012-06-13 �ڵ�� ��
        
        '2012-08-18 �ڵ�� ��
        '���������¸��ġ�
        Dim mobj������챣����� As Object
        Set mobj������챣����� = CreateObject("ְҵ������.clsMedicalExam")
        '2012-08-18 �ڵ�� ��
        
        '2012-07-11 �ڵ�� ��
        '���У��ͨ����ֻ�ܴ洢�ֳ���Ƭ������ʱ�����ʱ�־=2�����������������ݡ���
        MousePointer = 11
        If mintState = 1 And ���ʱ�־ = 1 Then
            If mblnTakePhoto Then
                Dim lobjPhoto As StdPicture
                '����Ƭ��Ϊ�գ��򱣴浽��ӦĿ¼�����õ��� ͨ�ö���.clsͼƬ����.cls
                If Not cctlCatchPhoto.Photo Is Nothing Then
                    Set lobjPhoto = cctlCatchPhoto.Photo
                    pmsub����ͼƬ lobjPhoto, Trim(clblsysno.Text), "ְҵ�����"
                End If
            End If
            MousePointer = 0
            Set lobjrec���� = Nothing
            Set lobj������ = Nothing
            Exit Sub
        End If
        '2012-07-11 �ڵ�� ��
        
        '2012-07-11 �ڵ�� ��
        'δ�������֤�Ż����֤�Ŵ��󣬲������浱ǰ��Ϣ
        If Len(ctxt���֤��.Text) = 0 Then
            MousePointer = 0
            MsgBox ("δ�������֤�ţ��������浱ǰ���ݣ�")
            Exit Sub
        End If
        sub���ݹ�����ݺ����ȡ���պ��Ա� ctxt���֤��.Text, lstrBirth, lstrSex
        If ccmbSex.Text = "" Or lstrSex <> ccmbSex.Text Or Format(lstrBirth, "yyyy-mm-dd") <> Format(cdtp����.Value, "yyyy-mm-dd") Then
            MsgBox ("���֤�뵱ǰ������Ϣ�������������浱ǰ���ݣ�")
            Exit Sub
        End If
        '2012-07-11 �ڵ�� ��
        
        '���� ϵͳ��� �ı��� ����
        ctxt����.SetFocus
        
        '2012-06-13 �ڵ�� ��
        'ʡ������Ҫ��������Ƭ�����֤��Ƭ�ֿ��洢���ֿ���ʾ
        '���ﵥ���洢���֤��Ƭ��������Ƭ��ԭ�������洢��
        Set lobjRec���֤��Ƭ = CreateObject("ְҵ������.clsPersonExamed")
        lobjRec���֤��Ƭ.func�������֤��Ƭ Picture2.Image, pstr����ϵͳ��� & "IDcard", "ְҵ�����"
        Set lobjRec���֤��Ƭ = Nothing
        '2012-06-13 �ڵ�� ��
        
        MousePointer = 11
        
        Set lobj������ = CreateObject("ְҵ������.clsmedicalexamsheet")
        lobj������.������ = ccmbTemplate.Text
        
        pstrϵͳ��� = clblsysno.Text
        '�����Թܱ�Ų�����
        With mobj������챣�����
            '2012-06-14 �ڵ�� ��
            'ϵͳ��ű������������¸�ֵ�������һ�λ�ʹ��form_loadʱ��ϵͳ���
            .ϵͳ��� = pstr����ϵͳ���
            '2012-06-14 �ڵ�� ��
            
            If .����.������ <> ccmbTemplate.Text Then
                .����.������ = ccmbTemplate.Text
            End If
            '�޸ģ�2004-1-9���Թܱ�ſ������룩
            If pobjҵ�����.ҵ������("�Թܱ���Զ�����") = "��" Then
                If .����.�Թܱ����ĸ <> clblLetter.Caption Then
                    .����.�Թܱ����ĸ = clblLetter.Caption
                End If
            Else
                .����.�Թܱ����ĸ = clblLetter.Caption
                .�Թܱ�� = ctxtTubeNo.Text
            End If
            .�����Ա.ϵͳ��� = pstr����ϵͳ���
            .�����Ա.���� = ctxtName
            .�����Ա.�Ա� = ccmbSex.Text
            .�����Ա.��λ���� = ccmbUnit.Text
            .�����Ա.Σ������ = ccmbΣ������.Text
            .�����Ա.����Դ = ccmb����Դ.Text
            .�����Ա.ְҵ���� = ccmbְҵ���.Text
            .�����Ա.�ֹ��� = ccmb�ֹ���.Text
            .�����Ա.ְ���ְ�� = ccmbְ��.Text
            .�����Ա.ְҵΣ������ = ctxtΣ������.Text
            .�����Ա.������� = Trim(ctxt�������.Text)
            .�����Ա.���� = Trim(ctxt����.Text)
            .�����Ա.�ʱ� = Trim(ctxt�ʱ�.Text)
            .�����Ա.סַ = Trim(ctxtסַ.Text)
            .�����Ա.��� = Ccmb���.Text
            .�����Ա.�绰���� = Trim(ctxt�绰.Text)
            .�����Ա.���� = Trim(ctxt����.Text)
            .�����Ա.������ = Trim(ctxt������.Text)
            .�����Ա.������ = ctxt������.Text
            .�����Ա.��ϵ�绰 = ctxt��ϵ�绰.Text
            .�����Ա.�������� = ccmb��������.Text
            .�����Ա.��ҵ��� = Ccmb��ҵ���.Text
            .�����Ա.��λ��ַ = ctxt��λ��ַ.Text
            If mblnTakePhoto Then
                .�����Ա.��Ƭ = cctlCatchPhoto.Photo
'                .�����Ա.��Ƭѹ�� = cctlCatchPhoto.Photo
            ElseIf Not Picture1.Picture Is Nothing Then
                .�����Ա.��Ƭ = Picture1.Picture
            End If
            If Val(ctxtAge.Text) > 0 Then
'                If Val(ctxtAge.Text) > 200 Then
'                    Err.Raise 6666, , "���䳬��ϵͳ������������200��"
'                End If
                .�����Ա.�������� = DateAdd("yyyy", -Val(ctxtAge.Text), Date)
            Else
                '��������ַ������������䡣
                mobj����.sub���Ǽ���ֵ "�������", ctxtAge.Text
                mstrĬ������ = ctxtAge.Text
            End If
            .�����Ա.���� = ctxtAge.Text
            
            On Error Resume Next
            .�����Ա.������ݺ��� = ctxt���֤��.Text
            .�����Ա.�Ļ��̶� = ccmb�Ļ��̶�.Text
            .�����Ա.���� = ccmb����.Text
            If ccmbUnit.Text = "" Then
                .�����Ա.��λ������ = ""
            Else
                If .�����Ա.��λ������ <> mstr��λ������ Then
                    '����λ������¸�ֵ���������»�ȡ���������ࡢ��ҵ���Ƭ����
                    .�����Ա.��λ������ = mstr��λ������
                End If
            End If
            
            .������� = cdtpDate.Value ' ,Format(cdtpDate.Value, "yyyy-mm-dd hh:mm:ss")
            
            '�޸ģ�2004-1-9��������쵥�ţ�
            .��������� = ccmb���������.Text
            .�������� = Ccmb��������.Text
            
            'On Error GoTo errHandler
            On Error Resume Next
            If mcol�����Ŀ.Count = 0 Then
'                mobj������챣�����.����.mbln�Ƿ��Ѵ��� = True
'                mobj������챣�����.����.mbln�Ƿ��ѻ�ȡ�����Ŀ = False
'                mobj������챣�����.����.mbln�Ƿ��ѻ�ȡ������Ŀ = False
                Set mcol�����Ŀ = mobj���.����.�����Ŀ��("")
                frmSelectItem.pstr�������� = ccmbTemplate.Text
                Set frmSelectItem.pcol������Ŀ = mcol�����Ŀ
                frmSelectItem.Hide
                frmSelectItem.ccmdOk_Click
                Set mcol�����Ŀ = frmSelectItem.pcol������Ŀ
            End If
            Set .col�����Ŀ = mcol�����Ŀ
        
        End With
        
        '���ܣ����������Ŀ
        'ʱ�䣺2012-06-04
        '���ߣ�����
        save�Ż��������Ŀ mcol�����Ŀ, pstr����ϵͳ���
        'ʱ�䣺2012-06-04
        
        If mcol�շ���Ŀ.Count > 0 Then
            pobjҵ�����.Sub���Ǽ� mobj������챣�����, , , mcol�շ���Ŀ, Val(ctxt����)
        Else
            pobjҵ�����.Sub���Ǽ� mobj������챣�����, , , , Val(ctxt����)
        End If
        
        Set lobjRec = CreateObject("ְҵ������.clsMoney")
        lobjRec.mstrϵͳ��� = pstr����ϵͳ���
        lobjRec.mstr�����Ա���� = ctxtName.Text
        Set lobjRec.col�����Ŀ = mcol�����Ŀ
        Dim lstr�շ����� As String
        lstrError = lobjRec.func�շ�(lstr�շ�����)
        mobj������챣�����.�շ����� = lstr�շ�����
        If lstrError <> "" And lstrError <> "Cancel" Then
            MsgBox lstrError, vbOKOnly + vbExclamation, "ϵͳ��ʾ"
        End If
    
        cstbMain.Panels(1) = "�ϴα�������ϵͳ��ţ�" & mobj������챣�����.ϵͳ��� '& " ���Թܱ�ţ�" & mobj������챣�����.�Թܱ��
        If mobj������챣�����.�շ����� <> "" Then
            cstbMain.Panels(1) = cstbMain.Panels(1) & "���շ����ţ�" & mobj������챣�����.�շ�����
        End If
        
        '2012-06-25 �ڵ�� ��
        '��ʼ����������Ϣ���С��������״̬���ֶ�
        subInit�������״̬ mcol�����Ŀ, pstr����ϵͳ���
        '2012-06-25 �ڵ�� ��
        '�ǼǺ��޸ĸ���״̬�����ǣ�2012-10-30
        dafuncGetData "update ְҵ�����_��������Ϣ�� set ����״̬ = '1' where ����ϵͳ��� = '" & pstr����ϵͳ��� & "'"
        '2012-08-18 �ڵ�� ��
        '����ǼǺ���Ϊ��У�˺�δ��ӡ�嵥״̬����Ҫд�뵱ǰ���״̬�����ࡣ
        '��������Ǽǲ�ͬ��
        If mintState = 1 Then
            pobjҵ�����.funcд�뵥�˵�ǰ���״̬ pstr����ϵͳ���, mintState
            pobjҵ�����.funcд�뵥�˵�ǰ���״̬ pstrϵͳ���, 5   '5�������״̬"���½���"
            mobjGUI_BeforeOperate "У��ͨ��", False
        End If
        mintState = 2
        '2012-06-15 �ڵ�� ��
        
        '�ָ����ࡣ
        If mblnTakePhoto Then
            If cctlCatchPhoto.Status = "�ָ�" Then
                cctlCatchPhoto.subת��״̬
            End If
        End If
        
        If cchkClear = 1 Then
            subClear
            ���ʱ�־ = 0
            clblsysno.Text = mobj���.Func����ְҵ�����ϵͳ��� & (ccmb���������.ListIndex + 1)
        End If
        
        Set mcol�����Ŀ = New Collection
       
        mobj���.����.������ = ccmbTemplate.Text
               
        '�Թ���ĸ������ѡ��
        cvscLetter.Enabled = False
        ctxt����.SetFocus
        frmRegisterManage.sub��ѯ����ʾ
        Timer2.Enabled = True

        MousePointer = 0
        Set lobjrec���� = Nothing
        Set lobj������ = Nothing
        Exit Sub
errHandler:
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����沿��", "FrmRegisterAnnual", "mobjGUI_BeforeOperate", 6666, lstrError, False
    
    MousePointer = 0
    cstbMain.Panels(1) = lstrError
    Exit Sub
    Resume
End Sub

'
