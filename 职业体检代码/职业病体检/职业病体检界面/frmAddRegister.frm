VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{2327075A-FED2-4AEF-82B8-DD0C2B1AC8E1}#2.0#0"; "dyCatchPhoto.ocx"
Object = "{F48D4DEC-1198-11D5-91BE-0050BA06B70C}#5.9#0"; "¼��ؼ�.ocx"
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
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   6600
      Top             =   360
   End
   Begin VB.CheckBox cchkClear 
      Caption         =   "��������"
      Height          =   345
      Left            =   8520
      TabIndex        =   82
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.CheckBox Check���֤ 
      Caption         =   "ˢ�������֤"
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
      TabCaption(0)   =   "������Ϣ¼��            "
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
      Tab(0).Control(19)=   "cdtp����"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "cgrdHistory"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "ccmbSex"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Picture2"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Ccmb��������"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "ccmb���������"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "clblsysno"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "ctxt���֤��"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "ctxtAge"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "ctxtName"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "ccmbTemplate"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "ccmb�������"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "ccmb������"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).ControlCount=   32
      TabCaption(1)   =   "������Ϣ¼��           "
      TabPicture(1)   =   "frmAddRegister.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(2)=   "Frame3"
      Tab(1).ControlCount=   3
      Begin VB.ComboBox ccmb������ 
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
      Begin VB.ComboBox ccmb������� 
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
      Begin VB.TextBox ctxt���֤�� 
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
         Caption         =   "  ���˸�����Ϣ      "
         ForeColor       =   &H000080FF&
         Height          =   2295
         Left            =   -74520
         TabIndex        =   37
         Top             =   480
         Width           =   9735
         Begin VB.TextBox ctxt���� 
            Height          =   300
            Left            =   480
            TabIndex        =   46
            Top             =   600
            Width           =   1815
         End
         Begin VB.TextBox ctxt�ʱ� 
            Height          =   300
            Left            =   2640
            TabIndex        =   45
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox ctxtסַ 
            Height          =   300
            Left            =   4560
            TabIndex        =   44
            Top             =   600
            Width           =   4215
         End
         Begin VB.ComboBox ccmb�Ļ��̶� 
            Height          =   300
            ItemData        =   "frmAddRegister.frx":0055
            Left            =   480
            List            =   "frmAddRegister.frx":0074
            TabIndex        =   43
            Text            =   "ccmb�Ļ��̶�"
            Top             =   1320
            Width           =   1815
         End
         Begin VB.TextBox ctxt���� 
            Height          =   270
            Left            =   7920
            TabIndex        =   42
            Top             =   1320
            Width           =   975
         End
         Begin VB.TextBox ctxt�绰 
            Height          =   300
            Left            =   5400
            TabIndex        =   41
            Top             =   1320
            Width           =   1815
         End
         Begin VB.ComboBox Ccmb��� 
            Height          =   300
            ItemData        =   "frmAddRegister.frx":00C6
            Left            =   2640
            List            =   "frmAddRegister.frx":00D3
            TabIndex        =   40
            Text            =   "Ccmb���"
            Top             =   1320
            Width           =   855
         End
         Begin VB.ComboBox ccmb���� 
            Height          =   300
            ItemData        =   "frmAddRegister.frx":00E9
            Left            =   3840
            List            =   "frmAddRegister.frx":00F0
            TabIndex        =   39
            Text            =   "ccmb����"
            Top             =   1320
            Width           =   1215
         End
         Begin VB.TextBox ctxt������ 
            Height          =   300
            Left            =   1200
            TabIndex        =   38
            Top             =   1750
            Width           =   4455
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "���᣺"
            Height          =   180
            Left            =   480
            TabIndex        =   55
            Top             =   360
            Width           =   540
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "������ţ�"
            Height          =   180
            Left            =   2640
            TabIndex        =   54
            Top             =   360
            Width           =   900
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "סַ��"
            Height          =   180
            Left            =   4560
            TabIndex        =   53
            Top             =   360
            Width           =   540
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "���䣺"
            Height          =   180
            Left            =   7920
            TabIndex        =   52
            Top             =   1080
            Width           =   540
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "�绰���룺"
            Height          =   180
            Left            =   5400
            TabIndex        =   51
            Top             =   1080
            Width           =   900
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "�Ļ��̶ȣ�"
            Height          =   180
            Left            =   480
            TabIndex        =   50
            Top             =   1080
            Width           =   900
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "���"
            Height          =   180
            Left            =   2640
            TabIndex        =   49
            Top             =   1080
            Width           =   540
         End
         Begin VB.Label label80 
            AutoSize        =   -1  'True
            Caption         =   "���壺"
            Height          =   180
            Left            =   3840
            TabIndex        =   48
            Top             =   1080
            Width           =   540
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "�����أ�"
            Height          =   180
            Left            =   480
            TabIndex        =   47
            Top             =   1800
            Width           =   720
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "  Σ��������Ϣ¼��    "
         ForeColor       =   &H000080FF&
         Height          =   1815
         Left            =   -74520
         TabIndex        =   22
         Top             =   2880
         Width           =   9735
         Begin VB.TextBox ctxtΣ������ 
            Height          =   270
            Left            =   4560
            TabIndex        =   29
            Top             =   1200
            Width           =   1815
         End
         Begin VB.TextBox ctxt������� 
            Height          =   270
            Left            =   6720
            TabIndex        =   28
            Top             =   1200
            Width           =   1815
         End
         Begin VB.ComboBox ccmbְ�� 
            Height          =   300
            Left            =   2640
            TabIndex        =   27
            Top             =   1200
            Width           =   1575
         End
         Begin VB.ComboBox ccmb�ֹ��� 
            Height          =   300
            Left            =   480
            TabIndex        =   26
            Top             =   1200
            Width           =   1815
         End
         Begin VB.ComboBox ccmb����Դ 
            Height          =   300
            ItemData        =   "frmAddRegister.frx":00F8
            Left            =   480
            List            =   "frmAddRegister.frx":016A
            TabIndex        =   25
            Top             =   480
            Width           =   1815
         End
         Begin VB.ComboBox ccmbְҵ��� 
            Height          =   300
            ItemData        =   "frmAddRegister.frx":02F4
            Left            =   2640
            List            =   "frmAddRegister.frx":02F6
            TabIndex        =   24
            Top             =   480
            Width           =   1575
         End
         Begin VB.ComboBox ccmbΣ������ 
            Height          =   300
            Left            =   4560
            TabIndex        =   23
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "ְҵΣ�����䣺"
            Height          =   180
            Left            =   4560
            TabIndex        =   36
            Top             =   960
            Width           =   1260
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "���������"
            Height          =   180
            Left            =   6720
            TabIndex        =   35
            Top             =   960
            Width           =   900
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "ְҵ/ְ�ƣ�"
            Height          =   180
            Left            =   2640
            TabIndex        =   34
            Top             =   960
            Width           =   990
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "�ֹ��֣�"
            Height          =   180
            Left            =   480
            TabIndex        =   33
            Top             =   960
            Width           =   720
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "����Դ��"
            Height          =   180
            Left            =   480
            TabIndex        =   32
            Top             =   240
            Width           =   720
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "ְҵ���"
            Height          =   180
            Left            =   2640
            TabIndex        =   31
            Top             =   240
            Width           =   900
         End
         Begin VB.Label Label17 
            Caption         =   "Σ�����أ�"
            Height          =   255
            Left            =   4560
            TabIndex        =   30
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "  �ֵ�λ��Ϣ¼��   "
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
         Begin VB.CommandButton ccmd��λ��λ 
            Caption         =   "��λ(&T)"
            Height          =   375
            Left            =   4080
            Style           =   1  'Graphical
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   360
            Width           =   945
         End
         Begin VB.CheckBox cchk¼�뵥λ���� 
            Caption         =   "¼�뵥λ����"
            Height          =   255
            Left            =   6600
            TabIndex        =   13
            Top             =   240
            Value           =   1  'Checked
            Width           =   1575
         End
         Begin VB.TextBox ctxt������ 
            Height          =   300
            Left            =   480
            TabIndex        =   12
            Top             =   1200
            Width           =   1695
         End
         Begin VB.TextBox ctxt��ϵ�绰 
            Height          =   300
            Left            =   2520
            TabIndex        =   11
            Top             =   1200
            Width           =   1575
         End
         Begin VB.TextBox ctxt��λ��ַ 
            Height          =   300
            Left            =   480
            TabIndex        =   10
            Top             =   1920
            Width           =   5775
         End
         Begin VB.ComboBox ccmb�������� 
            Height          =   300
            Left            =   4560
            TabIndex        =   9
            Text            =   "ccmb��������"
            Top             =   1200
            Width           =   1695
         End
         Begin VB.ComboBox Ccmb��ҵ��� 
            Height          =   300
            Left            =   6720
            TabIndex        =   8
            Text            =   "Ccmb��ҵ���"
            Top             =   1200
            Width           =   1815
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��λ���ƣ�"
            Height          =   180
            Index           =   5
            Left            =   480
            TabIndex        =   21
            Top             =   240
            Width           =   900
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "�����ˣ�"
            Height          =   180
            Left            =   480
            TabIndex        =   20
            Top             =   960
            Width           =   720
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "��ϵ�绰��"
            Height          =   180
            Left            =   2520
            TabIndex        =   19
            Top             =   960
            Width           =   900
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "��λ��ַ��"
            Height          =   180
            Left            =   480
            TabIndex        =   18
            Top             =   1680
            Width           =   900
         End
         Begin VB.Label Label28 
            Caption         =   "�������ʣ�"
            Height          =   255
            Left            =   4560
            TabIndex        =   17
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "��ҵ���"
            Height          =   180
            Left            =   6720
            TabIndex        =   16
            Top             =   960
            Width           =   900
         End
      End
      Begin VB.ComboBox ccmb��������� 
         Height          =   300
         ItemData        =   "frmAddRegister.frx":02F8
         Left            =   2880
         List            =   "frmAddRegister.frx":030E
         TabIndex        =   5
         Top             =   1320
         Width           =   1695
      End
      Begin VB.ComboBox Ccmb�������� 
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
      Begin MSComCtl2.DTPicker cdtp���� 
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
         TabIndex        =   64
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
         Caption         =   "������"
         Height          =   180
         Index           =   3
         Left            =   4800
         TabIndex        =   80
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label Label4 
         Caption         =   "�����Ա���ͣ�"
         Height          =   255
         Left            =   2880
         TabIndex        =   79
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   360
         TabIndex        =   78
         Top             =   1800
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������ڣ�"
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
         Caption         =   "������"
         Height          =   180
         Left            =   2640
         TabIndex        =   76
         Top             =   2760
         Width           =   540
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         Height          =   180
         Left            =   360
         TabIndex        =   75
         Top             =   3480
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���䣺"
         Height          =   180
         Index           =   6
         Left            =   1440
         TabIndex        =   74
         Top             =   3480
         Width           =   540
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "���֤�ţ�"
         Height          =   180
         Left            =   360
         TabIndex        =   73
         Top             =   2760
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ϵͳ��ţ�"
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
         Caption         =   "ע��ˢ����ǰ��ȷ���ı���������Ϊ��"
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
         Caption         =   "�ǿ���¼��ʱ��ɫΪ��¼�����¼��ʱֻ��ˢ�������֤"
         Height          =   180
         Left            =   360
         TabIndex        =   70
         Top             =   480
         Visible         =   0   'False
         Width           =   4680
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "�������ڣ�"
         Height          =   180
         Left            =   2400
         TabIndex        =   69
         Top             =   3480
         Width           =   900
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "ע����ˢ���룬��ˢ���֤"
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
         Caption         =   "�뽫�������֤���ڶ������ϣ�"
         Height          =   180
         Left            =   360
         TabIndex        =   67
         Top             =   2520
         Width           =   2520
      End
      Begin VB.Label clblHintCheck 
         Caption         =   "ע�⣺У��֮��ֻ�������࣬�������ݼ�ʹ�޸ģ�Ҳ���ᱣ�档"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3480
         TabIndex        =   66
         Top             =   720
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.Label clblHistory 
         Caption         =   "˫���У�������������Ϣ�͸�����Ϣ��"
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

Private Sub Ccmb��������_Click()
    Dim lobj As Object
    Dim lsql As String
    Dim a As String  '�������
    Dim b As String  '������
    Dim c As String  '����
    Dim i As Integer
    a = ccmb���������.Text
    b = Ccmb��������.Text
    ccmbTemplate.Clear
    lsql = "select �������� from ְҵ�����_����ģ�������Ϣ�� where �����Ա����='" & a & "' and ������='" & b & "'"
    Set lobj = dafuncGetData(lsql)
    If lobj.RecordCount > 0 Then
        lobj.MoveFirst
        For i = 0 To lobj.RecordCount - 1
            c = lobj("��������")
            ccmbTemplate.AddItem c, i
        lobj.MoveNext
        Next
    End If
End Sub

Private Sub ccmb���������_Click()
    Dim lobj As Object
    Dim lsql As String
    Dim a As String  '�������
    Dim b As String  '������
    Dim c As String  '����
    Dim i As Integer
    a = ccmb���������.Text
    b = Ccmb��������.Text
    ccmbTemplate.Clear
    lsql = "select �������� from ְҵ�����_����ģ�������Ϣ�� where �����Ա����='" & a & "' and ������='" & b & "'"
    Set lobj = dafuncGetData(lsql)
    If lobj.RecordCount > 0 Then
        lobj.MoveFirst
        For i = 0 To lobj.RecordCount - 1
            c = lobj("��������")
            ccmbTemplate.AddItem c, i
        lobj.MoveNext
        Next
    End If
End Sub
Sub subInit�������״̬(paraCol As Collection, paraSysNo As String)
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
    strSQL = "update ְҵ�����_��������Ϣ�� set �������״̬='" & paraState & "' where ϵͳ���='" & paraSysNo & "'"
    dafuncGetData strSQL
End Sub
Private Sub save�Ż��������Ŀ(ByRef para�����Ŀ As Collection, ByVal paraϵͳ��� As String)
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
        MedicProjt = Left(Trim(col�����Ŀ(i)("����")), 2)
        
        lstrSql = "select ���� from ϵͳ����_�ֵ�_�ֵ����ݱ� where ID = (select ID from ϵͳ����_�ֵ�_�ֵ���б� where ����='ְҵ���������ֵ�') and ���= '" & MedicProjt & "'"
        Set rs = dafuncGetData(lstrSql)
        
        lstrSql = "insert into ְҵ�����_�����Ϣ_" & rs("����") & "(ϵͳ���,�����Ŀ) values(" _
            & "'" & paraϵͳ��� & "','" & col�����Ŀ(i)("����") & "')"
        dafuncGetData lstrSql
    Next i
    
    
    Exit Sub
errHandler:
   sfsub������ "ְҵ������", "frmImportExcel", "save�Ż��������Ŀ", Err.Number, Err.Description, False
End Sub

Private Sub ccmd��λ��λ_Click()
    FrmCompany.Command1.Visible = True
    FrmCompany.Show 1
'    ccmbUnit.Text = FrmCompany.cgrdMain.TextMatrix(FrmCompany.cgrdMain.Row, 1)
End Sub

Private Sub clblsysno_Click()
    If clblsysno.Text = "" Or (clblsysno.Text <> "" And Len(clblsysno.Text) = 15) Then
'        clblsysno.Text = mobj���.Func����ְҵ�����ϵͳ��� & (ccmb���������.ListIndex + 1)
        clblsysno.Text = mobj���.Func����ְҵ�����ϵͳ��� & "1"    'ϵͳ���β��ȫ��Ϊ"1"
    End If
End Sub
'Private Sub Form_Activate()
'    On Error Resume Next
'    If mblnTakePhoto Then
'        '���³�ʼ������ؼ���
'        cctlCatchPhoto.funcInitVideo
'    End If
'    'ctxtName.SetFocus
'End Sub
Private Sub Form_Load()
    Set mobj��� = CreateObject("ְҵ������.clsMedicalExam")
    Set mobjGUI = New cls����ͨ�ö���
    
    '���ù�����������Ҫ�ĸ��ְ�ť��
    Dim lcol��������ť As New Collection           '�������ϵİ�ť��ʼ�����ϡ�
    With lcol��������ť
        .Add "���(&Cl)110"
'        .Add "|"
'        .Add "�����Ŀ(&T)102"
'        .Add "������Ƭ(&E)103"
        .Add "����(&R)101"
        .Add "|"
        .Add "����"
        .Add "|"
        .Add "�˳�"
    End With
    With mobjGUI
        Set .Form = Me
        Set .c������ = ctbMain
        '���ý���ͨ�ö����ṩ�ķ������Խ���ؼ����г�ʼ����
        .subInitialize lcol��������ť, ""
    End With
'    ctbMain.Buttons(3).Visible = False
'    ctbMain.Buttons(4).Visible = False
    cdtpDate.Value = Now
    subClear
    subAddList
    sub��������ʼ��
    Timer2.Enabled = True
End Sub
Private Sub subClear()
   clblsysno.Text = ""
   ccmb���������.Text = ""
   Ccmb��������.Text = ""
   ccmbTemplate.Text = ""
   ctxt���֤��.Text = ""
   ctxtName.Text = ""
   ccmbSex.Text = ""
   ctxtAge.Text = ""
   ctxt����.Text = ""
   ctxt�ʱ�.Text = ""
   ctxtסַ.Text = ""
   ccmb�Ļ��̶�.Text = ""
   Ccmb���.Text = ""
   ccmb����.Text = ""
   ctxt�绰.Text = ""
   ctxt����.Text = ""
   ctxt������.Text = ""
   ccmb����Դ.Text = ""
   ccmbְҵ���.Text = ""
   ccmbΣ������.Text = ""
   ccmb�ֹ���.Text = ""
   ccmbְ��.Text = ""
   ctxtΣ������.Text = ""
   ctxt�������.Text = ""
   ccmbUnit.Text = ""
   ctxt������.Text = ""
   ctxt��ϵ�绰.Text = ""
   ccmb��������.Text = ""
   Ccmb��ҵ���.Text = ""
   ctxt��λ��ַ.Text = ""
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
            End If
            ccmbSex.Text = lstrSex
    End If
    Exit Sub
errHandler:
    sfsub������ "ְҵ������", "frmAddregister", "Sub ctxt���֤��_lostfocus", Err.Number, Err.Description, True
End Sub
Private Sub subAddList()
    ccmb���������.ListIndex = 1
    Ccmb��������.ListIndex = 1
    
End Sub

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
'    sfsub������ "ְҵ�����沿��", "frmregister", "func��������ʼ��", Err.Number, Err.Description, True
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
    ChDir (App.Path)                '�ı䵱ǰĬ��·��ΪӦ�ó�������·��
    ret = Authenticate()
    If (ret) Then
       ret = ReadBaseInfos(iname, isex, folk, birthday, code, addr, agency, startdate, enddate)
       ctxtName.Text = Trim(Split(iname, "")(0))
        Timer2.Enabled = False

        '����⵽�����֤����ȡ���ݳɹ��󣬹ر�timer2
        ctxt���֤��.Enabled = True
        ctxt���֤��.SetFocus      '��setfocus��lostfocusĿ���ǣ���ȡ�����֤�ź����  ctxt���֤��.lostfocus �¼��������Լ�����Ա����䣬��������
        'Call sub��ȡ֤��
        ctxt���֤��.Text = Trim(code)
'        ctxt���֤��_KeyDown 13, 1
        ctxt���֤��.Enabled = False
        ctxtName.Enabled = True
        ctxtName.SetFocus
        ctxtName.Text = Trim(Split(iname, "")(0))
        ctxtName.Enabled = False
        ctxtסַ.Text = Trim(addr)
        ccmb����.Text = Trim(folk)
        Picture2.Picture = LoadPicture(App.Path & "\photo.bmp")
    End If
    Exit Sub
errHandler:
'    sfsub������ "ְҵ�����沿��", "frmregister", "timer2_timer", Err.Number, Err.Description, True
End Sub
Private Sub sub����()
    Dim lsql As String
    Dim SNO As String
    Dim name, sex, age As String
    Dim a, b, c, D, E, F As String
    Dim G, H, i, j, K, L, M, n As String
    Dim o, P, Q, R, S, T, U, V, W, X, Y, Z As String
            U = ccmb���������.Text
            V = Ccmb��������.Text
            W = cdtp����.Value
            X = Left(ccmb�Ļ��̶�.Text, 2)
            Y = Trim(Split(ccmbTemplate.Text, "-")(2))
            age = Trim(ctxtAge.Text)
            SNO = Trim(clblsysno.Text)
            name = ctxtName.Text
           sex = ccmbSex.Text
            a = ccmbUnit.Text
            b = ccmbΣ������.Text
            c = Right(ccmb����Դ.Text, 2)
            D = ccmbְҵ���.Text
            E = ccmb�ֹ���.Text
            F = ccmbְ��.Text
            G = ctxtΣ������.Text
            H = Trim(ctxt�������.Text)
            i = Trim(ctxt����.Text)
            j = Trim(ctxt�ʱ�.Text)
            K = Trim(ctxtסַ.Text)
            L = Ccmb���.Text
            M = Trim(ctxt�绰.Text)
            n = Trim(ctxt����.Text)
            o = Trim(ctxt������.Text)
            P = ctxt������.Text
            Q = ctxt��ϵ�绰.Text
            R = ccmb��������.Text
            S = Ccmb��ҵ���.Text
            T = ctxt��λ��ַ.Text
    lsql = "update ְҵ�����_�����Ա������Ϣ�� set ����='" & name & "',�Ա�='" & sex & "',����='" & age & "',��������='" & W & "',������='" & o & "',Σ������='" & Y & "',ְҵ����='" & D & "',����Դ='" & c & "',�ֹ���='" & E & "',ְ���ְ��='" & F & "',�������='" & H & "',����='" & n & "',ְҵΣ������='" & G & "',�绰����='" & M & "',סַ='" & K & "',�ʱ�='" & j & "',�Ļ��̶�='" & X & "',����='" & i & "',����='" & ccmb����.Text & "',���='" & L & "',��������='" & U & "',�������='" & V & "'where ϵͳ���='" & SNO & "'"
    dafuncGetData (lsql)
End Sub

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

    Select Case Operate
    Case "���"
        subClear
    
    Case "����"
        '����������֮ǰ�Ƚ�������Ϣ��������
        Dim lobj1 As Object
        Dim S As String
        '��λ����
        If ccmbUnit.Text = "" Then
            MsgBox "û��¼�뵥λ��Ϣ"
            Exit Sub
        Else
            Dim a As String
            Dim OB As Object
            Set OB = dafuncGetData("select * from  ��λ����_��λ������Ϣ�� where ��λ����='" & Trim(ccmbUnit.Text) & "'")
            If OB.RecordCount > 0 Then
                a = OB("������")
            ElseIf OB.RecordCount = 0 Then
                MsgBox "��û��¼�뵥λ��Ϣ���밴��λ����¼�뵥λ��Ϣ"
                Exit Sub
            Else
                MsgBox "��ѡ��ĵ�λ�ظ�¼��,��ɾ���ظ���λ����ѡ��"
                Exit Sub
            End If
        End If
        '��������Ա������Ϣ��
        S = "select * from ְҵ�����_�����Ա������Ϣ�� where ϵͳ���='" & clblsysno.Text & "'"
        Set lobj1 = dafuncGetData(S)
        If lobj1.RecordCount = 0 Then
            dafuncGetData ("insert into ְҵ�����_�����Ա������Ϣ�� values('" & clblsysno.Text & "','','','','','','','" & a & "','" & Trim(ccmbUnit.Text) & "','" & Now & "','','','','','','','','','','','','','','','',null,null,null,null,null,'','')")
            dafuncGetData ("update  ְҵ�����_�����Ա������Ϣ�� set ������ݺ���='" & ctxt���֤��.Text & "',��������='" & Now & "'where ϵͳ���='" & clblsysno.Text & "'")
        End If
        '�����������Ϣ��
        S = "select * from ְҵ�����_��������Ϣ�� where ϵͳ���='" & clblsysno.Text & "'"
        Set lobj1 = dafuncGetData(S)
        If lobj1.RecordCount = 0 Then
            dafuncGetData ("insert into ְҵ�����_��������Ϣ�� values('" & clblsysno.Text & "','','" & ccmbTemplate.Text & "','" & ccmb���������.Text & "','" & Ccmb��������.Text & "','" & Now & "',null,null,null,null,null,'0','',null,null,'',null,null,null)")
        End If
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
        sub����
            Set mcol�����Ŀ = New Collection
            mobj���.����.������ = ccmbTemplate.Text
        
        '���ϵͳ��ų���С��15�������ж��ǲ���ʧ��
        If Len(clblsysno.Text) < 15 Then
            MousePointer = 0
            MsgBox "ϵͳ��Ŵ������飡", vbInformation, "ϵͳ��ʾ"
            Exit Sub
        End If
        If Len(ctxt���֤��.Text) = 0 Then
            MsgBox ("δ�������֤�ţ��������浱ǰ���ݣ�")
            Exit Sub
        End If
        
        If dafuncGetData("select ������� from ְҵ�����_���������ݿ� where ������ݺ��� = '" & Trim(ctxt���֤��.Text) & "' and convert(varchar(10),�������,111) = convert(varchar(10),getdate(),111)  and ���״̬ <> 0 ").RecordCount > 0 Then
            MsgBox "ͬһ�����֤�ŵ��첻�ܶ��¼�룡����"
            subClear
             Exit Sub
        End If
        '�������֤��Ƭ
        If mblnTakePhoto Then
            Dim lobjPhoto As Object
            Set lobjPhoto = cctlCatchPhoto.Photo
'        ElseIf Not Picture1.Picture Is Nothing Then
'            Set lobjPhoto = Picture1.Picture
            pmsub����ͼƬ lobjPhoto, Trim(clblsysno.Text), "ְҵ�����"
         Else
            Set lobjRec���֤��Ƭ = CreateObject("ְҵ������.clsPersonExamed")
            lobjRec���֤��Ƭ.func�������֤��Ƭ Picture2.Image, clblsysno.Text, "ְҵ�����"
            lobjRec���֤��Ƭ.func�������֤��Ƭ Picture2.Image, clblsysno.Text & "IDcard", "ְҵ�����"
            Set lobjRec���֤��Ƭ = Nothing
        End If
        
        Set lobj������ = CreateObject("ְҵ������.clsmedicalexamsheet")
        lobj������.������ = ccmbTemplate.Text
        With mobj���
            .ϵͳ��� = Trim(clblsysno.Text)
            
            If .����.������ <> ccmbTemplate.Text Then
                .����.������ = ccmbTemplate.Text
            End If

              '����Ա������ 2015-12-25 by Ĳ��
'            mobj���.�����Ա.�Ա� = Trim(ccmbSex.Text)
            mobj���.�����Ա.���� = Trim(ctxtAge.Text)

            .�����Ա.ϵͳ��� = Trim(clblsysno.Text)
            .�����Ա.���� = ctxtName.Text
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

            If Not Picture2.Picture Is Nothing Then
                .�����Ա.��Ƭ = Picture2.Picture
        
            End If
            If Val(ctxtAge.Text) > 0 Then
                .�����Ա.�������� = DateAdd("yyyy", -Val(ctxtAge.Text), Date)
            Else
                '��������ַ������������䡣
                mobj����.sub���Ǽ���ֵ "�������", ctxtAge.Text
                mstrĬ������ = ctxtAge.Text
            End If
            .�����Ա.���� = ctxtAge.Text
            
            On Error Resume Next
            .�����Ա.������ݺ��� = ctxt���֤��.Text

            .�����Ա.�Ļ��̶� = Left(ccmb�Ļ��̶�.Text, 2)
            
            .�����Ա.���� = ccmb����.Text
'            If ccmbUnit.Text = "" Then
'                .�����Ա.��λ������ = ""
'            Else
'                If .�����Ա.��λ������ <> mstr��λ������ Then
'                    .�����Ա.��λ������ = mstr��λ������
'                End If
'            End If

            .������� = cdtpDate.Value
            .��������� = ccmb���������.Text
            .�������� = Ccmb��������.Text
            Set lobj������ = CreateObject("ְҵ������.clsmedicalexamsheet")
            lobj������.������ = Trim(ccmbTemplate.Text)
            If mcol�����Ŀ.Count = 0 Then
                Set mcol�����Ŀ = lobj������.������Ŀ��("")
            End If
            Set .col�����Ŀ = mcol�����Ŀ
            save�Ż��������Ŀ mcol�����Ŀ, Trim(clblsysno.Text)
            subInit�������״̬ mcol�����Ŀ, Trim(clblsysno.Text)
        End With
        
        Set mcol�����Ŀ = New Collection
       
        mobj���.����.������ = ccmbTemplate.Text
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
            lobjFile.func��ӡ�Թܱ�ǩ mobj���.ϵͳ���, mobj���.�����Ա.����, mobj���.�����Ա.�Ա�, mobj���.�����Ա.����, "����"
        End If
'        If objds1("��Ŀ") = "17" Then  '17 ����
'            '����������ֱ�Ӵ�ӡ���������Ǵ�ӡ������Ŀ  2015-12-10 by Ĳ��   ��
'            Dim lobject As Object
''            Set lobject = dafuncGetData("select distinct right(�����Ŀ,2) as ��Ŀ  from ְҵ�����_����ģ�������Ŀ�� where ��������='" & ccmbTemplate.Text & "'and �����Ŀ like '1702%'and �����Ŀ<>'17020'")
'            Set lobject = dafuncGetData("select distinct right(�����Ŀ,2) as ��Ŀ���,���� as ��Ŀ  from ְҵ�����_����ģ�������Ŀ�� a,ְҵ�����_�����Ŀ���ñ� b where a.��������='" & ccmbTemplate.Text & "'and a.�����Ŀ like '1702%'and a.�����Ŀ=b.���� and a.�����Ŀ<>'17020' and b.����='����'")
'            If lobject.RecordCount > 0 Then
'                Dim xiangmu As String
'                Dim zhongjian As String
'                Dim X As Integer
'                lobject.MoveFirst
'                For X = 0 To lobject.RecordCount - 1
'                zhongjian = zhongjian + "," + lobject("��Ŀ")
'                lobject.MoveNext
'                Next X
'                xiangmu = Right(zhongjian, Len(zhongjian) - 1)
'                 lobjFile.func��ӡ�Թܱ�ǩ mobj���.ϵͳ���, mobj���.�����Ա.����, mobj���.�����Ա.�Ա�, mobj���.�����Ա.����, xiangmu
'            End If
'            '2015-12-10 by  Ĳ��  ��
'        End If
        objds1.MoveNext
        Next i
        Unload frmProcess
        MousePointer = 0
        Dim strSQL As String
        strSQL = "update ְҵ�����_��������Ϣ�� set ���״̬= '2'  where ϵͳ���='" & mobj���.ϵͳ��� & "'"
        dafuncGetData (strSQL)

        '2015-3-15ֹͣ���� ��ΰ
        Dim ret
        ret = CloseComm()
        Unload Me
        
    Case "����"
        mblnTakePhoto = True
        If mblnTakePhoto Then
        '���³�ʼ������ؼ���
            cctlCatchPhoto.funcInitVideo
        End If

    Case "�˳�"
        sub���ر��
        Unload Me
    End Select
End Sub

Private Sub sub���ر��()
        If clblsysno.Text <> "" And Len(clblsysno.Text) = 15 Then
'            mobj���.Func�˻�ְҵ�����ϵͳ��� Trim(clblsysno.Text)
            Dim lobjRec1 As Object
            Dim lobjRec2 As Object
            dasubBeginTran
                
            '�����жϸ�ϵͳ��ŵļ�¼�Ƿ��Ѵ��ڡ�
            dasubSetQueryTimeout 6000
            Set lobjRec1 = dafuncGetData("select ϵͳ��� from ְҵ�����_��������Ϣ�� where ϵͳ���='" & clblsysno.Text & "'")
            If lobjRec1.RecordCount = 0 Then
                '��ϵͳ���û�б��棬�����˻ء�
                dasubSetQueryTimeout 6000
                Dim max As Integer
                Dim XXX As String
                Set lobjRec2 = dafuncGetData("select ����ֵ from ְҵ�����_ҵ��������Ϣ�� where ������Ŀ = '��������ˮ��' and datepart(yyyy,˵��) = datepart(yyyy,getdate())")
                max = Val(lobjRec2("����ֵ"))
                max = max - 1
                XXX = Str(max)
                dafuncGetData ("update ְҵ�����_ҵ��������Ϣ�� set ����ֵ='" & XXX & "' where ������Ŀ = '��������ˮ��' and datepart(yyyy,˵��) = datepart(yyyy,getdate())")
            End If
            dasubCommitTran
        End If
End Sub
