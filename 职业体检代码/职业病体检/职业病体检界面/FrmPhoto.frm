VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{2327075A-FED2-4AEF-82B8-DD0C2B1AC8E1}#2.0#0"; "dyCatchPhoto.ocx"
Object = "{F48D4DEC-1198-11D5-91BE-0050BA06B70C}#5.9#0"; "¼��ؼ�.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmPhoto 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ְҵ�������-����"
   ClientHeight    =   8505
   ClientLeft      =   240
   ClientTop       =   375
   ClientWidth     =   11520
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8938.523
   ScaleMode       =   0  'User
   ScaleWidth      =   11520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame SSTab1 
      Height          =   7095
      Left            =   240
      TabIndex        =   23
      Top             =   840
      Width           =   10935
      Begin VB.TextBox ccmbSex 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   58
         Top             =   3480
         Width           =   735
      End
      Begin VB.ComboBox ccmb������ 
         Height          =   300
         Left            =   5880
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   720
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.ComboBox ccmb������� 
         Height          =   300
         Left            =   6960
         TabIndex        =   38
         Top             =   720
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.ComboBox ccmbTemplate 
         Enabled         =   0   'False
         Height          =   300
         Left            =   240
         TabIndex        =   37
         Top             =   1800
         Width           =   3480
      End
      Begin VB.TextBox ctxtName 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   300
         Left            =   2520
         TabIndex        =   36
         Top             =   2760
         Width           =   1410
      End
      Begin VB.TextBox ctxtAge 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   35
         Top             =   3480
         Width           =   495
      End
      Begin VB.TextBox ctxt���֤�� 
         Height          =   300
         Left            =   240
         TabIndex        =   34
         Top             =   2760
         Width           =   2130
      End
      Begin VB.TextBox clblsysno 
         Height          =   270
         Left            =   240
         TabIndex        =   33
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Frame frmPhoto 
         Caption         =   "����"
         ClipControls    =   0   'False
         ForeColor       =   &H00800000&
         Height          =   4275
         Left            =   5760
         TabIndex        =   29
         Top             =   1560
         Width           =   4905
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00E0E0E0&
            Height          =   1935
            Left            =   3240
            ScaleHeight     =   1875
            ScaleWidth      =   1515
            TabIndex        =   31
            Top             =   360
            Width           =   1575
         End
         Begin VB.CommandButton ccmdTakePhotoAgain 
            Caption         =   "����ȡ��"
            Height          =   495
            Left            =   3240
            TabIndex        =   30
            Top             =   2280
            Visible         =   0   'False
            Width           =   1575
         End
         Begin dyCatchPhoto.ctlCatchPhoto cctlCatchPhoto 
            Height          =   3570
            Left            =   240
            TabIndex        =   32
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
      Begin VB.ComboBox ccmb��������� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2760
         TabIndex        =   27
         Top             =   1080
         Width           =   1695
      End
      Begin VB.ComboBox Ccmb�������� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4680
         TabIndex        =   26
         Top             =   1080
         Width           =   1935
      End
      Begin VB.PictureBox Picture2 
         Height          =   1935
         Left            =   4080
         ScaleHeight     =   1875
         ScaleWidth      =   1515
         TabIndex        =   25
         Top             =   1800
         Width           =   1575
      End
      Begin VSFlex8Ctl.VSFlexGrid cgrdHistory 
         Height          =   2535
         Left            =   240
         TabIndex        =   24
         Top             =   4320
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
         Left            =   2280
         TabIndex        =   28
         Top             =   3480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   21233664
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
         Left            =   8640
         TabIndex        =   40
         Top             =   720
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   21233664
         CurrentDate     =   36950
         MaxDate         =   73050
         MinDate         =   17899
      End
      Begin ¼��ؼ�.ctlInputDictGrid ctlInputDictGrid1 
         Height          =   2535
         Left            =   8040
         TabIndex        =   41
         Top             =   6240
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Index           =   3
         Left            =   4680
         TabIndex        =   57
         Top             =   840
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Label4 
         Caption         =   "�����Ա���ͣ�"
         Height          =   255
         Left            =   2760
         TabIndex        =   56
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   240
         TabIndex        =   55
         Top             =   1560
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������ڣ�"
         Height          =   180
         Index           =   2
         Left            =   8640
         TabIndex        =   54
         Top             =   480
         Width           =   900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   2520
         TabIndex        =   53
         Top             =   2520
         Width           =   540
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         Height          =   180
         Left            =   240
         TabIndex        =   52
         Top             =   3240
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���䣺"
         Height          =   180
         Index           =   6
         Left            =   1320
         TabIndex        =   51
         Top             =   3240
         Width           =   540
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "���֤�ţ�"
         Height          =   180
         Left            =   240
         TabIndex        =   50
         Top             =   2520
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ϵͳ��ţ�"
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   49
         Top             =   840
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "ע��ˢ����ǰ��ȷ���ı���������Ϊ��"
         Height          =   180
         Left            =   240
         TabIndex        =   48
         Top             =   480
         Visible         =   0   'False
         Width           =   3060
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C000&
         Caption         =   "�ǿ���¼��ʱ��ɫΪ��¼�����¼��ʱֻ��ˢ�������֤"
         Height          =   180
         Left            =   240
         TabIndex        =   47
         Top             =   240
         Visible         =   0   'False
         Width           =   4680
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "�������ڣ�"
         Height          =   180
         Left            =   2280
         TabIndex        =   46
         Top             =   3240
         Width           =   900
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "ע����ˢ���룬��ˢ���֤"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   7560
         TabIndex        =   45
         Top             =   6000
         Visible         =   0   'False
         Width           =   2160
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "�뽫�������֤���ڶ������ϣ�"
         Height          =   180
         Left            =   240
         TabIndex        =   44
         Top             =   2280
         Width           =   2520
      End
      Begin VB.Label clblHintCheck 
         Caption         =   "ע�⣺У��֮��ֻ�������࣬�������ݼ�ʹ�޸ģ�Ҳ���ᱣ�档"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3360
         TabIndex        =   43
         Top             =   480
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.Label clblHistory 
         Caption         =   "˫���У�������������Ϣ�͸�����Ϣ��"
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   4080
         Visible         =   0   'False
         Width           =   3375
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   350
      Left            =   5280
      Top             =   480
   End
   Begin VB.CheckBox Check���֤ 
      Caption         =   "ˢ�������֤"
      Height          =   255
      Left            =   8520
      TabIndex        =   17
      Top             =   480
      Value           =   1  'Checked
      Visible         =   0   'False
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
      Interval        =   3
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
         TabIndex        =   22
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
Attribute VB_Name = "FrmPhoto"
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

'Private Sub cchk¼�뵥λ����_Click()
'    Dim lblnVisible As Boolean
'    On Error Resume Next
'    If cchk¼�뵥λ����.Value = 1 Then
'        lblnVisible = True
'    Else
'        lblnVisible = False
'    End If
'    ccmbUnit.Visible = lblnVisible
'    ccmd��λ��λ.Visible = lblnVisible
'    Label2(5).Visible = lblnVisible
'    ctxtName.SetFocus
'End Sub

'Private Sub ccmbSex_GotFocus()
'    On Error Resume Next
'    If ccmbSex = "" And ccmbSex.ListCount > 0 Then
'        ccmbSex.ListIndex = 0
'    End If
'End Sub
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
    If Len(clblsysno.Text) = 0 Then
        clblsysno.Text = lobj�������.Func����ְҵ�����ϵͳ��� & (ccmb���������.ListIndex + 1)
    Else
        clblsysno.Text = Left(clblsysno.Text, Len(clblsysno.Text) - 1) & (ccmb���������.ListIndex + 1)
    End If
    mobj���.ϵͳ��� = Trim(clblsysno.Text)
    mobj���.�����Ա.ϵͳ��� = Trim(clblsysno.Text)
    '2012-06-14 �ڵ�� ��
    '2012-12-18 ������  ��
    'BUG�ţ�0000092
'    If InStr(ccmb���������.Text, "����") > 0 Then
        Ccmb��������.Text = "�ڸ��ڼ�"
'    Else
'        Ccmb��������.ListIndex = 0
'    End If
    Call Ccmb��������_Click
    '2012-12-18 ������  ��
    Exit Sub
errHandler:
    sfsub������ "ְҵ������", "frmregister", "Private Sub ccmb���������_Click", Err.Number, Err.Description, True
End Sub

'Private Sub ccmb����Դ_click()
'    ccmbְҵ���.Visible = True
'    Call funcְҵ���
'    Exit Sub
'End Sub

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
'    lstrSysNo = clblsysno.Text
    clblsysno.Text = cgrdHistory.TextMatrix(cgrdHistory.Row, 0)
    lstrSysNo = clblsysno.Text
    clblsysno_LostFocus
    clblsysno.Text = lstrSysNo
End Sub

'Private Sub ctxt������_Change()
'    If Len(Trim(ctxt������.Text)) > 50 Then
'        ctxt������.Text = Left(Trim(ctxt������.Text), 50)
'    End If
'End Sub

'Private Sub ctxt�绰_Change()
'    If Len(Trim(ctxt�绰.Text)) > 11 Then
'        ctxt�绰.Text = Left(Trim(ctxt�绰.Text), 11)
'    End If
'End Sub

'Private Sub ctxt����_Change()
'    If Len(Trim(ctxt����.Text)) > 2 Then
'        ctxt����.Text = Left(Trim(ctxt����.Text), 2)
'    End If
'End Sub

'Private Sub ctxt����_Change()
'    If Len(Trim(ctxt����.Text)) > 2 Then
'        ctxt����.Text = Left(Trim(ctxt����.Text), 2)
'    End If
'End Sub

'Private Sub ctxt��ϵ�绰_Change()
'    If Len(Trim(ctxt��ϵ�绰.Text)) > 11 Then
'        ctxt��ϵ�绰.Text = Left(Trim(ctxt��ϵ�绰.Text), 11)
'    End If
'End Sub

'Private Sub ctxt���֤��_KeyDown(KeyCode As Integer, Shift As Integer)
'    On Error Resume Next
'    If KeyCode = 13 Then
'        ctxtName.SetFocus
'        sub�鿴��ʷ��Ϣ (ctxt���֤��.Text)
'    End If
'End Sub
'Private Sub ccmbTemplate_KeyDown(KeyCode As Integer, Shift As Integer)
'    On Error Resume Next
'    If KeyCode = 13 Then
'        If ctxtTubeNo.Visible Then
'            ctxtTubeNo.SetFocus
'        Else
'            ctxt��쵥��.SetFocus
'        End If
'    End If
'End Sub

'���ܣ����Ʋ��������������ƣ�ֻ��ѡ��
Private Sub ccmbTemplate_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

'Private Sub ccmbUnit_Click()
'    On Error GoTo errHandler
'    Dim i As Integer
'
'    '�ж�¼��ĵ�λ�Ƿ����б��д��ڣ�������������б�
'    i = gffuncItemIsInComboBox(ccmbUnit, ccmbUnit.Text)
'    If i = -1 Then
'        '���뵽�б����
'        ccmbUnit.AddItem ccmbUnit.Text
'
'        '���ص��������䲾�ļ���
'        pobjҵ�����.���չ������䲾.sub���ӵ�λ���� mstr��λ������ & "|" & ccmbUnit.Text
'    Else
'        '�޸ģ�2001-8-23��
'        On Error Resume Next
'        mstr��λ������ = pobjҵ�����.���չ������䲾.��λ���(ccmbUnit.Text)
'        sub��ʾ��λ���� ciptBase, mstr��λ������, mobjGUI
'    End If
'    Exit Sub
'errHandler:
'    sfsub������ "ְҵ������", "frmregister", "Sub ccmbUnit_Click", Err.Number, Err.Description, True
'
'End Sub

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
'    If Check���֤.Value = 0 Then   '����
        Timer2.Enabled = False
'        If mintState = 1 Then  'mintstate=1��ʾͨ��У�ˣ���������
            Picture1.Visible = False
            cctlCatchPhoto.Visible = True
            cctlCatchPhoto.funcInitVideo
            cctlCatchPhoto.Enabled = True
            mblnTakePhoto = True
            '2012-04-14 �ڵ�� ��
            '����ˢ�������֤ʱ������������Ƭ��Ҳ����������ͷ����
'            ctbMain.Buttons(4).Enabled = True
            '2012-04-14 �ڵ�� ��
'        End If
        
        '2012-07-11 �ڵ�� ��
        'ˢ�������֤���������Ա����䣬�������ڣ����֤��  enabled=false ,����true
        'У��֮��(mintstate=1)���������޸Ļ�����Ϣ��
        '�������Ͽ��ƣ�ֻ��Ϊ�˲鿴ʱ������Ϊ�����޸ġ�ʵ���ϣ�ֻ�д�ʱ���յ���Ƭ���Ա��档��
'        ctxt���֤��.Enabled = (mintState <> 1) 'True
'        ctxtName.Enabled = (mintState <> 1)  'True
'        ccmbSex.Enabled = (mintState <> 1)  'True
'        ctxtAge.Enabled = (mintState <> 1)  'True
'        cdtp����.Enabled = (mintState <> 1)  'True
        '2012-07-11 �ڵ�� ��
        
        Label31.Visible = False
'    Else                            'ˢ���֤
'        Picture1.Visible = True
'        cctlCatchPhoto.Visible = False
'        ctxt���֤��.Enabled = False
'        ctxtName.Enabled = False
'        ccmbSex.Enabled = False
'        ctxtAge.Enabled = False
'        cdtp����.Enabled = False
'        Label31.Visible = True
'        If mblnTakePhoto Then
'            cctlCatchPhoto.subDisconnect
'            mblnTakePhoto = False
'        End If
        '2012-04-14 �ڵ�� ��
        '��ˢ�������֤ʱ������������Ƭ
'        ctbMain.Buttons(4).Enabled = False
        '2012-04-14 �ڵ�� ��
'    End If
    
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
        '���³�ʼ������ؼ���
         If cctlCatchPhoto.Status = "�ָ�" Then
            cctlCatchPhoto.subת��״̬
            cctlCatchPhoto.subClear
         End If
        '����ϵͳ��Ų�����Ϣ
        subLoad clblsysno
'        Timer2.Enabled = True
    End If
End Sub
''����ϵͳ��Ų��ұ����Ϣ
Private Sub subLoad(ByVal paraϵͳ��� As String)
    Dim lobjRec As Object
     On Error GoTo errHandler
        Set lobjRec = dafuncGetData("select �������,������,������,��������,������ݺ���,����,�Ա�,����,�������� from ְҵ�����_���������ݿ� where ϵͳ���='" & paraϵͳ��� & "'")
        If Not (lobjRec.BOF Or lobjRec.EOF) Then
            mobj���.ϵͳ��� = Trim(clblsysno.Text)
            ccmb���������.Text = IIf(IsNull(lobjRec!�������), "", lobjRec!�������)
            Ccmb��������.Text = IIf(IsNull(lobjRec!������), "", lobjRec!������)
            ccmbTemplate.Text = IIf(IsNull(lobjRec!������), "", lobjRec!������)
            '2015-10-16
            'cdtpDate.Value = IIf(IsNull(lobjRec!��������), "", lobjRec!��������)
            cdtpDate.Value = Now
            ctxt���֤�� = IIf(IsNull(lobjRec!������ݺ���), "", lobjRec!������ݺ���)
            ctxtName = IIf(IsNull(lobjRec!����), "", lobjRec!����)
            ccmbSex = IIf(IsNull(lobjRec!�Ա�), "", lobjRec!�Ա�)
            ctxtAge = IIf(IsNull(lobjRec!����), "", lobjRec!����)
            cdtp���� = IIf(IsNull(lobjRec!��������), "", lobjRec!��������)
        Else
            MsgBox "��Ч���룡", vbExclamation, "ϵͳ��ʾ��"
            subClear
        End If
     Exit Sub
errHandler:
    sfsub������ "Frmphoto", "frmregister", "Sub subLoad", Err.Number, Err.Description, True
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
'    clblSysNo.Enabled = False
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
        cdtpDate.Value = IIf(IsNull(lobjRec!��������), "", lobjRec!��������)
        ctxt���֤�� = IIf(IsNull(lobjRec!������ݺ���), "", lobjRec!������ݺ���)
        ctxtName = IIf(IsNull(lobjRec!����), "", lobjRec!����)
        ccmbSex = IIf(IsNull(lobjRec!�Ա�), "", lobjRec!�Ա�)
        ctxtAge = IIf(IsNull(lobjRec!����), "", lobjRec!����)
        cdtp���� = IIf(IsNull(lobjRec!��������), "", lobjRec!��������)
'        ctxt���� = IIf(IsNull(lobjRec!����), "", lobjRec!����)
'        ctxt�ʱ� = IIf(IsNull(lobjRec!�ʱ�), "", lobjRec!�ʱ�)
'        ctxtסַ = IIf(IsNull(lobjRec!סַ), "", lobjRec!סַ)
'        ccmb�Ļ��̶� = IIf(IsNull(lobjRec!�Ļ��̶�), "", lobjRec!�Ļ��̶�)
'        Ccmb��� = IIf(IsNull(lobjRec!���), "", lobjRec!���)
'        ccmb���� = IIf(IsNull(lobjRec!����), "", lobjRec!����)
'        ctxt�绰 = IIf(IsNull(lobjRec!�绰����), "", lobjRec!�绰����)
'        ctxt���� = IIf(IsNull(lobjRec!����), "", lobjRec!����)
'        ctxt������ = IIf(IsNull(lobjRec!������), "", lobjRec!������)
'        ccmb����Դ = IIf(IsNull(lobjRec!����Դ), "", lobjRec!����Դ)
'        ccmbְҵ��� = IIf(IsNull(lobjRec!ְҵ����), "", lobjRec!ְҵ����)
'        ccmbΣ������ = IIf(IsNull(lobjRec!Σ������), "", lobjRec!Σ������)
'        ccmb�ֹ��� = IIf(IsNull(lobjRec!�ֹ���), "", lobjRec!�ֹ���)
'        ccmbְ�� = IIf(IsNull(lobjRec!ְ���ְ��), "", lobjRec!ְ���ְ��)
'        ctxtΣ������ = IIf(IsNull(lobjRec!ְҵΣ������), "", lobjRec!ְҵΣ������)
'        ctxt������� = IIf(IsNull(lobjRec!�������), "", lobjRec!�������)
        str��λ������ = IIf(IsNull(lobjRec!��λ������), "", lobjRec!��λ������)
        
        '��ȡ��Ƭ
'        Set lobjRec = CreateObject("ְҵ������.clspersonexamed")
'        lobjRec.ϵͳ��� = Trim(clblsysno.Text)
'        If lobjRec.��Ƭ <> 0 Then
'            Picture1.Picture = lobjRec.��Ƭ
'            Picture1.Visible = True
'            cctlCatchPhoto.Visible = False
'            If mblnTakePhoto Then
'                cctlCatchPhoto.subDisconnect
'                mblnTakePhoto = False
'            End If
'        End If
        
'''        '2012-07-11 �ڵ�� ��
'''        '������ȡ�ֳ���Ƭ
'''        Set lobj�ֳ���Ƭ = lobjRec.func��ȡ�ֳ���Ƭ(Trim(clblsysno.Text), "ְҵ�����")
'''        If Not lobj�ֳ���Ƭ Is Nothing Then Picture1.Picture = lobj�ֳ���Ƭ
'''        Picture1.Visible = True
'''        '2012-07-11 �ڵ�� ��
        
        
        '2012-06-13 �ڵ�� ��
        '��ȡ�������Ա���֤��Ƭ
'        Set lobj���֤��Ƭ = lobjRec.func�������֤��Ƭ(Trim(clblsysno.Text) & "IDcard", "ְҵ�����")
'        If Not lobj���֤��Ƭ Is Nothing Then
'            Picture2.Picture = lobj���֤��Ƭ
'            If mintState = 1 And mblnTakePhoto = False Then ccmdTakePhotoAgain.Visible = True
'        End If
'        Picture2.Visible = True
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
'Private Sub ctxt���֤��_lostfocus()
'    Dim ldatBirth As String
'    Dim lstrSex As String
'    On Error GoTo errHandler
'    If Trim(ctxt���֤��.Text) <> "" Then
'            '��ȷʱ�����֤���л�ȡ�������ڡ�
'            sub���ݹ�����ݺ����ȡ���պ��Ա� ctxt���֤��.Text, ldatBirth, lstrSex
'            If Not IsDate(ldatBirth) Then
'                MsgBox ("���֤�Ų��Ϸ���")
'                Exit Sub
'            End If
'
'            '�����Ƿ���Ҫ¼��������ڣ���Ҫʱ�Զ��������֤����д��������
'            On Error Resume Next
'            If IsDate(ldatBirth) Then
'                cdtp����.Value = ldatBirth
''                ctxtAge.Text = DateDiff("yyyy", ldatBirth, Date)
'                ctxtAge.Text = Year(Date) - Year(ldatBirth)
''����� 2012-12-11 ��
''˵�����������ж����֤�������Ƿ���˵�ǰ���ڣ��������һ�ꡣ
''             If Month(Date) > Month(ldatBirth) Then
''                 ctxtAge.Text = Year(Date) - Year(ldatBirth) + 1
''             ElseIf Month(Date) = Month(ldatBirth) Then
''                If Day(Date) >= Day(ldatBirth) Then
''                    ctxtAge.Text = Year(Date) - Year(ldatBirth) + 1
''                Else
''                    ctxtAge.Text = Year(Date) - Year(ldatBirth)
''                End If
''             Else
''                 ctxtAge.Text = Year(Date) - Year(ldatBirth)
''             End If
'
'    '����� 2012-12-11 ��
'            End If
'            ccmbSex.Text = lstrSex
'    End If
'    Exit Sub
'errHandler:
'    sfsub������ "ְҵ������", "frmregister", "Sub ctxt���֤��_lostfocus", Err.Number, Err.Description, True
'End Sub

'Private Sub ctxt��쵥��_KeyDown(KeyCode As Integer, Shift As Integer)
'    On Error Resume Next
'    If KeyCode = 13 Then
'        If ccmbUnit.Visible Then
'            ccmbUnit.SetFocus
'        Else
'            If ctxt����.Visible Then
'                ctxt����.SetFocus
'            Else
'                If ciptBase.Visible Then
'                    ciptBase.SetFocus
'                End If
'            End If
'        End If
'    End If
'End Sub
'
'Private Sub ctxt�ʱ�_Change()
'    If Len(Trim(ctxt�ʱ�.Text)) > 6 Then
'        ctxt�ʱ�.Text = Left(Trim(ctxt�ʱ�.Text), 6)
'    End If
'End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If mblnTakePhoto Then
        '���³�ʼ������ؼ���
        cctlCatchPhoto.funcInitVideo
    End If
    'ctxtName.SetFocus
End Sub

'Private Sub Form_Deactivate()
'    On Error Resume Next
'    gfsubHideComboList ccmbUnit
'End Sub
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
        .Add "����"
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
    
    If ���ʱ�־ = 2 Then
        ccmb���������.Enabled = False
        Ccmb��������.Enabled = False
        ccmbTemplate.Enabled = False
        ctbMain.Buttons(1).Enabled = False
    End If

    '���
    subClear
    cdtpDate.Value = Now
    
    '�·���ϵͳ���
    'clblsysno.Caption = mobj���.Func����ϵͳ���
    'clblSysNo.Text = ""
    mobj���.ϵͳ��� = Trim(clblsysno.Text)
    'pstrϵͳ��� = Trim(clblSysNo.Text)
'    If Check���֤.Value = 0 Then
'        Check���֤.Value = 1
'    Else
'        Check���֤_Click
'    End If
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
'    clblSysNo.Enabled = True

    'Ϊ�˼ӿ촰������ٶȣ����³�ʼ���������ڶ�ʱ������ɡ�
'    Timer1.Enabled = True
'    Timer2.Enabled = True
    subLoad pstrPhoto
        'У��
'    mintState = 1
'    Check���֤.Value = 1
    Check���֤_Click
'    mintState = 1
    pobjҵ�����.funcд�뵥�˵�ǰ���״̬ clblsysno, mintState
    pobjҵ�����.funcд��У������Ϣ clblsysno, um�û����
    MousePointer = 0
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
'        ccmbUnit.AddItem lcolInfo(i)
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
'        cchk¼�뵥λ����.Value = 1
    Else
'        cchk¼�뵥λ����.Value = 0
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
'    If lstrTmp = "δУ��" Or lstrTmp = "" Or (���ʱ�־ = 1 And lstrTmp <> "δ���嵥") Then   'ֻ�е��Ǽ�δУ��ʱ����ˢ���֤
'        If Check���֤.Value = 0 Then
'            Check���֤.Value = 1
'        Else
'            Check���֤_Click
'        End If
'        mintState = 0
''        clblHintCheck.Visible = False
'        sub��������ʼ��
'        func���֤��֤
'    ElseIf lstrTmp = "δ���嵥" Then
'        mintState = 1
'        clblHintCheck.Visible = True
'        'If ���ʱ�־ = 0 Then Check���֤.Value = 0:Check���֤_Click
'        If Check���֤.Value = 1 Then
'            Check���֤.Value = 0
'        Else
'            Check���֤_Click
'        End If
'    ElseIf lstrTmp = "������" Then  '�����������ҵ���δ���嵥���ʼ��������ˢ���֤���ܡ�
If lstrTmp = "������" Then  '�����������ҵ���δ���嵥���ʼ��������ˢ���֤����
        mintState = 1
        clblHintCheck.Visible = True
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
'    ccmb�Ļ��̶�.Clear
'    ccmb�Ļ��̶�.AddItem ""
'    For i = 1 To lobjRec.RecordCount
'        ccmb�Ļ��̶�.AddItem lobjRec("����")
'        ccmb�Ļ��̶�.ItemData(ccmb�Ļ��̶�.NewIndex) = lobjRec("���")
'        lobjRec.MoveNext
'    Next
'    ccmb�Ļ��̶�.ListIndex = 0
'
'    '��ȡ���
'    Set lobjRec = pobjDict.FetchEx("�����ֵ�")
'    Ccmb���.Clear
'    Ccmb���.AddItem ""
'    For i = 1 To lobjRec.RecordCount
'        Ccmb���.AddItem lobjRec("����")
'        Ccmb���.ItemData(Ccmb���.NewIndex) = lobjRec("���")
'        lobjRec.MoveNext
'    Next
'    Ccmb���.ListIndex = 0
    
'     '��ȡ����
'    Set lobjRec = pobjDict.FetchEx("�����ֵ�")
'    ccmb����.Clear
'    ccmb����.AddItem ""
'    For i = 1 To lobjRec.RecordCount
'        ccmb����.AddItem lobjRec("����")
'        ccmb����.ItemData(ccmb����.NewIndex) = lobjRec("���")
'        lobjRec.MoveNext
'    Next
'    ccmb����.ListIndex = 0
'
'     '��ȡ��������
'    Set lobjRec = pobjDict.FetchEx("���������ֵ�")
'    ccmb��������.Clear
'    ccmb��������.AddItem ""
'    For i = 1 To lobjRec.RecordCount
'        ccmb��������.AddItem lobjRec("����")
'        ccmb��������.ItemData(ccmb��������.NewIndex) = lobjRec("���")
'        lobjRec.MoveNext
'    Next
'    ccmb��������.ListIndex = 0
'
'    '��ȡΣ������
'    Set lobjRec = pobjDict.FetchEx("Σ�������ֵ�")
'    ccmbΣ������.Clear
'    ccmbΣ������.AddItem ""
'    For i = 1 To lobjRec.RecordCount
'        ccmbΣ������.AddItem lobjRec("����")
'       ccmbΣ������.ItemData(ccmbΣ������.NewIndex) = lobjRec("���")
'        lobjRec.MoveNext
'    Next
'    ccmbΣ������.ListIndex = 0
'
'    '��ȡְҵ��ְ��
'    Set lobjRec = pobjDict.FetchEx("ְҵ��ְ���ֵ�")
'    ccmbְ��.Clear
'    ccmbְ��.AddItem ""
'    For i = 1 To lobjRec.RecordCount
'        ccmbְ��.AddItem lobjRec("����")
'        ccmbְ��.ItemData(ccmbְ��.NewIndex) = lobjRec("���")
'        lobjRec.MoveNext
'    Next
'    ccmbְ��.ListIndex = 0
'
'    '��ȡ�����ֵ�
'    Set lobjRec = pobjDict.FetchEx("�����ֵ�")
'    ccmb�ֹ���.Clear
'    ccmb�ֹ���.AddItem ""
'    For i = 1 To lobjRec.RecordCount
'        ccmb�ֹ���.AddItem lobjRec("����")
'        ccmb�ֹ���.ItemData(ccmb�ֹ���.NewIndex) = lobjRec("���")
'        lobjRec.MoveNext
'    Next
'    ccmb�ֹ���.ListIndex = 0
'    Call func��ȡ��ҵ����ֵ�
'    Call func����Դ
'    clblSysNo.Visible = True
'    clblSysNo.SetFocus
     '2012-07-11 �ڵ�� ��
    'ϵͳ�������֮�󣬲����Ըı䡣ת��focus֮ǰ���������趨setfocus��������֤������Ϣ���Լ��ؽ�ȥ��
    '�Ҽ�����󣬿ؼ�enabled=false
    'If ���ʱ�־ = 1 Then
'        ctxt����.SetFocus
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
'Private Sub ccmbUnit_GotFocus()
'    On Error GoTo errHandler
''    gfsubShowComboList ccmbUnit
'    Exit Sub
'errHandler:
'    'sfsub������ "ְҵ�����沿��", "FrmRegisterAnnual", "ccmbUnit_GotFocus", Err.Number, Err.Description, False
'End Sub

'Private Sub ccmbUnit_LostFocus()
'    On Error GoTo errHandler
'    Dim i As Integer
'    If Trim(ccmbUnit.Text) = "" Then Exit Sub
'
'    '�ж�¼��ĵ�λ�Ƿ����б��д��ڣ�������������б�
'    i = gffuncItemIsInComboBox(ccmbUnit, ccmbUnit.Text)
'    If i = -1 Then
'        '���뵽�б����
'        ccmbUnit.AddItem ccmbUnit.Text
'
'        '���ص��������䲾�ļ���
'        pobjҵ�����.���չ������䲾.sub���ӵ�λ���� mstr��λ������ & "|" & ccmbUnit.Text
'    Else
'        '�޸ģ�2001-8-26������λ�����Ų�ͬ���޸Ĺ������䲾����
'        If mstr��λ������ <> pobjҵ�����.���չ������䲾.��λ���(ccmbUnit.Text) And mstr��λ������ <> "" Then
'            pobjҵ�����.���չ������䲾.sub���ӵ�λ���� mstr��λ������ & "|" & ccmbUnit.Text
'        End If
'    End If
'    Exit Sub
'errHandler:
'    sfsub������ "ְҵ������", "frmregister", "Sub ccmbUnit_LostFocus", Err.Number, Err.Description, True
'End Sub


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
'Private Sub ccmd��λ��λ_Click()
'    On Error GoTo errHandler
'    Dim lobjRec As Object  '��λ��λ���صĽ����¼��
'    Dim lobj��λ As Object
'    Dim lobj��λ��Ϣ As Object
'    '������λ��λ���档
'    Set lobjRec = pobjҵ�����.func��λ��λ
'    '��ȡ��λ�ĵ�λ����ʾ�ڡ���λ���ơ�¼����С�
'    If Not lobjRec Is Nothing Then
'        If lobjRec.RecordCount > 0 Then
'            ccmbUnit.Text = IIf(IsNull(lobjRec("��λ����")), "", lobjRec("��λ����"))
'            mstr��λ������ = lobjRec!������
'            'Set lobj��λ = CreateObject("ְҵ������.class1")
'            'lobj��λ.��λ��Ϣ���� = lobjRec!������
'            'Set lobj��λ��Ϣ���� = lobj��λ.��λ��Ϣ
'
'
'
'            If mstr��λ������ <> "" Then
'                '�޸ģ�2001-8-23����ʾ��λ���ԣ���
'                On Error Resume Next
'                'sub��ʾ��λ���� ciptBase, mstr��λ������, mobjGUI
'                func��ȡ��λ��Ϣ lobjRec!������
'            End If
'        End If
'    End If
'
'    '�ѽ���ص���λ¼��򡣱����ܱ����µ�λ��λ��Ϣ��
'    ccmbUnit.SetFocus
'    SendKeys vbTab
'    Exit Sub
'errHandler:
'    Dim lstrError As String
'    lstrError = func������(Err.Number, Err.Description)
'    sfsub������ "ְҵ�����沿��", "FrmRegisterAnnual", "ccmd��λ��λ_Click", 6666, lstrError, False
'End Sub

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
    clblsysno.Text = ""
    ctxt���֤��.Text = ""
    ctxtName.Text = ""
    ccmbSex.Text = ""
    ctxtAge = ""
    ctxtTubeNo = ""
    ctxt��쵥�� = ""
    mstr��λ������ = ""
    Ccmb�������� = ""
    ccmbTemplate.Text = ""
    ccmb���������.Text = ""
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
    
'    mobj����.sub���Ǽ���ֵ "���Ǽ�ʱ¼�뵥λ����", IIf(cchk¼�뵥λ����.Value = 1, "��", "��")
     
    Set mobj��� = Nothing
    Set mobj��켯 = Nothing
    Set mobj����ģ�� = Nothing
    '�ر������
    If mblnTakePhoto Then
        cctlCatchPhoto.subDisconnect
    End If
    mblnInUse = False
    pstrϵͳ������� = ""
    Dim ret
    
    ret = CloseComm()
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
    Dim lobjFile As Object
    '2012-06-13 �ڵ�� ��
    
    On Error GoTo errHandler
    
    Select Case Operate
    
    Case "���"
        subClear
        '��ս���������ţ���ʾ��¼��������Ա��
        mobj���.�����Ա.����������� = ""
        clblsysno.Text = ""
        Cancel = True
        
    Case "����"
        '2012-07-11 �ڵ�� ��
        '���У��ͨ����ֻ�ܴ洢�ֳ���Ƭ
        Cancel = True
        MousePointer = 11
'        If mintState = 1 And ���ʱ�־ = 1 Then
       
        '2015-3-13 liuwei �޸ı�����Ƭ
          Dim lobjPhoto As StdPicture
          Dim strSQL As String
        If mblnTakePhoto Then
                Set lobjPhoto = cctlCatchPhoto.Photo
            ElseIf Not Picture1.Picture Is Nothing Then
               Set lobjPhoto = Picture1.Picture
                
            End If
         pmsub����ͼƬ lobjPhoto, Trim(clblsysno.Text), "ְҵ�����"
         '2015-10-16   �������Ϊ��������
         strSQL = "update ְҵ�����_��������Ϣ�� set �������='" & Now & "' where ϵͳ���='" & clblsysno & "'"
         dafuncGetData strSQL
'         dafuncGetData ("update ְҵ�����_��������Ϣ�� set  �������='" & Now & "'where ϵͳ���='" & paraSysNo & "'")   '2015-10-16
        ���ʱ�־ = 0
'           clblSysNo.Text = mobj���.Func����ְҵ�����ϵͳ��� & (ccmb���������.ListIndex + 1)
        Set mcol�����Ŀ = New Collection
        mobj���.����.������ = ccmbTemplate.Text
        frmRegisterManage.sub��ѯ����ʾ
        Cancel = True
        cgrdHistory.rows = 1
        cgrdHistory.Visible = False
        clblHistory.Visible = False
'        Check���֤_Click


        '��ӱ�ǩ��ӡ   2015-11-30 by Ĳ�� ��
        With mobj���
        mobj���.ϵͳ��� = Trim(clblsysno.Text)
        mobj���.�����Ա.���� = Trim(ctxtName.Text)
        '����Ա������ 2015-12-25 by Ĳ��
        mobj���.�����Ա.�Ա� = Trim(ccmbSex.Text)
        mobj���.�����Ա.���� = Trim(ctxtAge.Text)
        End With
        Dim strsql1 As String
        strsql1 = "select distinct left(�����Ŀ,2) as ��Ŀ  from ְҵ�����_����ģ�������Ŀ�� where ��������='" & ccmbTemplate.Text & "'"
        Dim objds1 As Object
        Set objds1 = dafuncGetData(strsql1)
'        Dim lobjFile As Object
        Set lobjFile = CreateObject("ְҵ������.cls����")
        Dim zxcsysno As Collection
        Set zxcsysno = New Collection
        
        zxcsysno.Add (mobj���.ϵͳ���)
        lobjFile.func��ӡ��������嵥 Trim(clblsysno)
'        lobjFile.func��ӡ����嵥 zxcsysno
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
                '��ӡ��ǩ����Ŀ���ǹ̶��ģ��Ǵ�ӡ�������еĳ�����Ŀ 2015-12-25 by Ĳ��
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
'        Unload frmProcess
        MousePointer = 0
'        'Update ���״̬
''        Dim strSQL As String
'        strSQL = "update ְҵ�����_��������Ϣ�� set ���״̬= '2'  where ϵͳ���='" & mobj���.ϵͳ��� & "'"
'        dafuncGetData (strSQL)
        
'        Set lobjFile = CreateObject("ְҵ������.cls����")
'        lobjFile.func��ӡ��������嵥 Trim(clblsysno)
        pobjҵ�����.funcд�뵥�˵�ǰ���״̬ Trim(clblsysno), 2
                
        frmRegisterManage.sub��ѯ����ʾ
'           frmRegisterManage.sub��ʾ��ѯ���
        clblsysno.Text = ""
        subClear
         
        '���³�ʼ������ؼ���
        If cctlCatchPhoto.Status = "�ָ�" Then
           cctlCatchPhoto.subת��״̬
           cctlCatchPhoto.subClear
        End If
    End Select
    Set lobjrec���� = Nothing
    Set lobj������ = Nothing
    MousePointer = 0
    Unload Me         '������ɺ����������ʧ   2016-1-6 by Ĳ��
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
'        ccmbUnit.Text = .��λ����
'        ccmbUnit_LostFocus
        
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




'��ȡ��λ������Ϣ
Private Function func��ȡ��λ��Ϣ(��λ��� As String)
    Dim lobj��λ As Object
    On Error GoTo errHandler
    Set lobj��λ = dafuncGetData("select * from ��λ����_��λ������Ϣ�� where ������='" & ��λ��� & "'")
    If Not lobj��λ.RecordCount = 0 Then
'        ccmbUnit.Text = IIf(IsNull(lobj��λ("��λ����")), "", lobj��λ("��λ����"))
        mstr��λ������ = ��λ���
'        ctxt������.Text = IIf(IsNull(lobj��λ("������")), "", lobj��λ("������"))
'        ctxt��ϵ�绰.Text = IIf(IsNull(lobj��λ("�绰")), "", lobj��λ("�绰"))
'        ccmb��������.Text = IIf(IsNull(lobj��λ("��Ӫ��ʽ")), "", lobj��λ("��Ӫ��ʽ"))
'        Ccmb��ҵ���.Text = IIf(IsNull(lobj��λ("��λ���")), "", lobj��λ("��λ���"))
'        ctxt��λ��ַ.Text = IIf(IsNull(lobj��λ("��ַ")), "", lobj��λ("��ַ"))
    End If
    Exit Function
errHandler:
    sfsub������ "ְҵ������", "frmregister", "func��ȡ��λ��Ϣ", Err.Number, Err.Description, True
End Function

'�������֤����������ʼ����PC���ն˵�����
Private Sub sub��������ʼ��()
    'CVR_InitComm
    On Error GoTo errHandler
    'If Option1.Value = True Then
    '    List1.AddItem "�����ӻ��ߡ� ���� " & comS.ListIndex + 1
    '    List1.AddItem "���� " & CVR_InitComm(comS.ListIndex + 1)
    'Else
    '    List1.AddItem "�����ӻ��ߡ� USB�� " & comU.ListIndex + 1
    '   List1.AddItem " ���� " & CVR_InitComm(0 + 1001)
    '���Ӵ��ڣ�COM1~COM16����USB��(1001~1016)���Ӵ��ڣ�COM1~COM16����USB��(1001~1016)
      ' CVR_InitComm (0 + 1001)
    'End If
    Exit Sub
errHandler:
    sfsub������ "ְҵ�����沿��", "frmregister", "func��������ʼ��", Err.Number, Err.Description, True
End Sub

'ֻ�о������֤��֤�󣬲��ܶ�ȡ��Ϣ����Ƭ��
Private Function func���֤��֤() As Integer
'    'CVR_Authenticate
'    'On Error GoTo errHandler
'    'Dim temp As Integer
'    'List1.AddItem "�����֤��֤��"
'    'List1.AddItem " ���� " & CVR_Authenticate()
'    func���֤��֤ = CVR_Authenticate()
'    Exit Function
errHandler:
    sfsub������ "ְҵ�����沿��", "frmregister", "func���֤��֤", Err.Number, Err.Description, True
End Function




'��timer2 ���м���Ƿ������֤���ڶ������ϣ�����Ϊ350ms
Private Sub Timer2_Timer()
     '��֪Ϊ�Σ����ǵ����Ҳ���termb.dll�ļ������ǣ������޸Ĵ�����
'    'On Error GoTo errHandler
'    On Error Resume Next
'    '2012-07-11 �ڵ�� ��
'     Dim n, ret, nLen
'    Dim iname As String * 31
'    Dim isex  As String * 3
'    Dim folk As String * 10
'    Dim code As String * 19
'    Dim addr As String * 71
'    Dim birthday As String * 9
'    Dim startdate As String * 9
'    Dim enddate As String * 9
'    Dim agency As String * 31
'    Dim Msg As String * 300
'    Dim Msg1 As String * 256
'    Dim IINSNDN As String * 64
'    Dim SAMID As String * 36
'    Dim LenT As Integer
'    ChDir (App.Path)                '�ı䵱ǰĬ��·��ΪӦ�ó�������·��
'    ret = Authenticate()
'    If (ret) Then '
'       ret = ReadBaseInfos(iname, isex, folk, birthday, code, addr, agency, startdate, enddate)
'       If Trim(ctxtName.Text) <> Trim(Split(iname, "")(0)) Then
'       Dim msgs
'       msgs = MsgBox("���֤��Ϣ�������Ϣ��ƥ�䡣", vbOKOnly + vbInformation, "��ʾ")
'       Exit Sub
'       End If
'       Timer2.Enabled = False
'
'        '����⵽�����֤����ȡ���ݳɹ��󣬹ر�timer2
'        'Call sub��ȡ��Ϣ
'        ctxt���֤��.Enabled = True
'        ctxt���֤��.SetFocus      '��setfocus��lostfocusĿ���ǣ���ȡ�����֤�ź����  ctxt���֤��.lostfocus �¼��������Լ�����Ա����䣬��������
'        'Call sub��ȡ֤��
'        ctxt���֤��.Text = Trim(code)
'        ctxt���֤��_KeyDown 13, 1
'        ctxt���֤��.Enabled = False
'        ctxtName.Enabled = True
'        ctxtName.SetFocus
'        'Call sub��ȡ����
'        ctxtName.Text = Trim(Split(iname, "")(0))
'        ctxtName.Enabled = False
'        '2012-06-13 �ڵ�� ��
'        'ʡ������Ҫ������ͼƬ�����֤ͼƬ���洢
'        '��ʾ���֤ͼƬ
'        Picture2.Picture = LoadPicture(App.Path & "\photo.bmp")
'        '2012-06-13 �ڵ�� ��
'        'Call sub��ȡסַ
'        ctxtסַ.Text = Trim(addr)
'        'Call sub��ȡ����
'        ccmb����.Text = Trim(folk)
'        '2012-07-12 �ڵ�� ��
'        'ʵ�ֹر����֤�Ķ����������ظ���ʱ����
'        'ret = CloseComm()
'        '2012-07-12 �ڵ�� ��
'
'        '2012-07-15 �ڵ�� ��
'        'ÿ�εõ��µ����֤ʱ������������������е����ʱ���¼
'        sub�鿴��ʷ��Ϣ (Trim(ctxt���֤��.Text))
'        '2012-07-15 �ڵ�� ��
'
'
'       '���ʱ��2015-2-25
''�������ѯ�����Ա������Ϣ�����������ͺ��������Լ�Σ�����ء���ɵĽṹ�硰ְҵ���-�ڸ��ڼ�-�۳���
'
'Dim sΣ������ As String
'Dim s������� As String
'Dim s�������� As String
'Dim strs As String
'Dim strb  As String
'Dim strx As String
' Dim rs As Object
'strs = "select Σ������,�������,�������� from ְҵ�����_�����Ա������Ϣ�� where ϵͳ���='" & clblsysno.Text & "'"
'Set rs = dafuncGetData(strs)
'sΣ������ = rs("Σ������")
's������� = rs("�������")
's�������� = rs("��������")
'Dim s������Ϣ As String
's������Ϣ = s�������� + "-" + s������� + "-" + sΣ������
'ccmbTemplate.Text = s������Ϣ
'ccmb���������.Text = s��������
'Ccmb��������.Text = s�������
'   End If
'    Exit Sub
'errHandler:
'    sfsub������ "ְҵ�����沿��", "frmregister", "timer2_timer", Err.Number, Err.Description, True
'End Sub
'
''���֤��֤�Ժ󣬲Ŷ�ȡ��Ϣ��ֻ�ж�ȡ��Ϣ�Ժ󣬲Ż��ڵ�ǰĿ¼�������֤����Ƭ�ļ�zp.mbp
''Private Sub sub��ȡ��Ϣ()
''    'CVR_Read_Content
''    Dim mode As Integer
''    On Error GoTo errHandler
''    'modeȡֵ��
''    '1: ��������wz.txt?��Ƭ����xp.wlt����Ƭzp.bmp (����)
''    '2: ��������wz.txt����Ƭ����xp.wlt
''    '4: ����wz.txt(����)����Ƭzp.bmp(����)
''    '6: �������豸ģ�����������.txt�ļ�(����)����Ƭ.bmp�ļ�(����)
''    mode = 4
''    CVR_Read_Content (mode)
''    Exit Sub
''errHandler:
''    sfsub������ "ְҵ�����沿��", "frmregister", "sub��ȡ��Ϣ", Err.Number, Err.Description, True
''End Sub
'''���ʱ��2015-2-25
''   '�������ѯ�����Ա������Ϣ�����������ͺ��������Լ�Σ�����ء���ɵĽṹ�硰ְҵ���-�ڸ��ڼ�-�۳���
''   Private Sub sub��ȡ������Ϣ(ByVal paraϵͳ��� As String)
''
''Dim sΣ������ As String
''Dim s������ As String
''Dim s������� As String
''Dim strs As String
''strs = "select Σ������ from ְҵ�����_�����Ա������Ϣ�� where ϵͳ���='" & paraSysNo & "'"
''strb = "select ������ from ְҵ�����_�����Ա������Ϣ�� where ϵͳ���='" & paraSysNo & "'"
''strx = "select ������� from ְҵ�����_�����Ա������Ϣ�� where ϵͳ���='" & paraSysNo & "'"
''sΣ������ = dafuncGetData(strs)
''s������ = dafuncGetData(strb)
''s������� = dafuncGetData(strx)
''Dim s������Ϣ As String
''s������Ϣ = s������ + "-" + s������� + "-" + sΣ������
''ccmbTemplate.Text = s������Ϣ
''End Sub
'''��ȡ���֤��
''Private Sub sub��ȡ֤��()
''    Dim strTemp As String
''    Dim nReturnLen As Integer
''    Dim nReturn As Integer
''    strTemp = Space(255)
''    nReturn = GetPeopleIDCode(strTemp, nReturnLen)
''    ctxt���֤��.Text = Trim(strTemp)
''End Sub
''
''Private Sub sub��ȡ����()
''    Dim strTemp As String
''    Dim nReturnLen As Integer
''    Dim nReturn As Integer
''    strTemp = Space(255)
''    nReturn = GetPeopleName(strTemp, nReturnLen)
''    ctxtName.Text = Trim(strTemp)
''End Sub
''
''Private Sub sub��ȡסַ()
''    Dim strTemp As String
''    Dim nReturnLen As Integer
''    Dim nReturn As Integer
''    strTemp = Space(255)
''    nReturn = GetPeopleAddress(strTemp, nReturnLen)
'''    ctxtסַ.Text = Trim(strTemp)
''End Sub
''
''Private Sub sub��ȡ����()
''    Dim strTemp As String
''    Dim nReturnLen As Integer
''    Dim nReturn As Integer
''    strTemp = Space(10)
''    nReturn = GetPeopleNation(strTemp, nReturnLen)
'''    ccmb����.Text = Trim(strTemp)
''End Sub
'
''���ܣ�����ְҵ���Ǽ���ѡ��������Ŀ
''���ߣ�����
''ʱ�䣺2012-06-04
''˵��������Ҫ�鿴���ݿ������Ƿ�����ͬ�������Ŀ��Ȼ���ٽ������ӻ����޸�
'
'Public Sub save�Ż��������Ŀ(ByRef para�����Ŀ As Collection, ByVal paraϵͳ��� As String)
'    Dim lstrSql As String
'    Dim MedicProjt As String
'    Dim rs As Object
'    Dim i As Integer
'    Dim col�����Ŀ As Collection
'    On Error GoTo errHandler
'
'    Set rs = dafuncGetData("select ���� from ϵͳ����_�ֵ�_�ֵ����ݱ� where ID = (select ID from ϵͳ����_�ֵ�_�ֵ���б� where ����='ְҵ���������ֵ�') and ���� like '%��'")
'
'    For i = 1 To rs.RecordCount
'
'        lstrSql = "delete ְҵ�����_�����Ϣ_" & rs("����") & " where ϵͳ���='" & paraϵͳ��� & "'"
'        dafuncGetData lstrSql
'        rs.MoveNext
'    Next i
'
'    Set col�����Ŀ = para�����Ŀ
'
'    For i = 1 To col�����Ŀ.Count
'        MedicProjt = Left(Trim(col�����Ŀ(i).Item(1)), 2)
'
'        lstrSql = "select ���� from ϵͳ����_�ֵ�_�ֵ����ݱ� where ID = (select ID from ϵͳ����_�ֵ�_�ֵ���б� where ����='ְҵ���������ֵ�') and ���= '" & MedicProjt & "'"
'        Set rs = dafuncGetData(lstrSql)
'
'        lstrSql = "insert into ְҵ�����_�����Ϣ_" & rs("����") & "(ϵͳ���,�����Ŀ) values('" & paraϵͳ��� & "','" & col�����Ŀ(i).Item(1) & "')"
'        dafuncGetData lstrSql
'    Next i
'
'    Exit Sub
'errHandler:
'   sfsub������ "ְҵ��ʷ¼��", "clscareerhstregt", "public sub save�����Ŀ", Err.Number, Err.Description, False
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
'    strSQL = "select ϵͳ���,�������� ����ʱ��,������ ����,���״̬ from ְҵ�����_���������ݿ� where ������ݺ���='" & paraIDCard & "' and ��������<'" & Format(Now, "yyyy-mm-dd") & "'"
    strSQL = "select ϵͳ���,�������� ����ʱ��,������ ����,���״̬ from ְҵ�����_���������ݿ� where ���״̬='1'"
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
        initState(5) = "���½���"
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
'        ctxt����.SetFocus
        
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
'            .�����Ա.��λ���� = ccmbUnit.Text
'            .�����Ա.Σ������ = ccmbΣ������.Text
'            .�����Ա.����Դ = ccmb����Դ.Text
'            .�����Ա.ְҵ���� = ccmbְҵ���.Text
'            .�����Ա.�ֹ��� = ccmb�ֹ���.Text
'            .�����Ա.ְ���ְ�� = ccmbְ��.Text
'            .�����Ա.ְҵΣ������ = ctxtΣ������.Text
'            .�����Ա.������� = Trim(ctxt�������.Text)
'            .�����Ա.���� = Trim(ctxt����.Text)
'            .�����Ա.�ʱ� = Trim(ctxt�ʱ�.Text)
'            .�����Ա.סַ = Trim(ctxtסַ.Text)
'            .�����Ա.��� = Ccmb���.Text
'            .�����Ա.�绰���� = Trim(ctxt�绰.Text)
'            .�����Ա.���� = Trim(ctxt����.Text)
'            .�����Ա.������ = Trim(ctxt������.Text)
'            .�����Ա.������ = ctxt������.Text
'            .�����Ա.��ϵ�绰 = ctxt��ϵ�绰.Text
'            .�����Ա.�������� = ccmb��������.Text
'            .�����Ա.��ҵ��� = Ccmb��ҵ���.Text
'            .�����Ա.��λ��ַ = ctxt��λ��ַ.Text
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
'            .�����Ա.�Ļ��̶� = ccmb�Ļ��̶�.Text
'            .�����Ա.���� = ccmb����.Text
'            If ccmbUnit.Text = "" Then
'                .�����Ա.��λ������ = ""
'            Else
'                If .�����Ա.��λ������ <> mstr��λ������ Then
'                    '����λ������¸�ֵ���������»�ȡ���������ࡢ��ҵ���Ƭ����
'                    .�����Ա.��λ������ = mstr��λ������
'                End If
'            End If
            
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
       ' save�Ż��������Ŀ mcol�����Ŀ, pstr����ϵͳ���
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
'        ctxt����.SetFocus
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
