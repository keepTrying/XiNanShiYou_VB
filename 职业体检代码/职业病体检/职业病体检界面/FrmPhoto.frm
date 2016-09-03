VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{2327075A-FED2-4AEF-82B8-DD0C2B1AC8E1}#2.0#0"; "dyCatchPhoto.ocx"
Object = "{F48D4DEC-1198-11D5-91BE-0050BA06B70C}#5.9#0"; "录入控件.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmPhoto 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "职业健康体检-照相"
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
   StartUpPosition =   1  '所有者中心
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
      Begin VB.ComboBox ccmb体检类别 
         Height          =   300
         Left            =   5880
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   720
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.ComboBox ccmb体检类型 
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
      Begin VB.TextBox ctxt身份证号 
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
         Caption         =   "照像："
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
            Caption         =   "重新取像"
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
      Begin VB.ComboBox ccmb体检人类型 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2760
         TabIndex        =   27
         Top             =   1080
         Width           =   1695
      End
      Begin VB.ComboBox Ccmb体检人类别 
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
            Name            =   "宋体"
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
      Begin 录入控件.ctlInputDictGrid ctlInputDictGrid1 
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
            Name            =   "宋体"
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
         Caption         =   "体检类别："
         Height          =   180
         Index           =   3
         Left            =   4680
         TabIndex        =   57
         Top             =   840
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Label4 
         Caption         =   "体检人员类型："
         Height          =   255
         Left            =   2760
         TabIndex        =   56
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "体检表："
         Height          =   180
         Left            =   240
         TabIndex        =   55
         Top             =   1560
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "体检日期："
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
         Caption         =   "姓名："
         Height          =   180
         Left            =   2520
         TabIndex        =   53
         Top             =   2520
         Width           =   540
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别："
         Height          =   180
         Left            =   240
         TabIndex        =   52
         Top             =   3240
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄："
         Height          =   180
         Index           =   6
         Left            =   1320
         TabIndex        =   51
         Top             =   3240
         Width           =   540
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "身份证号："
         Height          =   180
         Left            =   240
         TabIndex        =   50
         Top             =   2520
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "系统编号："
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
         Caption         =   "注：刷条码前请确保文本框中内容为空"
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
         Caption         =   "非快速录入时黄色为必录项，快速录入时只需刷二代身份证"
         Height          =   180
         Left            =   240
         TabIndex        =   47
         Top             =   240
         Visible         =   0   'False
         Width           =   4680
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "出生日期："
         Height          =   180
         Left            =   2280
         TabIndex        =   46
         Top             =   3240
         Width           =   900
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "注：先刷条码，再刷身份证"
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
         Caption         =   "请将二代身份证放在读卡器上！"
         Height          =   180
         Left            =   240
         TabIndex        =   44
         Top             =   2280
         Width           =   2520
      End
      Begin VB.Label clblHintCheck 
         Caption         =   "注意：校核之后只允许照相，其它内容即使修改，也不会保存。"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3360
         TabIndex        =   43
         Top             =   480
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.Label clblHistory 
         Caption         =   "双击行，导入体检基本信息和附加信息。"
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
   Begin VB.CheckBox Check身份证 
      Caption         =   "刷二代身份证"
      Height          =   255
      Left            =   8520
      TabIndex        =   17
      Top             =   480
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CheckBox cchkClear 
      Caption         =   "保存后清空"
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
   Begin VB.Frame cfram基本信息 
      Caption         =   "登记基本信息（非快速录入时黄色为必录项，快速录入时只需照相):"
      ClipControls    =   0   'False
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   2520
      TabIndex        =   5
      Top             =   9360
      Width           =   6300
      Begin VB.TextBox ctxt民族 
         Height          =   300
         Left            =   4800
         TabIndex        =   21
         Top             =   2040
         Width           =   1695
      End
      Begin VB.ComboBox ccmb体检时期 
         Height          =   300
         Left            =   8160
         TabIndex        =   19
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox ctxt份数 
         Height          =   270
         Left            =   4440
         TabIndex        =   15
         Text            =   "1"
         Top             =   2880
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox ctxt体检单号 
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
      Begin 录入控件.ctlInputDictGrid c字典表 
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
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin 录入控件.ctlInputFrame ciptBase 
         Height          =   975
         Left            =   6120
         TabIndex        =   2
         Top             =   2280
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   1720
         BackColor       =   15791081
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         FormatString    =   "身份证号,1,0,12"
         Count           =   1
         titleInputBox0001=   "身份证号"
         statusinfoInputBox0001=   ""
         lengthInputBox0001=   12
         orderInputBox0001=   1
         valueInputBox0001=   ""
         datatypeInputBox0001=   3
         colInputBox0001 =   0
         rowInputBox0001 =   1
         PassWordCharInputBox0001=   0   'False
         主键InputBox0001=   0   'False
         允许等于最大值InputBox0001=   0   'False
         允许等于最小值InputBox0001=   0   'False
         字典名称InputBox0001=   ""
         显示字典字段InputBox0001=   ""
         保存字典字段InputBox0001=   ""
         名称InputBox0001=   "输入框 1"
         缺省值InputBox0001=   ""
         保存缺省值InputBox0001=   ""
         长度InputBox0001=   0
         MaxInputBox0001 =   ""
         MinInputBox0001 =   ""
         VisibleInputBox0001=   -1  'True
         PermitNullInputBox0001=   -1  'True
         TriggerstrInputBox0001=   ""
         允许多选InputBox0001=   0   'False
         ErrColor        =   12648447
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "请将二代身份证放在读卡器上！"
         Height          =   180
         Left            =   0
         TabIndex        =   22
         Top             =   0
         Width           =   2520
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "民族："
         Height          =   180
         Left            =   4800
         TabIndex        =   20
         Top             =   1800
         Width           =   540
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "体检时期："
         Height          =   180
         Left            =   8160
         TabIndex        =   18
         Top             =   480
         Width           =   900
      End
      Begin VB.Label clbl份数 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "份数："
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
         Caption         =   "体检单号："
         Height          =   180
         Index           =   7
         Left            =   4200
         TabIndex        =   12
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label clbl旧体检日期 
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
         Caption         =   "上次体检日期："
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
         Caption         =   "保存后请看状态栏"
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
         Caption         =   "试管编号："
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
         Name            =   "宋体"
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


'功能：记载当前窗体是否已加载，以便主导航界面判断当前窗体是否已执行过Form_Load。
Public Property Get pblnInUse() As Boolean
    pblnInUse = mblnInUse
End Property

Private Sub cchkClear_Click()
    On Error Resume Next
    ctxtName.SetFocus
End Sub

'Private Sub cchk录入单位名称_Click()
'    Dim lblnVisible As Boolean
'    On Error Resume Next
'    If cchk录入单位名称.Value = 1 Then
'        lblnVisible = True
'    Else
'        lblnVisible = False
'    End If
'    ccmbUnit.Visible = lblnVisible
'    ccmd单位定位.Visible = lblnVisible
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

Private Sub ccmb体检人类型_Click()
    Dim lobj体检类型 As Object
    On Error GoTo errHandler
    
    Set lobj体检类型 = CreateObject("职业病对象.clsmedicalexam")
    lobj体检类型.体检类型 = ccmb体检人类型.ItemData(ccmb体检人类型.ListIndex)
    
    '2012-06-14 于登淼 ↓
    '根据不同体检人员类型，生成不同的系统编号
    If Len(clblsysno.Text) = 0 Then
        clblsysno.Text = lobj体检类型.Func分配职业病体检系统编号 & (ccmb体检人类型.ListIndex + 1)
    Else
        clblsysno.Text = Left(clblsysno.Text, Len(clblsysno.Text) - 1) & (ccmb体检人类型.ListIndex + 1)
    End If
    mobj体检.系统编号 = Trim(clblsysno.Text)
    mobj体检.体检人员.系统编号 = Trim(clblsysno.Text)
    '2012-06-14 于登淼 ↑
    '2012-12-18 刘云乐  ↓
    'BUG号：0000092
'    If InStr(ccmb体检人类型.Text, "部队") > 0 Then
        Ccmb体检人类别.Text = "在岗期间"
'    Else
'        Ccmb体检人类别.ListIndex = 0
'    End If
    Call Ccmb体检人类别_Click
    '2012-12-18 刘云乐  ↑
    Exit Sub
errHandler:
    sfsub错误处理 "职业病界面", "frmregister", "Private Sub ccmb体检人类型_Click", Err.Number, Err.Description, True
End Sub

'Private Sub ccmb照射源_click()
'    ccmb职业类别.Visible = True
'    Call func职业类别
'    Exit Sub
'End Sub

'2012-07-11 于登淼
'判断是否重新照相
Private Sub ccmdTakePhotoAgain_Click()
    'Check身份证.Value = 1
    If Check身份证.Value = 1 Then
        Check身份证.Value = 0
    Else
        Check身份证_Click
    End If
    ccmdTakePhotoAgain.Visible = False
End Sub



'2012-07-16 于登淼
'双击某行载入当时的体检基本信息
Private Sub cgrdHistory_DblClick()
    Dim lstrSysNo As String
'    lstrSysNo = clblsysno.Text
    clblsysno.Text = cgrdHistory.TextMatrix(cgrdHistory.Row, 0)
    lstrSysNo = clblsysno.Text
    clblsysno_LostFocus
    clblsysno.Text = lstrSysNo
End Sub

'Private Sub ctxt出生地_Change()
'    If Len(Trim(ctxt出生地.Text)) > 50 Then
'        ctxt出生地.Text = Left(Trim(ctxt出生地.Text), 50)
'    End If
'End Sub

'Private Sub ctxt电话_Change()
'    If Len(Trim(ctxt电话.Text)) > 11 Then
'        ctxt电话.Text = Left(Trim(ctxt电话.Text), 11)
'    End If
'End Sub

'Private Sub ctxt工龄_Change()
'    If Len(Trim(ctxt工龄.Text)) > 2 Then
'        ctxt工龄.Text = Left(Trim(ctxt工龄.Text), 2)
'    End If
'End Sub

'Private Sub ctxt籍贯_Change()
'    If Len(Trim(ctxt籍贯.Text)) > 2 Then
'        ctxt籍贯.Text = Left(Trim(ctxt籍贯.Text), 2)
'    End If
'End Sub

'Private Sub ctxt联系电话_Change()
'    If Len(Trim(ctxt联系电话.Text)) > 11 Then
'        ctxt联系电话.Text = Left(Trim(ctxt联系电话.Text), 11)
'    End If
'End Sub

'Private Sub ctxt身份证号_KeyDown(KeyCode As Integer, Shift As Integer)
'    On Error Resume Next
'    If KeyCode = 13 Then
'        ctxtName.SetFocus
'        sub查看历史信息 (ctxt身份证号.Text)
'    End If
'End Sub
'Private Sub ccmbTemplate_KeyDown(KeyCode As Integer, Shift As Integer)
'    On Error Resume Next
'    If KeyCode = 13 Then
'        If ctxtTubeNo.Visible Then
'            ctxtTubeNo.SetFocus
'        Else
'            ctxt体检单号.SetFocus
'        End If
'    End If
'End Sub

'功能：控制不能输入体检表名称，只能选择。
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
'    '判断录入的单位是否在列表中存在，不存在则加入列表
'    i = gffuncItemIsInComboBox(ccmbUnit, ccmbUnit.Text)
'    If i = -1 Then
'        '加入到列表框中
'        ccmbUnit.AddItem ccmbUnit.Text
'
'        '加载到工作记忆簿文件中
'        pobj业务对象.当日工作记忆簿.sub增加单位名称 mstr单位申请编号 & "|" & ccmbUnit.Text
'    Else
'        '修改：2001-8-23。
'        On Error Resume Next
'        mstr单位申请编号 = pobj业务对象.当日工作记忆簿.单位编号(ccmbUnit.Text)
'        sub显示单位属性 ciptBase, mstr单位申请编号, mobjGUI
'    End If
'    Exit Sub
'errHandler:
'    sfsub错误处理 "职业病界面", "frmregister", "Sub ccmbUnit_Click", Err.Number, Err.Description, True
'
'End Sub

Private Sub ccmbUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        If ctxt份数.Visible Then
            ctxt份数.SetFocus
        Else
            If ciptBase.Visible Then
                ciptBase.SetFocus
            End If
        End If
    Else
        mstr单位申请编号 = ""
    End If
        
End Sub

Private Sub Ccmb体检人类别_Click()
    Dim lobj体检表模板集 As Object
    Dim lobj体检类别 As Object
    Dim lcolInfo As New Collection
    Dim lcol体检表编号 As Collection
    Dim i As Integer
    On Error GoTo errHandler
     '将体检类别加入组合框中
    'Set lobj体检类别 = CreateObject("职业病对象.clsmedicalexam")
    'lobj体检类别.体检类别 = ccmb体检人类别.ItemData(ccmb体检人类别.ListIndex)
    'lobj体检类别.体检表类别 = 1
    'Set lcol类别 = lobj体检类别.体检类别
    'ccmb体检人类别.AddItem ""
    'For i = 1 To lcol类别.recordCount
    '    ccmb体检人类别.AddItem lcol类别("类别")
    '    ccmb体检人类别.ItemData(ccmb体检人类别.NewIndex) = lcol类别("编号")
    '    lcol类别.movenext
    'Next
    'ccmb体检人类别.ListIndex = 0
    'Set lobj体检类别 = Nothing
   
    
    '将所有的非复查体检表模板加入到体检表下拉列表框中。再加体检类别条件
    ccmbTemplate.Clear
    Set lobj体检表模板集 = CreateObject("职业病对象.ClsMedicalExamTemplateSet")
    lobj体检表模板集.体检表类型 = Trim(ccmb体检人类型.Text)
    'lobj体检表模板集.体检表类别 = ccmb体检人类别.ItemData(ccmb体检人类别.ListIndex)
    lobj体检表模板集.体检表类别 = Trim(Ccmb体检人类别.Text)
    Set lcolInfo = lobj体检表模板集.元素集
    Set lcol体检表编号 = lobj体检表模板集.体检表编号元素集
    'ccmbTemplate.ListIndex = 0
    If lcolInfo.Count = 0 Then Exit Sub
    For i = 1 To lcolInfo.Count
        ccmbTemplate.AddItem lcolInfo(i)
        ccmbTemplate.ItemData(ccmbTemplate.NewIndex) = lcol体检表编号(i)
    Next
    ccmbTemplate.Text = ccmbTemplate.List(0)
    
    Set lobj体检表模板集 = Nothing
    Call ccmbTemplate_Click
    Exit Sub
errHandler:
    sfsub错误处理 "职业病界面", "frmregister", "ccmb体检人类别_click", Err.Number, Err.Description, True
End Sub

Private Sub ccmb体检人类型_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error Resume Next
    If KeyCode = 13 Then
        ctxt份数.SetFocus
    End If
End Sub

Private Sub cdtpDate_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ctxtName.SetFocus
    End If
End Sub

Private Sub Check身份证_Click()
    On Error GoTo errHandler
    '2012-06-20 于登淼 ↓
    '只有校核通过后，才会有照相的权限。
    '未通过时，只能刷身份证。
'    If Check身份证.Value = 0 Then   '照相
        Timer2.Enabled = False
'        If mintState = 1 Then  'mintstate=1表示通过校核，允许照相
            Picture1.Visible = False
            cctlCatchPhoto.Visible = True
            cctlCatchPhoto.funcInitVideo
            cctlCatchPhoto.Enabled = True
            mblnTakePhoto = True
            '2012-04-14 于登淼 ↓
            '当不刷二代身份证时，可以载入照片，也可以用摄像头拍照
'            ctbMain.Buttons(4).Enabled = True
            '2012-04-14 于登淼 ↑
'        End If
        
        '2012-07-11 于登淼 ↓
        '刷二代身份证，姓名，性别，年龄，出生日期，身份证号  enabled=false ,否则true
        '校核之后(mintstate=1)，不允许修改基本信息。
        '（界面上控制，只是为了查看时让人以为不能修改。实际上，只有此时拍照的照片可以保存。）
'        ctxt身份证号.Enabled = (mintState <> 1) 'True
'        ctxtName.Enabled = (mintState <> 1)  'True
'        ccmbSex.Enabled = (mintState <> 1)  'True
'        ctxtAge.Enabled = (mintState <> 1)  'True
'        cdtp出生.Enabled = (mintState <> 1)  'True
        '2012-07-11 于登淼 ↑
        
        Label31.Visible = False
'    Else                            '刷身份证
'        Picture1.Visible = True
'        cctlCatchPhoto.Visible = False
'        ctxt身份证号.Enabled = False
'        ctxtName.Enabled = False
'        ccmbSex.Enabled = False
'        ctxtAge.Enabled = False
'        cdtp出生.Enabled = False
'        Label31.Visible = True
'        If mblnTakePhoto Then
'            cctlCatchPhoto.subDisconnect
'            mblnTakePhoto = False
'        End If
        '2012-04-14 于登淼 ↓
        '当刷二代身份证时，不能载入照片
'        ctbMain.Buttons(4).Enabled = False
        '2012-04-14 于登淼 ↑
'    End If
    
    Exit Sub
errHandler:
    sfsub错误处理 "职业病界面", "frmregister", "Sub Check身份证_Click", Err.Number, Err.Description, True
End Sub

'此功能暂不用
'Private Sub ciptBase_LastLostFocus()
'    Dim blnCancel As Boolean
'    On Error Resume Next
    '自动保存。
 '   If ctbMain.Buttons(6).Enabled Then
 '       ctxtName.SetFocus
 '       SendKeys "{F2}"
 '   End If
'End Sub

'Private Sub ciptBase_LostFocus()
'    On Error Resume Next
'    If ActiveControl.Name <> "c字典表" Then
 '       c字典表.Visible = False
 '   End If
'End Sub


Private Sub clblsysno_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error Resume Next
    If KeyCode = 13 Then
        '重新初始化照相控件。
         If cctlCatchPhoto.Status = "恢复" Then
            cctlCatchPhoto.sub转换状态
            cctlCatchPhoto.subClear
         End If
        '根据系统编号查找信息
        subLoad clblsysno
'        Timer2.Enabled = True
    End If
End Sub
''根据系统编号查找编号信息
Private Sub subLoad(ByVal para系统编号 As String)
    Dim lobjRec As Object
     On Error GoTo errHandler
        Set lobjRec = dafuncGetData("select 体检类型,体检类别,体检表编号,建档日期,公民身份号码,姓名,性别,年龄,出生日期 from 职业病体检_体检基本数据库 where 系统编号='" & para系统编号 & "'")
        If Not (lobjRec.BOF Or lobjRec.EOF) Then
            mobj体检.系统编号 = Trim(clblsysno.Text)
            ccmb体检人类型.Text = IIf(IsNull(lobjRec!体检类型), "", lobjRec!体检类型)
            Ccmb体检人类别.Text = IIf(IsNull(lobjRec!体检类别), "", lobjRec!体检类别)
            ccmbTemplate.Text = IIf(IsNull(lobjRec!体检表编号), "", lobjRec!体检表编号)
            '2015-10-16
            'cdtpDate.Value = IIf(IsNull(lobjRec!建档日期), "", lobjRec!建档日期)
            cdtpDate.Value = Now
            ctxt身份证号 = IIf(IsNull(lobjRec!公民身份号码), "", lobjRec!公民身份号码)
            ctxtName = IIf(IsNull(lobjRec!姓名), "", lobjRec!姓名)
            ccmbSex = IIf(IsNull(lobjRec!性别), "", lobjRec!性别)
            ctxtAge = IIf(IsNull(lobjRec!年龄), "", lobjRec!年龄)
            cdtp出生 = IIf(IsNull(lobjRec!出生日期), "", lobjRec!出生日期)
        Else
            MsgBox "无效条码！", vbExclamation, "系统提示！"
            subClear
        End If
     Exit Sub
errHandler:
    sfsub错误处理 "Frmphoto", "frmregister", "Sub subLoad", Err.Number, Err.Description, True
End Sub
Private Sub clblsysno_LostFocus()
    Dim lobjRec As Object
    Dim strSQL As String
    'Dim lobj系统编号 As Object
    Dim strTmp As String
    Dim str单位申请编号 As String
    '2012-06-13 于登淼 ↓
    '获取身份证照片变量，获取现场照片变量
    Dim lobj身份证照片 As Object
    Dim lobj现场照片 As Object
    '2012-06-13 于登淼 ↑
    
    On Error GoTo errHandler
    strTmp = Trim(clblsysno.Text)
    
    '2012-07-11 于登淼 ↓
    '系统编号固定了，就不能再更改了。
'    clblSysNo.Enabled = False
    '2012-07-11 于登淼 ↑
    
    '2012-06-14 于登淼 ↓
    '因为存在先录入信息再生成条码号的情况，故这个判断取消
    If Len(clblsysno.Text) = 0 Then
        'MsgBox "系统编号错误，请检查！", vbInformation, "系统提示"
        Exit Sub
    End If
    '2012-06-14 于登淼 ↑
    
    Set lobjRec = dafuncGetData("select * from 职业病体检_体检基本数据库 where 系统编号='" & strTmp & "'")
    If lobjRec.RecordCount = 0 Then
        mobj体检.系统编号 = Trim(clblsysno.Text)
        'mobj体检. = Trim(clblSysNo.Text)
    ElseIf lobjRec.RecordCount = 1 Then
        '2012-07-11 于登淼 ↓
        '导入信息时，可能遇到文件导入时缺少部分信息。此时，忽略错误继续下一行代码。
        On Error Resume Next
        '2012-07-11 于登淼 ↑
        mobj体检.系统编号 = Trim(clblsysno.Text)
        ccmb体检人类型.Text = IIf(IsNull(lobjRec!体检类型), "", lobjRec!体检类型)
        Ccmb体检人类别.Text = IIf(IsNull(lobjRec!体检类别), "", lobjRec!体检类别)
        ccmbTemplate.Text = IIf(IsNull(lobjRec!体检表编号), "", lobjRec!体检表编号)
        cdtpDate.Value = IIf(IsNull(lobjRec!建档日期), "", lobjRec!建档日期)
        ctxt身份证号 = IIf(IsNull(lobjRec!公民身份号码), "", lobjRec!公民身份号码)
        ctxtName = IIf(IsNull(lobjRec!姓名), "", lobjRec!姓名)
        ccmbSex = IIf(IsNull(lobjRec!性别), "", lobjRec!性别)
        ctxtAge = IIf(IsNull(lobjRec!年龄), "", lobjRec!年龄)
        cdtp出生 = IIf(IsNull(lobjRec!出生日期), "", lobjRec!出生日期)
'        ctxt籍贯 = IIf(IsNull(lobjRec!籍贯), "", lobjRec!籍贯)
'        ctxt邮编 = IIf(IsNull(lobjRec!邮编), "", lobjRec!邮编)
'        ctxt住址 = IIf(IsNull(lobjRec!住址), "", lobjRec!住址)
'        ccmb文化程度 = IIf(IsNull(lobjRec!文化程度), "", lobjRec!文化程度)
'        Ccmb婚否 = IIf(IsNull(lobjRec!婚否), "", lobjRec!婚否)
'        ccmb民族 = IIf(IsNull(lobjRec!民族), "", lobjRec!民族)
'        ctxt电话 = IIf(IsNull(lobjRec!电话号码), "", lobjRec!电话号码)
'        ctxt工龄 = IIf(IsNull(lobjRec!工龄), "", lobjRec!工龄)
'        ctxt出生地 = IIf(IsNull(lobjRec!出生地), "", lobjRec!出生地)
'        ccmb照射源 = IIf(IsNull(lobjRec!照射源), "", lobjRec!照射源)
'        ccmb职业类别 = IIf(IsNull(lobjRec!职业分类), "", lobjRec!职业分类)
'        ccmb危害因素 = IIf(IsNull(lobjRec!危害因素), "", lobjRec!危害因素)
'        ccmb现工种 = IIf(IsNull(lobjRec!现工种), "", lobjRec!现工种)
'        ccmb职务 = IIf(IsNull(lobjRec!职务或职称), "", lobjRec!职务或职称)
'        ctxt危害工龄 = IIf(IsNull(lobjRec!职业危害工龄), "", lobjRec!职业危害工龄)
'        ctxt放射剂量 = IIf(IsNull(lobjRec!放射剂量), "", lobjRec!放射剂量)
        str单位申请编号 = IIf(IsNull(lobjRec!单位申请编号), "", lobjRec!单位申请编号)
        
        '获取像片
'        Set lobjRec = CreateObject("职业病对象.clspersonexamed")
'        lobjRec.系统编号 = Trim(clblsysno.Text)
'        If lobjRec.像片 <> 0 Then
'            Picture1.Picture = lobjRec.像片
'            Picture1.Visible = True
'            cctlCatchPhoto.Visible = False
'            If mblnTakePhoto Then
'                cctlCatchPhoto.subDisconnect
'                mblnTakePhoto = False
'            End If
'        End If
        
'''        '2012-07-11 于登淼 ↓
'''        '单独获取现场照片
'''        Set lobj现场照片 = lobjRec.func获取现场照片(Trim(clblsysno.Text), "职业病体检")
'''        If Not lobj现场照片 Is Nothing Then Picture1.Picture = lobj现场照片
'''        Picture1.Visible = True
'''        '2012-07-11 于登淼 ↑
        
        
        '2012-06-13 于登淼 ↓
        '获取该体检人员身份证照片
'        Set lobj身份证照片 = lobjRec.func查找身份证照片(Trim(clblsysno.Text) & "IDcard", "职业病体检")
'        If Not lobj身份证照片 Is Nothing Then
'            Picture2.Picture = lobj身份证照片
'            If mintState = 1 And mblnTakePhoto = False Then ccmdTakePhotoAgain.Visible = True
'        End If
'        Picture2.Visible = True
        '2012-06-13 于登淼 ↑
        
        If FrmRegister.pstr复查系统编号 <> "" Then
            ccmdTakePhotoAgain.Visible = False
            Me.ctbMain.Buttons(7).Visible = False
        End If
        
        On Error GoTo errHandler
        If Not IsNull(str单位申请编号) Then
            func获取单位信息 str单位申请编号
        End If
    Else
        MsgBox "系统编号不唯一，请检查！", vbInformation, "系统提示"
        Exit Sub
    End If
    
    Set lobjRec = Nothing
    Exit Sub
errHandler:
    sfsub错误处理 "职业病界面", "frmregister", "Sub clblsysno_LostFocus", Err.Number, Err.Description, True
End Sub

Private Sub ctxtAge_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        ctxt体检单号.SetFocus
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
'不用
Private Sub ctxt份数_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        '若录入板没有录入项目，则直接保存。
        If ciptBase.Visible Then
            ciptBase.SetFocus
            ciptBase.ItemSetFocus 0
        End If
    End If
End Sub
'录完身份证号后，获取年龄，性别，出生日期
'Private Sub ctxt身份证号_lostfocus()
'    Dim ldatBirth As String
'    Dim lstrSex As String
'    On Error GoTo errHandler
'    If Trim(ctxt身份证号.Text) <> "" Then
'            '正确时从身份证号中获取出生日期。
'            sub根据公民身份号码获取生日和性别 ctxt身份证号.Text, ldatBirth, lstrSex
'            If Not IsDate(ldatBirth) Then
'                MsgBox ("身份证号不合法！")
'                Exit Sub
'            End If
'
'            '查找是否需要录入出生日期，需要时自动根据身份证号填写出生日期
'            On Error Resume Next
'            If IsDate(ldatBirth) Then
'                cdtp出生.Value = ldatBirth
''                ctxtAge.Text = DateDiff("yyyy", ldatBirth, Date)
'                ctxtAge.Text = Year(Date) - Year(ldatBirth)
''罗李奎 2012-12-11 ↓
''说明：岁数是判断身份证上日期是否过了当前日期，过了则加一岁。
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
'    '罗李奎 2012-12-11 ↑
'            End If
'            ccmbSex.Text = lstrSex
'    End If
'    Exit Sub
'errHandler:
'    sfsub错误处理 "职业病界面", "frmregister", "Sub ctxt身份证号_lostfocus", Err.Number, Err.Description, True
'End Sub

'Private Sub ctxt体检单号_KeyDown(KeyCode As Integer, Shift As Integer)
'    On Error Resume Next
'    If KeyCode = 13 Then
'        If ccmbUnit.Visible Then
'            ccmbUnit.SetFocus
'        Else
'            If ctxt份数.Visible Then
'                ctxt份数.SetFocus
'            Else
'                If ciptBase.Visible Then
'                    ciptBase.SetFocus
'                End If
'            End If
'        End If
'    End If
'End Sub
'
'Private Sub ctxt邮编_Change()
'    If Len(Trim(ctxt邮编.Text)) > 6 Then
'        ctxt邮编.Text = Left(Trim(ctxt邮编.Text), 6)
'    End If
'End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If mblnTakePhoto Then
        '重新初始化照相控件。
        cctlCatchPhoto.funcInitVideo
    End If
    'ctxtName.SetFocus
End Sub

'Private Sub Form_Deactivate()
'    On Error Resume Next
'    gfsubHideComboList ccmbUnit
'End Sub
'初始化界面

Private Sub Form_Load()
    On Error GoTo errHandler
    
    '如果窗体已经初始化过则不再进行初始化。
    If mblnInUse Then Exit Sub
    
    '设置窗体正在使用的标志。
    mblnInUse = True
    
    MousePointer = 11
    
    '界面不可操作。
'    cfram基本信息.Enabled = False
    'ctbMain.Enabled = False
    'clblsysno.Visible = False
    Set mcol收费项目 = New Collection
    Set mcol体检项目 = New Collection
    
    Set mobj旧体检 = CreateObject("职业病对象.clsMedicalExam")
    
    Set mobj体检 = CreateObject("职业病对象.clsMedicalExam")
    '修改：2002-10-10（设置系统编号名称）。
    If pstr系统编号名称 <> "" Then
        mobj体检.系统编号名称 = pstr系统编号名称
    End If
    
    Set mobj体检集 = CreateObject("职业病对象.clsMedicalExamSet")
    Set mobj体检表模板 = CreateObject("职业病对象.ClsMedicalExamTemplate")
    
    '初始化界面通用对象（每个界面对应一个界面通用对象实例，不可混用，切记切记）。
    Set mobjGUI = New cls界面通用对象
    mobjGUI.pbln自动设置字典高度 = False
    
    '设置工具栏上所需要的各种按钮。
    Dim lcol工具栏按钮 As New Collection           '工具栏上的按钮初始化集合。
    With lcol工具栏按钮
        .Add "保存"
        .Add "|"
        .Add "退出"
    End With
    
    '为需要通过界面通用对象控制的各种控件赋初始值。
    With mobjGUI
        Set .Form = Me
        Set .c工具栏 = ctbMain
        Set .c录入板 = ciptBase
        Set .c字典表 = c字典表
        Set .c状态栏 = cstbMain
        
        '调用界面通用对象提供的方法，对界面控件进行初始化。
        .subInitialize lcol工具栏按钮, ""
    End With
    
    If 访问标志 = 2 Then
        ccmb体检人类型.Enabled = False
        Ccmb体检人类别.Enabled = False
        ccmbTemplate.Enabled = False
        ctbMain.Buttons(1).Enabled = False
    End If

    '清空
    subClear
    cdtpDate.Value = Now
    
    '新分配系统编号
    'clblsysno.Caption = mobj体检.Func分配系统编号
    'clblSysNo.Text = ""
    mobj体检.系统编号 = Trim(clblsysno.Text)
    'pstr系统编号 = Trim(clblSysNo.Text)
'    If Check身份证.Value = 0 Then
'        Check身份证.Value = 1
'    Else
'        Check身份证_Click
'    End If
    'cctlCatchPhoto.Visible = False
    'cctlCatchPhoto.Visible = True
    
'    If pobj业务对象.业务设置("试管编号自动生成") = "否" Then
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
    ccmb体检人类型.Visible = True
    Label2(3).Visible = True
    clbl旧体检日期.Visible = True
'    clblSysNo.Enabled = True

    '为了加快窗体加载速度，余下初始化工作放在定时器中完成。
'    Timer1.Enabled = True
'    Timer2.Enabled = True
    subLoad pstrPhoto
        '校核
'    mintState = 1
'    Check身份证.Value = 1
    Check身份证_Click
'    mintState = 1
    pobj业务对象.func写入单人当前体检状态 clblsysno, mintState
    pobj业务对象.func写入校核人信息 clblsysno, um用户编号
    MousePointer = 0
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病界面部件", "FrmRegisterAnnual", "Form_Load", 6666, lstrError, False
    '恢复工具栏可用。
    ctbMain.Enabled = True
    MousePointer = 0
    cstbMain.Panels(1) = lstrError
End Sub



'功能：完成form_load余下的初始化工作。
Private Sub Timer1_Timer()
    Dim lobj体检表模板集 As Object  '体检表模板集，获取所有的非复查体检表模板名称。
    Dim lcolInfo As Collection
    Dim lcol类别 As Object
    Dim lcol类型 As Object
    Dim i As Integer
    Dim lobj体检类别 As Object
    Dim lobj体检类型 As Object
    Dim lobjRec As Object
    On Error GoTo errHandler
    
    '定时器不再起作用。
    Timer1.Enabled = False
    
    '从当日工作及已簿中获取当天录入过的单位名称。
    Set lcolInfo = pobj业务对象.当日工作记忆簿.单位名称集
    For i = 1 To lcolInfo.Count
'        ccmbUnit.AddItem lcolInfo(i)
    Next
    
    '将体检类别加入组合框中
    Set lobjRec = pobjDict.FetchEx("体检类型字典")
    Ccmb体检人类别.Clear
    'Ccmb体检人类别.AddItem ""
    For i = 1 To lobjRec.RecordCount
        Ccmb体检人类别.AddItem lobjRec("名称")
        Ccmb体检人类别.ItemData(Ccmb体检人类别.NewIndex) = lobjRec("编号")
        lobjRec.MoveNext
    Next
'    Ccmb体检人类别.ListIndex = 0
    '2012-06-15 于登淼 ↓
    '修改体检人员信息，弹出窗体时，防止体检编号出错。
    If clblsysno.Text = "" Then
        Ccmb体检人类别.ListIndex = 0
    Else
        If Right(clblsysno.Text, 1) < "0" Or Right(clblsysno.Text, 1) > "9" Then
            Ccmb体检人类别.ListIndex = CInt(Left(Right(clblsysno.Text, 2), 1) - 1)
        Else
            Ccmb体检人类别.ListIndex = CInt(Right(clblsysno.Text, 1) - 1)
'            Ccmb体检人类别.ListIndex = 0
        End If
    End If
    '2012-06-15 于登淼 ↑
   
    Set lobjRec = pobjDict.FetchEx("体检人类别字典")
    ccmb体检人类型.Clear
    'ccmb体检人类型.AddItem ""
    For i = 1 To lobjRec.RecordCount
        ccmb体检人类型.AddItem lobjRec("名称")
        ccmb体检人类型.ItemData(ccmb体检人类型.NewIndex) = lobjRec("编号")
        lobjRec.MoveNext
    Next
    '2012-06-15 于登淼 ↓
    '修改体检人员信息，弹出窗体时，防止体检编号出错。
    If clblsysno.Text = "" Then
        ccmb体检人类型.ListIndex = 0
    End If
    '2012-06-15 于登淼 ↑
    
    '将所有的非复查体检表模板加入到体检表下拉列表框中。
    'Set lobj体检表模板集 = CreateObject("职业病对象.ClsMedicalExamTemplateSet")
    'lobj体检表模板集.体检表类型 = 3
    
    'lobj体检表模板集.体检表类别 = ccmb体检人类别.ItemData(ccmb体检人类别.ListIndex)
    'Set lcolInfo = lobj体检表模板集.元素集
    'For i = 1 To lcolInfo.Count
    '    ccmbTemplate.AddItem lcolInfo(i)
    'Next
    'ccmbTemplate.Text = ccmbTemplate.List(0)
    'Set lobj体检表模板集 = Nothing
    
    '2012-06-15 于登淼 ↓
    '该功能在该函数后面已有，故注释掉
'''    '根据业务设置判断是否照相,判断时加上界面上的是否刷二代身份证。
'''    If pobj业务对象.业务设置("是否照像") = "是" Then
'''        mblnTakePhoto = True
'''    Else
'''        mblnTakePhoto = False
'''    End If
    '2012-06-15 于登淼 ↑
    
    If pobj业务对象.业务设置("是否快速登记") = "是" Then
        mbln快速录入 = True
    Else
        mbln快速录入 = False
    End If
    
    '只有初检，而且快速登记才可以批量登记。
    If Not mbln快速录入 Or pstr系统编号 <> "" Then
        clbl份数.Visible = False
        ctxt份数.Visible = False
    End If
    
    'ccmb体检人类型.ListIndex = 0
    
    If ccmbTemplate.ListCount > 0 Then
        'ccmbTemplate.ListIndex = 0
'        ccmbTemplate.Text = ccmbTemplate.List(0)
'        subChangeTemplate
        
    End If
    
'    If pstr系统编号 <> "" Then
'        '是年检登记。
'        '显示体检人员基本信息。
'        SubGetPersonInfo pstr系统编号
'    End If
    
    On Error Resume Next
    Set mobj记忆 = New cls用户操作记忆
    mobj记忆.用户编号 = "*"
    mobj记忆.业务名 = "体检管理"
    mstr默认年龄 = mobj记忆.记忆项值("体检年龄")
'    If mstr默认年龄 <> "" And ctxtAge = "" Then
'        ctxtAge = mstr默认年龄
'    End If
    
    If mobj记忆.记忆项值("体检登记时录入单位名称") = "" Or mobj记忆.记忆项值("体检登记时录入单位名称") = "是" Then
'        cchk录入单位名称.Value = 1
    Else
'        cchk录入单位名称.Value = 0
    End If
    cfram基本信息.Enabled = True
    ctbMain.Enabled = True
    
    '2012-06-15 于登淼 ↓
    '省疾控新要求，初始登记只刷身份证，校核通过后，再照相。
'''    '需要照相时初始化照相控件。
'''    If mblnTakePhoto And Check身份证 = False Then
'''        '初始化控件。
'''        cctlCatchPhoto.funcInitVideo
'''        '照相控件先visible=false再visible=true，刷新一次，这样 “取像” 按钮才能正常显示
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
    Set lobjRec = CreateObject("职业病对象.clsManageMedicalExam")
    lstrTmp = lobjRec.func获取单人当前体检状态(Trim(clblsysno.Text))
'    If lstrTmp = "未校核" Or lstrTmp = "" Or (访问标志 = 1 And lstrTmp <> "未打清单") Then   '只有当登记未校核时，才刷身份证
'        If Check身份证.Value = 0 Then
'            Check身份证.Value = 1
'        Else
'            Check身份证_Click
'        End If
'        mintState = 0
''        clblHintCheck.Visible = False
'        sub读卡器初始化
'        func身份证认证
'    ElseIf lstrTmp = "未打清单" Then
'        mintState = 1
'        clblHintCheck.Visible = True
'        'If 访问标志 = 0 Then Check身份证.Value = 0:Check身份证_Click
'        If Check身份证.Value = 1 Then
'            Check身份证.Value = 0
'        Else
'            Check身份证_Click
'        End If
'    ElseIf lstrTmp = "待复查" Then  '待复查项暂且当做未打清单类初始化照相与刷身份证功能。
If lstrTmp = "待复查" Then  '待复查项暂且当做未打清单类初始化照相与刷身份证功能
        mintState = 1
        clblHintCheck.Visible = True
        If Check身份证.Value = 1 Then
            Check身份证.Value = 0
        Else
            Check身份证_Click
        End If
    End If
    '2012-06-15 于登淼 ↑
    
    MousePointer = 0

    '获取文化程度
    Set lobjRec = pobjDict.FetchEx("文化程度字典")
'    ccmb文化程度.Clear
'    ccmb文化程度.AddItem ""
'    For i = 1 To lobjRec.RecordCount
'        ccmb文化程度.AddItem lobjRec("名称")
'        ccmb文化程度.ItemData(ccmb文化程度.NewIndex) = lobjRec("编号")
'        lobjRec.MoveNext
'    Next
'    ccmb文化程度.ListIndex = 0
'
'    '获取婚否
'    Set lobjRec = pobjDict.FetchEx("婚姻字典")
'    Ccmb婚否.Clear
'    Ccmb婚否.AddItem ""
'    For i = 1 To lobjRec.RecordCount
'        Ccmb婚否.AddItem lobjRec("名称")
'        Ccmb婚否.ItemData(Ccmb婚否.NewIndex) = lobjRec("编号")
'        lobjRec.MoveNext
'    Next
'    Ccmb婚否.ListIndex = 0
    
'     '获取民族
'    Set lobjRec = pobjDict.FetchEx("民族字典")
'    ccmb民族.Clear
'    ccmb民族.AddItem ""
'    For i = 1 To lobjRec.RecordCount
'        ccmb民族.AddItem lobjRec("名称")
'        ccmb民族.ItemData(ccmb民族.NewIndex) = lobjRec("编号")
'        lobjRec.MoveNext
'    Next
'    ccmb民族.ListIndex = 0
'
'     '获取经济性质
'    Set lobjRec = pobjDict.FetchEx("经济性质字典")
'    ccmb经济性质.Clear
'    ccmb经济性质.AddItem ""
'    For i = 1 To lobjRec.RecordCount
'        ccmb经济性质.AddItem lobjRec("名称")
'        ccmb经济性质.ItemData(ccmb经济性质.NewIndex) = lobjRec("编号")
'        lobjRec.MoveNext
'    Next
'    ccmb经济性质.ListIndex = 0
'
'    '获取危害种类
'    Set lobjRec = pobjDict.FetchEx("危害种类字典")
'    ccmb危害因素.Clear
'    ccmb危害因素.AddItem ""
'    For i = 1 To lobjRec.RecordCount
'        ccmb危害因素.AddItem lobjRec("名称")
'       ccmb危害因素.ItemData(ccmb危害因素.NewIndex) = lobjRec("编号")
'        lobjRec.MoveNext
'    Next
'    ccmb危害因素.ListIndex = 0
'
'    '获取职业或职称
'    Set lobjRec = pobjDict.FetchEx("职业或职称字典")
'    ccmb职务.Clear
'    ccmb职务.AddItem ""
'    For i = 1 To lobjRec.RecordCount
'        ccmb职务.AddItem lobjRec("名称")
'        ccmb职务.ItemData(ccmb职务.NewIndex) = lobjRec("编号")
'        lobjRec.MoveNext
'    Next
'    ccmb职务.ListIndex = 0
'
'    '获取工种字典
'    Set lobjRec = pobjDict.FetchEx("工种字典")
'    ccmb现工种.Clear
'    ccmb现工种.AddItem ""
'    For i = 1 To lobjRec.RecordCount
'        ccmb现工种.AddItem lobjRec("名称")
'        ccmb现工种.ItemData(ccmb现工种.NewIndex) = lobjRec("编号")
'        lobjRec.MoveNext
'    Next
'    ccmb现工种.ListIndex = 0
'    Call func获取行业类别字典
'    Call func照射源
'    clblSysNo.Visible = True
'    clblSysNo.SetFocus
     '2012-07-11 于登淼 ↓
    '系统编号生成之后，不可以改变。转移focus之前，必须先设定setfocus，这样保证个人信息可以加载进去。
    '且加载完后，控件enabled=false
    'If 访问标志 = 1 Then
'        ctxt籍贯.SetFocus
    'End If
    '访问标志 = 0       '不知何时出现的代码。现在必须用这个变量控制登记的部分保存过程。
    '2012-07-11 于登淼 ↑
    
    '2012-08-18 于登淼 ↓
    '载入个人基本信息
    If pstr复查系统编号 <> "" Then
        ctbMain.Buttons(4).Enabled = False
    End If
    
    Form_Activate
    clblsysno_LostFocus
    
    '2012-08-18 于登淼 ↑
    
'''    '2012-08-19 于登淼 ↓
'''    '复检时更改体检日期为今日，否则为初检建档日期
'''    If 访问标志 = 2 And pstr复查系统编号 <> "" Then cdtpDate.Value = Now
'''    '2012-08-19 于登淼 ↑
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病界面部件", "FrmRegisterAnnual", "Timer1_Timer", 6666, lstrError, False
    '恢复界面可操作。
    cfram基本信息.Enabled = True
    ctbMain.Enabled = True
    MousePointer = 0
End Sub

Private Sub ccmbTemplate_Click()
    On Error GoTo errHandler
    '选择体检表
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
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病界面部件", "FrmRegisterAnnual", "ccmbTemplate_Click", 6666, lstrError, False
    
    Exit Sub
    Resume
End Sub

Private Sub subChangeTemplate()
    On Error GoTo errHandler
    '选择体检表
    Dim lcolInfo As Collection
    Dim lstrTubeNo As String
    Dim lstrTemp As String
    Dim i As Integer, j As Integer
    
    '获取新试管编号。
    If mobj体检.体检表.体检表名 <> ccmbTemplate.Text Then
'        mobj体检.体检表.体检表名 = ccmbTemplate.Text

        '根据体检表模板获取该体检表所有可用的字母。
        mobj体检表模板.体检表名 = ccmbTemplate.Text

'        If pobj业务对象.业务设置("试管编号自动生成") = "是" Then
'            '试管编号字母为空时cvscLetter可用
'            If mobj体检.体检表.试管编号字母 = "" Then
'                '将字母按逗号分开，加入mcoltubeNo中
'                lstrTubeNo = mobj体检表模板.试管字母编号
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
'                    '试管字母改变了，给出提示。
'                    If clblLetter.Caption <> "" And clblLetter.Caption <> mcolTubeNo(1) Then
'                        sffuncMsg "请注意，你现在选择的体检表使用的试管字母与前一个（" & clblLetter.Caption & "）不同了。"
'                    End If
'
'                    '赋值给clblLetter
'                    clblLetter.Caption = mcolTubeNo(1)
'                    cvscLetter.Enabled = True
'                    cvscLetter.Min = 1
'                    cvscLetter.Max = mcolTubeNo.Count
'                    cvscLetter.Value = 1
'                Else
'                    ctbMain.Buttons(6).Enabled = False
'                    '提示该体检表无可用的字母。
'                    Err.Raise 6666, , "该体检表无可用试管字母编号，请先设置体检表对应的试管字母编号"
'                End If
'            Else
                '有字母，不能选择字母。
'                clblLetter.Caption = mobj体检.体检表.试管编号字母
'                cvscLetter.Enabled = False
'            End If
'        Else
'            clblLetter.Caption = mobj体检表模板.试管字母编号
'        End If
        
        '初始化附加信息。
        On Error Resume Next
        mobjGUI.sub初始化录入板 ccmbTemplate.Text
        
        '修改：2001-8-23（显示单位属性）。
        If mstr单位申请编号 <> "" Then
            sub显示单位属性 ciptBase, mstr单位申请编号, mobjGUI
        End If

        '重新填写附加信息值。
'        If mobj体检表模板.基本附加项目集.Count > 0 Then
'            Set lcolInfo = mobj旧体检.体检表.附加信息
'            If lcolInfo.Count > 0 Then
'                sub填录入板值 ciptBase, mobjGUI, lcolInfo
'            End If
'        End If

        '修改：2002-7-26（杨春）根据“是否年检表”选择体检表类型。
        'If mobj体检表模板.是否年检表 Then
        '    ccmb体检人类型.ListIndex = 1
        'Else
        '    ccmb体检人类型.ListIndex = 0
        'End If

        '修改：2002-10-10（杨春）嘉定定制：显示体检金额。
        On Error Resume Next
        ciptBase.Box1("体检金额").Text = mobj体检表模板.收费标准金额
    End If
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病界面部件", "FrmRegisterAnnual", "subChangeTemplate", 6666, lstrError, True
    
    Exit Sub
    Resume
End Sub

'自动弹出列表框
'Private Sub ccmbUnit_GotFocus()
'    On Error GoTo errHandler
''    gfsubShowComboList ccmbUnit
'    Exit Sub
'errHandler:
'    'sfsub错误处理 "职业病界面部件", "FrmRegisterAnnual", "ccmbUnit_GotFocus", Err.Number, Err.Description, False
'End Sub

'Private Sub ccmbUnit_LostFocus()
'    On Error GoTo errHandler
'    Dim i As Integer
'    If Trim(ccmbUnit.Text) = "" Then Exit Sub
'
'    '判断录入的单位是否在列表中存在，不存在则加入列表
'    i = gffuncItemIsInComboBox(ccmbUnit, ccmbUnit.Text)
'    If i = -1 Then
'        '加入到列表框中
'        ccmbUnit.AddItem ccmbUnit.Text
'
'        '加载到工作记忆簿文件中
'        pobj业务对象.当日工作记忆簿.sub增加单位名称 mstr单位申请编号 & "|" & ccmbUnit.Text
'    Else
'        '修改：2001-8-26（若单位申请编号不同，修改工作记忆簿）。
'        If mstr单位申请编号 <> pobj业务对象.当日工作记忆簿.单位编号(ccmbUnit.Text) And mstr单位申请编号 <> "" Then
'            pobj业务对象.当日工作记忆簿.sub增加单位名称 mstr单位申请编号 & "|" & ccmbUnit.Text
'        End If
'    End If
'    Exit Sub
'errHandler:
'    sfsub错误处理 "职业病界面", "frmregister", "Sub ccmbUnit_LostFocus", Err.Number, Err.Description, True
'End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    Select Case KeyCode
    Case vbKeyF8
        If mblnTakePhoto Then
            If cctlCatchPhoto.VideoIsOk Then
                cctlCatchPhoto.sub转换状态
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

'功能：窗体初始化。

'调用单位定位
'Private Sub ccmd单位定位_Click()
'    On Error GoTo errHandler
'    Dim lobjRec As Object  '单位定位返回的结果记录。
'    Dim lobj单位 As Object
'    Dim lobj单位信息 As Object
'    '启动单位定位界面。
'    Set lobjRec = pobj业务对象.func单位定位
'    '获取定位的单位，显示在“单位名称”录入框中。
'    If Not lobjRec Is Nothing Then
'        If lobjRec.RecordCount > 0 Then
'            ccmbUnit.Text = IIf(IsNull(lobjRec("单位名称")), "", lobjRec("单位名称"))
'            mstr单位申请编号 = lobjRec!申请编号
'            'Set lobj单位 = CreateObject("职业病对象.class1")
'            'lobj单位.单位信息申请 = lobjRec!申请编号
'            'Set lobj单位信息申请 = lobj单位.单位信息
'
'
'
'            If mstr单位申请编号 <> "" Then
'                '修改：2001-8-23（显示单位属性）。
'                On Error Resume Next
'                'sub显示单位属性 ciptBase, mstr单位申请编号, mobjGUI
'                func获取单位信息 lobjRec!申请编号
'            End If
'        End If
'    End If
'
'    '把焦点回到单位录入框。保存能保存新单位定位信息。
'    ccmbUnit.SetFocus
'    SendKeys vbTab
'    Exit Sub
'errHandler:
'    Dim lstrError As String
'    lstrError = func错误处理(Err.Number, Err.Description)
'    sfsub错误处理 "职业病界面部件", "FrmRegisterAnnual", "ccmd单位定位_Click", 6666, lstrError, False
'End Sub

Private Sub mobjGUI_ItemLostFocus(ByVal Index As Integer, ByVal 名称 As String, ByVal 内容 As String, ByVal 保存内容 As String, ByVal IsError As Boolean)
    On Error GoTo errHandler
    Dim lstrIDCard As String
    Dim i As Integer
    Dim ldatBirth As String
    Dim lstrSex As String
    

    ldatBirth = ""
    Select Case 名称
    Case "身份证号"
        lstrIDCard = ciptBase.ItemText(Index)
        If lstrIDCard <> "" Then
            '正确时从身份证号中获取出生日期。
            sub根据公民身份号码获取生日和性别 lstrIDCard, ldatBirth, lstrSex
            If Not IsDate(ldatBirth) Then
                Err.Raise 6666, , "身份证号不合法！"
            End If
            
            '查找是否需要录入出生日期，需要时自动根据身份证号填写出生日期
            On Error Resume Next
            If IsDate(ldatBirth) Then
                ciptBase.Box1("出生日期").Text = ldatBirth
                ctxtAge.Text = DateDiff("yyyy", ldatBirth, Date)
            End If
        End If
    Case "卫生种类"
        Dim lstrItemText  As String
        '设置行业类别录入框的字典。
        For i = 1 To ciptBase.InfoCollection.Count
            If ciptBase.InfoCollection(i).Title = "行业类别" Then
                If Not ciptBase.InfoCollection(Index + 1).DictRecordSet Is Nothing Then
                    If ciptBase.InfoCollection(Index + 1).DictRecordSet.EOF Then
                    Else
                        mobjGUI.sub初始化字典表 i, "Parent=" & ciptBase.InfoCollection(Index + 1).DictRecordSet("InnerId")
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
    Case "工龄"
        '有效性判断。
        If 内容 <> "" Then
            If Val(内容) > 100 Then
                Err.Raise 6666, , "工龄不能大于100！"
            End If
            If Val(内容) >= Val(ctxtAge.Text) Then
                Err.Raise 6666, , "工龄>=年龄，这是非法的数据！"
            End If
        End If
    End Select
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病界面部件", "FrmRegisterAnnual", "mobjGUI_ItemLostFocus", 6666, lstrError, False
    
    ciptBase.ItemBox(Index).Text = ""
    ciptBase.ItemSetFocus Index
End Sub

Private Sub cvscLetter_Change()
    On Error Resume Next
    '点击滚动条，获得相应的字母。
    If mcolTubeNo.Count > 0 Then
        clblLetter.Caption = mcolTubeNo.Item(cvscLetter.Value)
    End If
End Sub

'功能：清空界面。
Private Sub subClear()
    
    On Error Resume Next
    clblsysno.Text = ""
    ctxt身份证号.Text = ""
    ctxtName.Text = ""
    ccmbSex.Text = ""
    ctxtAge = ""
    ctxtTubeNo = ""
    ctxt体检单号 = ""
    mstr单位申请编号 = ""
    Ccmb体检人类别 = ""
    ccmbTemplate.Text = ""
    ccmb体检人类型.Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    '2012-07-11 于登淼 ↓
'''    '已经在unload之前，mobjGUI“退出”中判断过了
'''    '若新增体检记录没有保存，退回系统编号。
'''    If Not mobj体检 Is Nothing Then
'''        If mobj体检.系统编号 <> "" And Not mobj体检.是否已存在 Then
'''            '退回系统编号。
'''            mobj体检.sub退回系统编号 mobj体检.系统编号
'''        End If
'''    End If
    '2012-07-11 于登淼 ↑
    
'    mobj记忆.sub覆盖记忆值 "体检登记时录入单位名称", IIf(cchk录入单位名称.Value = 1, "是", "否")
     
    Set mobj体检 = Nothing
    Set mobj体检集 = Nothing
    Set mobj体检表模板 = Nothing
    '关闭相机。
    If mblnTakePhoto Then
        cctlCatchPhoto.subDisconnect
    End If
    mblnInUse = False
    pstr系统编号名称 = ""
    Dim ret
    
    ret = CloseComm()
End Sub


'功能：处理工具栏上按钮。
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
    Dim lobjFile As Object
    '2012-06-13 于登淼 ↑
    
    On Error GoTo errHandler
    
    Select Case Operate
    
    Case "清空"
        subClear
        '清空健康档案编号，表示新录入的体检人员。
        mobj体检.体检人员.健康档案编号 = ""
        clblsysno.Text = ""
        Cancel = True
        
    Case "保存"
        '2012-07-11 于登淼 ↓
        '如果校核通过，只能存储现场照片
        Cancel = True
        MousePointer = 11
'        If mintState = 1 And 访问标志 = 1 Then
       
        '2015-3-13 liuwei 修改保存照片
          Dim lobjPhoto As StdPicture
          Dim strSQL As String
        If mblnTakePhoto Then
                Set lobjPhoto = cctlCatchPhoto.Photo
            ElseIf Not Picture1.Picture Is Nothing Then
               Set lobjPhoto = Picture1.Picture
                
            End If
         pmsub保存图片 lobjPhoto, Trim(clblsysno.Text), "职业病体检"
         '2015-10-16   体检日期为当期日期
         strSQL = "update 职业病体检_体检基本信息表 set 体检日期='" & Now & "' where 系统编号='" & clblsysno & "'"
         dafuncGetData strSQL
'         dafuncGetData ("update 职业病体检_体检基本信息表 set  体检日期='" & Now & "'where 系统编号='" & paraSysNo & "'")   '2015-10-16
        访问标志 = 0
'           clblSysNo.Text = mobj体检.Func分配职业病体检系统编号 & (ccmb体检人类型.ListIndex + 1)
        Set mcol体检项目 = New Collection
        mobj体检.体检表.体检表名 = ccmbTemplate.Text
        frmRegisterManage.sub查询并显示
        Cancel = True
        cgrdHistory.rows = 1
        cgrdHistory.Visible = False
        clblHistory.Visible = False
'        Check身份证_Click


        '添加标签打印   2015-11-30 by 牟俊 ↓
        With mobj体检
        mobj体检.系统编号 = Trim(clblsysno.Text)
        mobj体检.体检人员.姓名 = Trim(ctxtName.Text)
        '添加性别和年龄 2015-12-25 by 牟俊
        mobj体检.体检人员.性别 = Trim(ccmbSex.Text)
        mobj体检.体检人员.年龄 = Trim(ctxtAge.Text)
        End With
        Dim strsql1 As String
        strsql1 = "select distinct left(体检项目,2) as 项目  from 职业病体检_体检表模板体检项目表 where 体检表名称='" & ccmbTemplate.Text & "'"
        Dim objds1 As Object
        Set objds1 = dafuncGetData(strsql1)
'        Dim lobjFile As Object
        Set lobjFile = CreateObject("职业病文书.cls文书")
        Dim zxcsysno As Collection
        Set zxcsysno = New Collection
        
        zxcsysno.Add (mobj体检.系统编号)
        lobjFile.func打印单个体检清单 Trim(clblsysno)
'        lobjFile.func打印体检清单 zxcsysno
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
        '对生化不是直接打印生化，而是打印常规项目  2015-12-10 by 牟俊   ↓
            Dim lobject As Object
'            Set lobject = dafuncGetData("select distinct right(体检项目,2) as 项目  from 职业病体检_体检表模板体检项目表 where 体检表名称='" & ccmbTemplate.Text & "'and 体检项目 like '1702%'and 体检项目<>'17020'")
            Set lobject = dafuncGetData("select distinct right(体检项目,2) as 项目编号,名称 as 项目  from 职业病体检_体检表模板体检项目表 a,职业病体检_体检项目设置表 b where a.体检表名称='" & ccmbTemplate.Text & "'and a.体检项目 like '1702%'and a.体检项目=b.编码 and a.体检项目<>'17020' and b.属性='常规'")
            If lobject.RecordCount > 0 Then
'                If lobject("项目") = "21" Then
'                lobjFile.func打印试管标签 mobj体检.系统编号, mobj体检.体检人员.姓名, "肝功1,肾功,两对半,GLU,血脂,ACP"
'                ElseIf lobject("项目") = "22" Then
'                lobjFile.func打印试管标签 mobj体检.系统编号, mobj体检.体检人员.姓名, "肝功2,肾功,GLU,血脂,ACP"
'                Else
'                lobjFile.func打印试管标签 mobj体检.系统编号, mobj体检.体检人员.姓名, "生化"
'                End If
                '打印标签的项目不是固定的，是打印体检表里有的常规项目 2015-12-25 by 牟俊
                Dim xiangmu As String
                Dim zhongjian As String
                Dim X As Integer
                lobject.MoveFirst
                For X = 0 To lobject.RecordCount - 1
                zhongjian = zhongjian + "," + lobject("项目")
                lobject.MoveNext
                Next X
                xiangmu = Right(zhongjian, Len(zhongjian) - 1)
'                lobjFile.func打印试管标签 mobj体检.系统编号, mobj体检.体检人员.姓名, xiangmu
                lobjFile.func打印试管标签 mobj体检.系统编号, mobj体检.体检人员.姓名, mobj体检.体检人员.性别, mobj体检.体检人员.年龄, xiangmu
            End If
            '2015-12-10 by  牟俊  ↑
'           lobjFile.func打印试管标签 mobj体检.系统编号, mobj体检.体检人员.姓名, "生化"
        End If
        objds1.MoveNext
        Next i
'        Unload frmProcess
        MousePointer = 0
'        'Update 体检状态
''        Dim strSQL As String
'        strSQL = "update 职业病体检_体检基本信息表 set 体检状态= '2'  where 系统编号='" & mobj体检.系统编号 & "'"
'        dafuncGetData (strSQL)
        
'        Set lobjFile = CreateObject("职业病文书.cls文书")
'        lobjFile.func打印单个体检清单 Trim(clblsysno)
        pobj业务对象.func写入单人当前体检状态 Trim(clblsysno), 2
                
        frmRegisterManage.sub查询并显示
'           frmRegisterManage.sub显示查询结果
        clblsysno.Text = ""
        subClear
         
        '重新初始化照相控件。
        If cctlCatchPhoto.Status = "恢复" Then
           cctlCatchPhoto.sub转换状态
           cctlCatchPhoto.subClear
        End If
    End Select
    Set lobjrec类型 = Nothing
    Set lobj体检表编号 = Nothing
    MousePointer = 0
    Unload Me         '照相完成后照相界面消失   2016-1-6 by 牟俊
    Exit Sub
    
errHandler:
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病界面部件", "FrmRegisterAnnual", "mobjGUI_BeforeOperate", 6666, lstrError, False
    
    MousePointer = 0
    cstbMain.Panels(1) = lstrError
    Cancel = True
    Exit Sub
    Resume
    Exit Sub
End Sub


'功能：显示指定系统编号的体检人员的信息在界面上。
Private Sub SubGetPersonInfo(ByVal para系统编号 As String)
    Dim lcolInfo As New Collection
    Dim i As Integer
    Dim j As Integer
    Dim lstrTemp As String
    Dim lstrTubeNo As String
    Dim lstrSysNo As String
    
    
    On Error GoTo errHandler
    MousePointer = 11
    
    '界面暂时不可操作。
    ctbMain.Enabled = False
    
    '先退回旧系统编号。
    If Not mobj体检.是否已存在 And mobj体检.系统编号 <> "" Then
        mobj体检.sub退回系统编号 mobj体检.系统编号
    End If
    
    '创建旧职业病对象。
    Set mobj旧体检 = CreateObject("职业病对象.clsMedicalExam")
    mobj旧体检.系统编号 = para系统编号
    
    '获得上次年检的体检表
    If ccmbTemplate.Text <> mobj旧体检.体检表.体检表名 Then
        ccmbTemplate.Text = mobj旧体检.体检表.体检表名
    
        '重新初始化录入板。
        On Error Resume Next
        mobjGUI.sub初始化录入板 mobj旧体检.体检表.体检表名
        On Error GoTo errHandler
    End If
    
    '获取旧体检记录的附加信息。
    Set lcolInfo = mobj旧体检.体检表.附加信息
    
    '填写附加信息值
    sub填录入板值 ciptBase, mobjGUI, lcolInfo
    
    '显示基本信息。
    With mobj旧体检.体检人员
        ctxtName.Text = .姓名
        ccmbSex.Text = .性别
        ctxtAge.Text = .年龄
'        ccmbUnit.Text = .单位名称
'        ccmbUnit_LostFocus
        
        '相片
        '获得并显示照片。
        If Not .像片 Is Nothing Then
            Set cctlCatchPhoto.Photo = .像片
        Else
            cctlCatchPhoto.subClear
        End If
        
        '修改：2001-8-23。
        On Error Resume Next
        mstr单位申请编号 = .单位申请编号
        
        On Error GoTo errHandler
    End With
    
    '修改：2001-12-30（显示上次体检日期）。
    Label2(4).Visible = True
    clbl旧体检日期.Visible = True
    clbl旧体检日期.Caption = mobj旧体检.体检日期
    
    '修改：2002-1-6（若时间间隔超过18个月，自动设置为初检）。
    'If IsDate(clbl旧体检日期.Caption) Then
    '    If DateDiff("m", clbl旧体检日期.Caption, Now) >= 18 Then
    '        ccmb体检人类型.ListIndex = 0
    '    Else
            '不到18个月，自动设置为年检。
    '        ccmb体检人类型.ListIndex = 1
    '    End If
    'End If
    '分配新的系统编号
    lstrSysNo = mobj体检.Func分配系统编号
    mobj体检.系统编号 = lstrSysNo
    clblsysno.Text = lstrSysNo
    
    '健康档案不变。
    mobj体检.体检人员.健康档案编号 = mobj旧体检.体检人员.健康档案编号
    
    
    '设置年检的体检表名，从而获取新试管编号。
    mobj体检.体检表.体检表名 = ccmbTemplate.Text
    
    If pobj业务对象.业务设置("试管编号自动生成") = "是" Then
        '获取本体检表的当天已使用的试管编号字母。
        clblLetter.Caption = mobj体检.体检表.试管编号字母
        If clblLetter.Caption = "" Then
            
            '该次体检登记是当天的第一个，从体检表模板对象中获取所有可选的字幕。
            mobj体检表模板.体检表名 = ccmbTemplate.Text
            lstrTubeNo = mobj体检表模板.试管字母编号
            
            '将字母按逗号分开，加入mcoltubeNo中。
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
                '试管字母改变了，给出提示。
                If clblLetter.Caption <> "" And clblLetter.Caption <> mcolTubeNo(1) Then
                    sffuncMsg "请注意，你现在选择的体检表使用的试管字母与前一个（" & clblLetter.Caption & "）不同了。"
                End If
            
                '赋值给clblLetter。
                clblLetter.Caption = mcolTubeNo(1)
                '字母可以选择。
                cvscLetter.Enabled = True
                cvscLetter.Min = 1
                cvscLetter.max = mcolTubeNo.Count
                cvscLetter.Value = 1
            Else
                ctbMain.Buttons(6).Enabled = False
                '提示该体检表无可用的字母。
                Err.Raise 6666, , "该体检表无可用试管字母编号，请先设置体检表对应的试管字母编号"
            End If
        Else
            '有字母，不能选择字母。
            cvscLetter.Enabled = False
        End If
    Else
        ctxtTubeNo = mobj体检.试管编号
    End If
    '保存按钮可用。
    ctbMain.Buttons(6).Enabled = True
    Err.Clear
    
errHandler:
    '恢复界面可操作。
    ctbMain.Enabled = True
    MousePointer = 0
    If Err <> 0 Then
        sfsub错误处理 "职业病界面部件", "FrmRegisterAnnual", "SubGetPersonInfo", Err.Number, Err.Description, True
    End If
    
    Exit Sub
    Resume
End Sub




'获取单位基本信息
Private Function func获取单位信息(单位编号 As String)
    Dim lobj单位 As Object
    On Error GoTo errHandler
    Set lobj单位 = dafuncGetData("select * from 单位档案_单位基本信息表 where 申请编号='" & 单位编号 & "'")
    If Not lobj单位.RecordCount = 0 Then
'        ccmbUnit.Text = IIf(IsNull(lobj单位("单位名称")), "", lobj单位("单位名称"))
        mstr单位申请编号 = 单位编号
'        ctxt负责人.Text = IIf(IsNull(lobj单位("负责人")), "", lobj单位("负责人"))
'        ctxt联系电话.Text = IIf(IsNull(lobj单位("电话")), "", lobj单位("电话"))
'        ccmb经济性质.Text = IIf(IsNull(lobj单位("经营形式")), "", lobj单位("经营形式"))
'        Ccmb行业类别.Text = IIf(IsNull(lobj单位("单位类别")), "", lobj单位("单位类别"))
'        ctxt单位地址.Text = IIf(IsNull(lobj单位("地址")), "", lobj单位("地址"))
    End If
    Exit Function
errHandler:
    sfsub错误处理 "职业病界面", "frmregister", "func获取单位信息", Err.Number, Err.Description, True
End Function

'二代身份证读卡器，初始化，PC与终端的连接
Private Sub sub读卡器初始化()
    'CVR_InitComm
    On Error GoTo errHandler
    'If Option1.Value = True Then
    '    List1.AddItem "【连接机具】 串口 " & comS.ListIndex + 1
    '    List1.AddItem "返回 " & CVR_InitComm(comS.ListIndex + 1)
    'Else
    '    List1.AddItem "【连接机具】 USB口 " & comU.ListIndex + 1
    '   List1.AddItem " 返回 " & CVR_InitComm(0 + 1001)
    '连接串口（COM1~COM16）或USB口(1001~1016)连接串口（COM1~COM16）或USB口(1001~1016)
      ' CVR_InitComm (0 + 1001)
    'End If
    Exit Sub
errHandler:
    sfsub错误处理 "职业病界面部件", "frmregister", "func读卡器初始化", Err.Number, Err.Description, True
End Sub

'只有经过身份证认证后，才能读取信息（照片）
Private Function func身份证认证() As Integer
'    'CVR_Authenticate
'    'On Error GoTo errHandler
'    'Dim temp As Integer
'    'List1.AddItem "【身份证认证】"
'    'List1.AddItem " 返回 " & CVR_Authenticate()
'    func身份证认证 = CVR_Authenticate()
'    Exit Function
errHandler:
    sfsub错误处理 "职业病界面部件", "frmregister", "func身份证认证", Err.Number, Err.Description, True
End Function




'用timer2 进行检测是否有身份证放在读卡器上，现设为350ms
Private Sub Timer2_Timer()
     '不知为何，总是弹出找不到termb.dll文件。于是，如下修改错误处理。
'    'On Error GoTo errHandler
'    On Error Resume Next
'    '2012-07-11 于登淼 ↑
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
'    ChDir (App.Path)                '改变当前默认路径为应用程序所在路径
'    ret = Authenticate()
'    If (ret) Then '
'       ret = ReadBaseInfos(iname, isex, folk, birthday, code, addr, agency, startdate, enddate)
'       If Trim(ctxtName.Text) <> Trim(Split(iname, "")(0)) Then
'       Dim msgs
'       msgs = MsgBox("身份证信息与体检信息不匹配。", vbOKOnly + vbInformation, "提示")
'       Exit Sub
'       End If
'       Timer2.Enabled = False
'
'        '当检测到有身份证并读取数据成功后，关闭timer2
'        'Call sub读取信息
'        ctxt身份证号.Enabled = True
'        ctxt身份证号.SetFocus      '先setfocus后lostfocus目的是，读取完身份证号后调用  ctxt身份证号.lostfocus 事件函数，以计算出性别，年龄，出生年月
'        'Call sub获取证号
'        ctxt身份证号.Text = Trim(code)
'        ctxt身份证号_KeyDown 13, 1
'        ctxt身份证号.Enabled = False
'        ctxtName.Enabled = True
'        ctxtName.SetFocus
'        'Call sub获取姓名
'        ctxtName.Text = Trim(Split(iname, "")(0))
'        ctxtName.Enabled = False
'        '2012-06-13 于登淼 ↓
'        '省疾控新要求，照相图片和身份证图片都存储
'        '显示身份证图片
'        Picture2.Picture = LoadPicture(App.Path & "\photo.bmp")
'        '2012-06-13 于登淼 ↑
'        'Call sub获取住址
'        ctxt住址.Text = Trim(addr)
'        'Call sub获取民族
'        ccmb民族.Text = Trim(folk)
'        '2012-07-12 于登淼 ↓
'        '实现关闭身份证阅读器，以免重复打开时出错
'        'ret = CloseComm()
'        '2012-07-12 于登淼 ↑
'
'        '2012-07-15 于登淼 ↓
'        '每次得到新的身份证时，查找这个人曾经所有的体检时间记录
'        sub查看历史信息 (Trim(ctxt身份证号.Text))
'        '2012-07-15 于登淼 ↑
'
'
'       '添加时间2015-2-25
''在这里查询体检人员基本信息表里的体检类型和体检类别以及危害因素。组成的结构如“职业体检-在岗期间-粉尘”
'
'Dim s危害因素 As String
'Dim s体检表类别 As String
'Dim s体检表类型 As String
'Dim strs As String
'Dim strb  As String
'Dim strx As String
' Dim rs As Object
'strs = "select 危害因素,体检表类别,体检表类型 from 职业病体检_体检人员基本信息表 where 系统编号='" & clblsysno.Text & "'"
'Set rs = dafuncGetData(strs)
's危害因素 = rs("危害因素")
's体检表类别 = rs("体检表类别")
's体检表类型 = rs("体检表类型")
'Dim s体检表信息 As String
's体检表信息 = s体检表类型 + "-" + s体检表类别 + "-" + s危害因素
'ccmbTemplate.Text = s体检表信息
'ccmb体检人类型.Text = s体检表类型
'Ccmb体检人类别.Text = s体检表类别
'   End If
'    Exit Sub
'errHandler:
'    sfsub错误处理 "职业病界面部件", "frmregister", "timer2_timer", Err.Number, Err.Description, True
'End Sub
'
''身份证认证以后，才读取信息，只有读取信息以后，才会在当前目录产生身份证上照片文件zp.mbp
''Private Sub sub读取信息()
''    'CVR_Read_Content
''    Dim mode As Integer
''    On Error GoTo errHandler
''    'mode取值：
''    '1: 生成文字wz.txt?相片数据xp.wlt和相片zp.bmp (解码)
''    '2: 生成文字wz.txt和相片数据xp.wlt
''    '4: 生成wz.txt(解码)，相片zp.bmp(解码)
''    '6: 生成以设备模块号码命名的.txt文件(解码)，相片.bmp文件(解码)
''    mode = 4
''    CVR_Read_Content (mode)
''    Exit Sub
''errHandler:
''    sfsub错误处理 "职业病界面部件", "frmregister", "sub读取信息", Err.Number, Err.Description, True
''End Sub
'''添加时间2015-2-25
''   '在这里查询体检人员基本信息表里的体检类型和体检类别以及危害因素。组成的结构如“职业体检-在岗期间-粉尘”
''   Private Sub sub获取体检表信息(ByVal para系统编号 As String)
''
''Dim s危害因素 As String
''Dim s体检类别 As String
''Dim s体检类型 As String
''Dim strs As String
''strs = "select 危害因素 from 职业病体检_体检人员基本信息表 where 系统编号='" & paraSysNo & "'"
''strb = "select 体检类别 from 职业病体检_体检人员基本信息表 where 系统编号='" & paraSysNo & "'"
''strx = "select 体检类型 from 职业病体检_体检人员基本信息表 where 系统编号='" & paraSysNo & "'"
''s危害因素 = dafuncGetData(strs)
''s体检类别 = dafuncGetData(strb)
''s体检类型 = dafuncGetData(strx)
''Dim s体检表信息 As String
''s体检表信息 = s体检类别 + "-" + s体检类型 + "-" + s危害因素
''ccmbTemplate.Text = s体检表信息
''End Sub
'''获取身份证号
''Private Sub sub获取证号()
''    Dim strTemp As String
''    Dim nReturnLen As Integer
''    Dim nReturn As Integer
''    strTemp = Space(255)
''    nReturn = GetPeopleIDCode(strTemp, nReturnLen)
''    ctxt身份证号.Text = Trim(strTemp)
''End Sub
''
''Private Sub sub获取姓名()
''    Dim strTemp As String
''    Dim nReturnLen As Integer
''    Dim nReturn As Integer
''    strTemp = Space(255)
''    nReturn = GetPeopleName(strTemp, nReturnLen)
''    ctxtName.Text = Trim(strTemp)
''End Sub
''
''Private Sub sub获取住址()
''    Dim strTemp As String
''    Dim nReturnLen As Integer
''    Dim nReturn As Integer
''    strTemp = Space(255)
''    nReturn = GetPeopleAddress(strTemp, nReturnLen)
'''    ctxt住址.Text = Trim(strTemp)
''End Sub
''
''Private Sub sub获取民族()
''    Dim strTemp As String
''    Dim nReturnLen As Integer
''    Dim nReturn As Integer
''    strTemp = Space(10)
''    nReturn = GetPeopleNation(strTemp, nReturnLen)
'''    ccmb民族.Text = Trim(strTemp)
''End Sub
'
''功能：保存职业病登记里选择的体检项目
''作者：翁乔
''时间：2012-06-04
''说明：首先要查看数据库里面是否有相同的体检项目，然后再进行增加或者修改
'
'Public Sub save优化的体检项目(ByRef para体检项目 As Collection, ByVal para系统编号 As String)
'    Dim lstrSql As String
'    Dim MedicProjt As String
'    Dim rs As Object
'    Dim i As Integer
'    Dim col体检项目 As Collection
'    On Error GoTo errHandler
'
'    Set rs = dafuncGetData("select 名称 from 系统管理_字典_字典内容表 where ID = (select ID from 系统管理_字典_字典表列表 where 名称='职业病体检科室字典') and 名称 like '%科'")
'
'    For i = 1 To rs.RecordCount
'
'        lstrSql = "delete 职业病体检_结果信息_" & rs("名称") & " where 系统编号='" & para系统编号 & "'"
'        dafuncGetData lstrSql
'        rs.MoveNext
'    Next i
'
'    Set col体检项目 = para体检项目
'
'    For i = 1 To col体检项目.Count
'        MedicProjt = Left(Trim(col体检项目(i).Item(1)), 2)
'
'        lstrSql = "select 名称 from 系统管理_字典_字典内容表 where ID = (select ID from 系统管理_字典_字典表列表 where 名称='职业病体检科室字典') and 编号= '" & MedicProjt & "'"
'        Set rs = dafuncGetData(lstrSql)
'
'        lstrSql = "insert into 职业病体检_结果信息_" & rs("名称") & "(系统编号,体检项目) values('" & para系统编号 & "','" & col体检项目(i).Item(1) & "')"
'        dafuncGetData lstrSql
'    Next i
'
'    Exit Sub
'errHandler:
'   sfsub错误处理 "职业病史录入", "clscareerhstregt", "public sub save体检项目", Err.Number, Err.Description, False
End Sub

'2012-06-25 于登淼
'添加初始化各科体检状态函数。
'用于判断每个体检人员各个体检结果(与结论)科室的体检状态。
'0代表不需要检验的科室；1代表需要检验的科室；2代表该科室已经检验完；
'3代表该科室体检结果与结论不可以再修改。(其中，2、3状态都可以下最终结论)
'状态是一个长度为13的字符串(6-25时有13个填写结果的科室，字符串长度为18)
Sub subInit各科体检状态(paraCol As Collection, paraSysNo As String)
    Dim i As Integer
    Dim paraDeptNo As Integer
    Dim paraState, strSQL As String
    
    
    For i = 1 To 17: paraState = paraState & "0": Next
    'paraState = paraState & "1"
    
    For i = 1 To paraCol.Count
        paraDeptNo = CInt(Left(paraCol.Item(i).Item(1), 2))
        paraState = Left(paraState, paraDeptNo - 1) & "1" & Right(paraState, Len(paraState) - (paraDeptNo))
    Next
    
    strSQL = "update 职业病体检_体检基本信息表 set 各科体检状态='" & paraState & "' where 系统编号='" & paraSysNo & "'"
    dafuncGetData strSQL
End Sub

'2012-07-16 于登淼
'查看固定身份证号的体检人员历时信息，并将查询结果放入cgrdHistory中
'when 体检状态=0 then '未校核'
'when 体检状态=1 then '未打清单'
'when 体检状态=2 then '未录入受检者个人信息'
'when 体检状态=3 then '体检中'
'when 体检状态=4 then '未下结论'
'when 体检状态=5 then '已下结论'
'when 体检状态=6 then '已复核'
'when 体检状态=7 then '已发报告'
'when 体检状态=8 then '待复查'
Sub sub查看历史信息(ByVal paraIDCard As String)
    Dim strSQL As String
    Dim lobjRec As Object
    Dim initState(0 To 8) As String
    Dim i, j As Integer
    '2012-12-18 刘云乐
    'bug No:0000087,0000084
'    strSQL = "select 系统编号,建档日期 建档时间,体检表编号 体检表,体检状态 from 职业病体检_体检基本数据库 where 公民身份号码='" & paraIDCard & "' and 建档日期>='" & Format(DateAdd("yyyy", -5, Now), "yyyy-mm-dd") & "'"
'    strSQL = "select 系统编号,建档日期 建档时间,体检表编号 体检表,体检状态 from 职业病体检_体检基本数据库 where 公民身份号码='" & paraIDCard & "' and 建档日期<'" & Format(Now, "yyyy-mm-dd") & "'"
    strSQL = "select 系统编号,建档日期 建档时间,体检表编号 体检表,体检状态 from 职业病体检_体检基本数据库 where 体检状态='1'"
    ''2012-12-18 刘云乐
    Set lobjRec = dafuncGetData(strSQL)
    If lobjRec.RecordCount > 0 Then
        '控制界面控件是否显示和显示格式
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
        '显示时，由于体检状态显示数字，故替换为文字。原始参照存储过程“职业病体检_体检管理界面查询”
        '初始化状态数组
        initState(0) = "未校核"
        initState(1) = "未打清单"
        initState(2) = "未录入受检者个人信息"
        initState(3) = "体检中"
        initState(4) = "未下结论"
        initState(5) = "已下结论"
        initState(6) = "已复核"
        initState(7) = "已发报告"
        initState(8) = "待复查"
        
        lobjRec.MoveFirst
        For i = 1 To lobjRec.RecordCount
            For j = 0 To lobjRec.Fields.Count - 1
                If lcolIndex(j) = "体检状态" Then cgrdHistory.TextMatrix(i, j) = initState(CInt(cgrdHistory.TextMatrix(i, j)))
            Next
            lobjRec.MoveNext
        Next
        
        cgrdHistory.AutoSize 0, cgrdHistory.cols - 1, 0, 0
    End If
End Sub

'2012-08-18 于登淼
'复查结果保存
Private Sub sub复查保存()
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
        '2012-06-13 于登淼 ↑
        
        '2012-08-18 于登淼 ↓
        '变量名重新更改。
        Dim mobj复查体检保存变量 As Object
        Set mobj复查体检保存变量 = CreateObject("职业病对象.clsMedicalExam")
        '2012-08-18 于登淼 ↑
        
        '2012-07-11 于登淼 ↓
        '如果校核通过，只能存储现场照片（复查时，访问标志=2，会继续保存后面内容。）
        MousePointer = 11
        If mintState = 1 And 访问标志 = 1 Then
            If mblnTakePhoto Then
                Dim lobjPhoto As StdPicture
                '若像片不为空，则保存到相应目录，调用的是 通用对象.cls图片管理.cls
                If Not cctlCatchPhoto.Photo Is Nothing Then
                    Set lobjPhoto = cctlCatchPhoto.Photo
                    pmsub保存图片 lobjPhoto, Trim(clblsysno.Text), "职业病体检"
                End If
            End If
            MousePointer = 0
            Set lobjrec类型 = Nothing
            Set lobj体检表编号 = Nothing
            Exit Sub
        End If
        '2012-07-11 于登淼 ↑
        
        '2012-07-11 于登淼 ↓
        '未填入身份证号或身份证号错误，不允许保存当前信息
        If Len(ctxt身份证号.Text) = 0 Then
            MousePointer = 0
            MsgBox ("未填入身份证号，不允许保存当前内容！")
            Exit Sub
        End If
        sub根据公民身份号码获取生日和性别 ctxt身份证号.Text, lstrBirth, lstrSex
        If ccmbSex.Text = "" Or lstrSex <> ccmbSex.Text Or Format(lstrBirth, "yyyy-mm-dd") <> Format(cdtp出生.Value, "yyyy-mm-dd") Then
            MsgBox ("身份证与当前个人信息不符，不允许保存当前内容！")
            Exit Sub
        End If
        '2012-07-11 于登淼 ↑
        
        '移走 系统编号 文本框 焦点
'        ctxt籍贯.SetFocus
        
        '2012-06-13 于登淼 ↓
        '省疾控新要求，照相照片与身份证照片分开存储，分开显示
        '这里单独存储身份证照片。照相照片按原来方法存储。
        Set lobjRec身份证照片 = CreateObject("职业病对象.clsPersonExamed")
        lobjRec身份证照片.func保存身份证照片 Picture2.Image, pstr复查系统编号 & "IDcard", "职业病体检"
        Set lobjRec身份证照片 = Nothing
        '2012-06-13 于登淼 ↑
        
        MousePointer = 11
        
        Set lobj体检表编号 = CreateObject("职业病对象.clsmedicalexamsheet")
        lobj体检表编号.体检表编号 = ccmbTemplate.Text
        
        pstr系统编号 = clblsysno.Text
        '产生试管编号并保存
        With mobj复查体检保存变量
            '2012-06-14 于登淼 ↓
            '系统编号必须在这里重新赋值，否则第一次会使用form_load时的系统编号
            .系统编号 = pstr复查系统编号
            '2012-06-14 于登淼 ↑
            
            If .体检表.体检表名 <> ccmbTemplate.Text Then
                .体检表.体检表名 = ccmbTemplate.Text
            End If
            '修改：2004-1-9（试管编号可以输入）
            If pobj业务对象.业务设置("试管编号自动生成") = "是" Then
                If .体检表.试管编号字母 <> clblLetter.Caption Then
                    .体检表.试管编号字母 = clblLetter.Caption
                End If
            Else
                .体检表.试管编号字母 = clblLetter.Caption
                .试管编号 = ctxtTubeNo.Text
            End If
            .体检人员.系统编号 = pstr复查系统编号
            .体检人员.姓名 = ctxtName
            .体检人员.性别 = ccmbSex.Text
'            .体检人员.单位名称 = ccmbUnit.Text
'            .体检人员.危害因素 = ccmb危害因素.Text
'            .体检人员.照射源 = ccmb照射源.Text
'            .体检人员.职业分类 = ccmb职业类别.Text
'            .体检人员.现工种 = ccmb现工种.Text
'            .体检人员.职务或职称 = ccmb职务.Text
'            .体检人员.职业危害工龄 = ctxt危害工龄.Text
'            .体检人员.放射剂量 = Trim(ctxt放射剂量.Text)
'            .体检人员.籍贯 = Trim(ctxt籍贯.Text)
'            .体检人员.邮编 = Trim(ctxt邮编.Text)
'            .体检人员.住址 = Trim(ctxt住址.Text)
'            .体检人员.婚否 = Ccmb婚否.Text
'            .体检人员.电话号码 = Trim(ctxt电话.Text)
'            .体检人员.工龄 = Trim(ctxt工龄.Text)
'            .体检人员.出生地 = Trim(ctxt出生地.Text)
'            .体检人员.负责人 = ctxt负责人.Text
'            .体检人员.联系电话 = ctxt联系电话.Text
'            .体检人员.经济性质 = ccmb经济性质.Text
'            .体检人员.行业类别 = Ccmb行业类别.Text
'            .体检人员.单位地址 = ctxt单位地址.Text
            If mblnTakePhoto Then
                .体检人员.像片 = cctlCatchPhoto.Photo
'                .体检人员.像片压缩 = cctlCatchPhoto.Photo
            ElseIf Not Picture1.Picture Is Nothing Then
                .体检人员.像片 = Picture1.Picture
            End If
            If Val(ctxtAge.Text) > 0 Then
'                If Val(ctxtAge.Text) > 200 Then
'                    Err.Raise 6666, , "年龄超过系统允许的最大数：200。"
'                End If
                .体检人员.出生日期 = DateAdd("yyyy", -Val(ctxtAge.Text), Date)
            Else
                '如果输入字符，则记忆该年龄。
                mobj记忆.sub覆盖记忆值 "体检年龄", ctxtAge.Text
                mstr默认年龄 = ctxtAge.Text
            End If
            .体检人员.年龄 = ctxtAge.Text
            
            On Error Resume Next
            .体检人员.公民身份号码 = ctxt身份证号.Text
'            .体检人员.文化程度 = ccmb文化程度.Text
'            .体检人员.民族 = ccmb民族.Text
'            If ccmbUnit.Text = "" Then
'                .体检人员.单位申请编号 = ""
'            Else
'                If .体检人员.单位申请编号 <> mstr单位申请编号 Then
'                    '给单位编号重新赋值，可以重新获取其卫生种类、行业类别、片区。
'                    .体检人员.单位申请编号 = mstr单位申请编号
'                End If
'            End If
            
            .体检日期 = cdtpDate.Value ' ,Format(cdtpDate.Value, "yyyy-mm-dd hh:mm:ss")
            
            '修改：2004-1-9（增加体检单号）
            .体检人类型 = ccmb体检人类型.Text
            .体检人类别 = Ccmb体检人类别.Text
            
            'On Error GoTo errHandler
            On Error Resume Next
            If mcol体检项目.Count = 0 Then
'                mobj复查体检保存变量.体检表.mbln是否已存在 = True
'                mobj复查体检保存变量.体检表.mbln是否已获取体检项目 = False
'                mobj复查体检保存变量.体检表.mbln是否已获取附加项目 = False
                Set mcol体检项目 = mobj体检.体检表.体检项目集("")
                frmSelectItem.pstr体检表名称 = ccmbTemplate.Text
                Set frmSelectItem.pcol复查项目 = mcol体检项目
                frmSelectItem.Hide
                frmSelectItem.ccmdOk_Click
                Set mcol体检项目 = frmSelectItem.pcol复查项目
            End If
            Set .col体检项目 = mcol体检项目
        
        End With
        
        '功能：保存体检项目
        '时间：2012-06-04
        '作者：翁乔
       ' save优化的体检项目 mcol体检项目, pstr复查系统编号
        '时间：2012-06-04
        
        If mcol收费项目.Count > 0 Then
            pobj业务对象.Sub体检登记 mobj复查体检保存变量, , , mcol收费项目, Val(ctxt份数)
        Else
            pobj业务对象.Sub体检登记 mobj复查体检保存变量, , , , Val(ctxt份数)
        End If
        
        Set lobjRec = CreateObject("职业病界面.clsMoney")
        lobjRec.mstr系统编号 = pstr复查系统编号
        lobjRec.mstr体检人员姓名 = ctxtName.Text
        Set lobjRec.col体检项目 = mcol体检项目
        Dim lstr收费批号 As String
        lstrError = lobjRec.func收费(lstr收费批号)
        mobj复查体检保存变量.收费批号 = lstr收费批号
        If lstrError <> "" And lstrError <> "Cancel" Then
            MsgBox lstrError, vbOKOnly + vbExclamation, "系统提示"
        End If
    
        cstbMain.Panels(1) = "上次保存的体检系统编号：" & mobj复查体检保存变量.系统编号 '& " ，试管编号：" & mobj复查体检保存变量.试管编号
        If mobj复查体检保存变量.收费批号 <> "" Then
            cstbMain.Panels(1) = cstbMain.Panels(1) & "，收费批号：" & mobj复查体检保存变量.收费批号
        End If
        
        '2012-06-25 于登淼 ↓
        '初始化体检基本信息表中“各科体检状态”字段
        subInit各科体检状态 mcol体检项目, pstr复查系统编号
        '2012-06-25 于登淼 ↑
        '登记后修改复查状态，翁乔；2012-10-30
        dafuncGetData "update 职业病体检_体检基本信息表 set 复查状态 = '1' where 复查系统编号 = '" & pstr复查系统编号 & "'"
        '2012-08-18 于登淼 ↓
        '复查登记后，认为是校核后，未打印清单状态，需要写入当前体检状态且照相。
        '这个与初检登记不同。
        If mintState = 1 Then
            pobj业务对象.func写入单人当前体检状态 pstr复查系统编号, mintState
            pobj业务对象.func写入单人当前体检状态 pstr系统编号, 5   '5代表体检状态"已下结论"
            mobjGUI_BeforeOperate "校核通过", False
        End If
        mintState = 2
        '2012-06-15 于登淼 ↑
        
        '恢复照相。
        If mblnTakePhoto Then
            If cctlCatchPhoto.Status = "恢复" Then
                cctlCatchPhoto.sub转换状态
            End If
        End If
        
        If cchkClear = 1 Then
            subClear
            访问标志 = 0
            clblsysno.Text = mobj体检.Func分配职业病体检系统编号 & (ccmb体检人类型.ListIndex + 1)
        End If
        
        Set mcol体检项目 = New Collection
       
        mobj体检.体检表.体检表名 = ccmbTemplate.Text
               
        '试管字母不能再选择。
        cvscLetter.Enabled = False
'        ctxt籍贯.SetFocus
        frmRegisterManage.sub查询并显示
        Timer2.Enabled = True

        MousePointer = 0
        Set lobjrec类型 = Nothing
        Set lobj体检表编号 = Nothing
        Exit Sub
errHandler:
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病界面部件", "FrmRegisterAnnual", "mobjGUI_BeforeOperate", 6666, lstrError, False
    
    MousePointer = 0
    cstbMain.Panels(1) = lstrError
    Exit Sub
    Resume
End Sub
