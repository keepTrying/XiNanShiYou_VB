VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FC07EBD4-FE92-11D0-A199-A0077383D901}#5.1#0"; "CCRPPRG.OCX"
Begin VB.Form frmURT_ResultInput 
   Caption         =   "尿常规科结果录入窗口"
   ClientHeight    =   10275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13350
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10275
   ScaleWidth      =   13350
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox Picture1 
      Height          =   10095
      Left            =   0
      ScaleHeight     =   10035
      ScaleWidth      =   13155
      TabIndex        =   0
      Top             =   0
      Width           =   13215
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   9975
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   13095
         Begin VB.OptionButton coptClasses 
            BackColor       =   &H00C0FFFF&
            Caption         =   "放射健康"
            Height          =   255
            Index           =   2
            Left            =   2400
            TabIndex        =   73
            Top             =   960
            Width           =   1095
         End
         Begin VB.OptionButton coptClasses 
            BackColor       =   &H00C0FFFF&
            Caption         =   "职业健康"
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   72
            Top             =   960
            Width           =   1095
         End
         Begin VB.OptionButton coptClasses 
            BackColor       =   &H00C0FFFF&
            Caption         =   "普通体检"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   71
            Top             =   960
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton coptClasses 
            BackColor       =   &H00C0FFFF&
            Caption         =   "涉核部队"
            Height          =   255
            Index           =   3
            Left            =   3480
            TabIndex        =   70
            Top             =   960
            Width           =   1095
         End
         Begin VB.OptionButton coptClasses 
            BackColor       =   &H00C0FFFF&
            Caption         =   "8023部队"
            Height          =   255
            Index           =   4
            Left            =   4560
            TabIndex        =   69
            Top             =   960
            Width           =   1095
         End
         Begin VB.CommandButton Cmd结论模板 
            Caption         =   "结论模板"
            Height          =   495
            Left            =   11760
            TabIndex        =   66
            Top             =   7080
            Width           =   1215
         End
         Begin VB.Frame Frame5 
            Caption         =   "结论录入 (不超过250个汉字)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   2295
            Left            =   7320
            TabIndex        =   12
            Top             =   7560
            Width           =   5655
            Begin VB.TextBox ctxtConclun 
               Height          =   1935
               Left            =   0
               MultiLine       =   -1  'True
               TabIndex        =   13
               Top             =   360
               Width           =   5655
            End
         End
         Begin VB.CommandButton ccmdAutoFull 
            Caption         =   "清 空 表 格"
            Height          =   375
            Index           =   3
            Left            =   10440
            TabIndex        =   11
            Top             =   2040
            Width           =   1695
         End
         Begin VB.CommandButton ccmdAutoFull 
            Caption         =   "恢 复 默 认"
            Height          =   375
            Index           =   2
            Left            =   7920
            TabIndex        =   10
            Top             =   2040
            Width           =   1695
         End
         Begin VB.CommandButton ccmdAutoFull 
            Caption         =   "全 部 正 常"
            Height          =   375
            Index           =   1
            Left            =   10440
            TabIndex        =   9
            Top             =   1560
            Width           =   1695
         End
         Begin VB.CommandButton ccmdAutoFull 
            Caption         =   "未填写项全部正常"
            Height          =   375
            Index           =   0
            Left            =   7920
            TabIndex        =   8
            Top             =   1560
            Width           =   1695
         End
         Begin VB.Frame Frame3 
            Height          =   580
            Left            =   7680
            TabIndex        =   4
            Top             =   840
            Visible         =   0   'False
            Width           =   4815
            Begin VB.CommandButton WriteConclun 
               Caption         =   "手写结论"
               Height          =   375
               Left            =   2640
               TabIndex        =   6
               Top             =   120
               Width           =   1455
            End
            Begin VB.TextBox ctxtDoctor 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   720
               TabIndex        =   5
               Top             =   120
               Width           =   1335
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "医师："
               Height          =   180
               Index           =   1
               Left            =   120
               TabIndex        =   7
               Top             =   240
               Width           =   540
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "体检项目结果填写： "
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   4455
            Left            =   7320
            TabIndex        =   2
            Top             =   2520
            Width           =   5700
            Begin VSFlex8Ctl.VSFlexGrid cgrdInput 
               Height          =   3975
               Left            =   240
               TabIndex        =   3
               Top             =   360
               Width           =   5175
               _cx             =   2088772520
               _cy             =   2088770403
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
               AllowBigSelection=   -1  'True
               AllowUserResizing=   1
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   1
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
         End
         Begin TabDlg.SSTab ctabPerson 
            Height          =   8415
            Left            =   120
            TabIndex        =   14
            Top             =   1440
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   14843
            _Version        =   393216
            Tabs            =   2
            Tab             =   1
            TabHeight       =   520
            ForeColor       =   8388608
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "   单个处理  "
            TabPicture(0)   =   "frmURT_ResultInput.frx":0000
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "Frame2"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "Frame4"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).ControlCount=   2
            TabCaption(1)   =   "   批量处理 "
            TabPicture(1)   =   "frmURT_ResultInput.frx":001C
            Tab(1).ControlEnabled=   -1  'True
            Tab(1).Control(0)=   "TotalPeopleBatch"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).Control(1)=   "Label6"
            Tab(1).Control(1).Enabled=   0   'False
            Tab(1).Control(2)=   "ccrp进度"
            Tab(1).Control(2).Enabled=   0   'False
            Tab(1).Control(3)=   "cdtpDateBatch"
            Tab(1).Control(3).Enabled=   0   'False
            Tab(1).Control(4)=   "cgrdInfoBatch"
            Tab(1).Control(4).Enabled=   0   'False
            Tab(1).Control(5)=   "Timerccrp"
            Tab(1).Control(5).Enabled=   0   'False
            Tab(1).Control(6)=   "ccmdSelInfo"
            Tab(1).Control(6).Enabled=   0   'False
            Tab(1).Control(7)=   "cchkCompanyBatch"
            Tab(1).Control(7).Enabled=   0   'False
            Tab(1).Control(8)=   "ctxtQueyCompanyBatch"
            Tab(1).Control(8).Enabled=   0   'False
            Tab(1).Control(9)=   "ccmd查询单位"
            Tab(1).Control(9).Enabled=   0   'False
            Tab(1).Control(10)=   "fraQueryBatch"
            Tab(1).Control(10).Enabled=   0   'False
            Tab(1).Control(11)=   "cchkDateBatch"
            Tab(1).Control(11).Enabled=   0   'False
            Tab(1).Control(12)=   "ccmdClear"
            Tab(1).Control(12).Enabled=   0   'False
            Tab(1).Control(13)=   "ccmdRemove"
            Tab(1).Control(13).Enabled=   0   'False
            Tab(1).Control(14)=   "cchkBchResult(0)"
            Tab(1).Control(14).Enabled=   0   'False
            Tab(1).Control(15)=   "cchkBchResult(1)"
            Tab(1).Control(15).Enabled=   0   'False
            Tab(1).ControlCount=   16
            Begin VB.CheckBox cchkBchResult 
               Caption         =   "未填结果"
               Height          =   255
               Index           =   1
               Left            =   3000
               TabIndex        =   77
               Top             =   4440
               Value           =   1  'Checked
               Width           =   1095
            End
            Begin VB.CheckBox cchkBchResult 
               Caption         =   "已填结果"
               Height          =   255
               Index           =   0
               Left            =   1680
               TabIndex        =   76
               Top             =   4440
               Width           =   1095
            End
            Begin VB.CommandButton ccmdRemove 
               Caption         =   "移 除"
               Height          =   375
               Left            =   5640
               TabIndex        =   63
               Top             =   3600
               Width           =   855
            End
            Begin VB.CommandButton ccmdClear 
               Caption         =   "清 空"
               Height          =   375
               Left            =   4440
               TabIndex        =   62
               Top             =   3600
               Width           =   855
            End
            Begin VB.CheckBox cchkDateBatch 
               BackColor       =   &H00C0FFC0&
               Caption         =   "体检日期"
               Height          =   255
               Left            =   360
               TabIndex        =   61
               Top             =   3600
               Width           =   1215
            End
            Begin VB.Frame fraQueryBatch 
               Caption         =   "批量查询体检人员"
               Height          =   3015
               Left            =   240
               TabIndex        =   46
               Top             =   480
               Width           =   6615
               Begin VB.PictureBox Picture4 
                  Height          =   1935
                  Left            =   4680
                  ScaleHeight     =   1875
                  ScaleWidth      =   1515
                  TabIndex        =   53
                  Top             =   240
                  Width           =   1575
               End
               Begin VB.TextBox ctxt单位名称 
                  Height          =   300
                  Left            =   1680
                  TabIndex        =   52
                  Top             =   2160
                  Width           =   2415
               End
               Begin VB.TextBox ctxt年龄 
                  Height          =   300
                  Left            =   1680
                  TabIndex        =   51
                  Top             =   1800
                  Width           =   2415
               End
               Begin VB.TextBox ctxt性别 
                  Height          =   300
                  Left            =   1680
                  TabIndex        =   50
                  Top             =   1440
                  Width           =   2415
               End
               Begin VB.TextBox ctxt姓名 
                  Height          =   300
                  Left            =   1680
                  TabIndex        =   49
                  Top             =   1080
                  Width           =   2415
               End
               Begin VB.TextBox ctxt体检条码 
                  Height          =   300
                  Left            =   1680
                  TabIndex        =   48
                  Top             =   720
                  Width           =   2415
               End
               Begin VB.CheckBox cchk套用体检结果 
                  BackColor       =   &H008080FF&
                  Caption         =   "该体检人员结果作为批量体检结果录入"
                  Height          =   255
                  Left            =   480
                  TabIndex        =   47
                  Top             =   2640
                  Value           =   1  'Checked
                  Width           =   3615
               End
               Begin MSComCtl2.DTPicker DTP录入日期 
                  Height          =   300
                  Left            =   1680
                  TabIndex        =   54
                  Top             =   360
                  Width           =   2415
                  _ExtentX        =   4260
                  _ExtentY        =   529
                  _Version        =   393216
                  Enabled         =   0   'False
                  Format          =   72679424
                  CurrentDate     =   40969
               End
               Begin VB.Label Label11 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "结论录入日期"
                  Height          =   255
                  Left            =   480
                  TabIndex        =   60
                  Top             =   360
                  Width           =   1095
               End
               Begin VB.Label Label14 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "单位名称"
                  Height          =   255
                  Left            =   480
                  TabIndex        =   59
                  Top             =   2160
                  Width           =   975
               End
               Begin VB.Label Label15 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "年龄"
                  Height          =   255
                  Left            =   480
                  TabIndex        =   58
                  Top             =   1800
                  Width           =   975
               End
               Begin VB.Label Label16 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "性别"
                  Height          =   255
                  Left            =   480
                  TabIndex        =   57
                  Top             =   1440
                  Width           =   975
               End
               Begin VB.Label Label17 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "姓名"
                  Height          =   255
                  Left            =   480
                  TabIndex        =   56
                  Top             =   1080
                  Width           =   975
               End
               Begin VB.Label Label18 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "体检条码号"
                  Height          =   255
                  Left            =   480
                  TabIndex        =   55
                  Top             =   720
                  Width           =   975
               End
            End
            Begin VB.CommandButton ccmd查询单位 
               Caption         =   "单位定位"
               Height          =   375
               Left            =   4440
               TabIndex        =   45
               Top             =   4080
               Width           =   855
            End
            Begin VB.TextBox ctxtQueyCompanyBatch 
               Height          =   300
               Left            =   1680
               TabIndex        =   44
               Top             =   4080
               Width           =   2415
            End
            Begin VB.CheckBox cchkCompanyBatch 
               BackColor       =   &H00C0FFC0&
               Caption         =   "单位名称"
               Height          =   255
               Left            =   360
               TabIndex        =   43
               Top             =   4080
               Width           =   1215
            End
            Begin VB.CommandButton ccmdSelInfo 
               Caption         =   "查 询"
               Height          =   375
               Left            =   5640
               TabIndex        =   42
               Top             =   4080
               Width           =   855
            End
            Begin VB.Timer Timerccrp 
               Left            =   6480
               Top             =   3960
            End
            Begin VB.Frame Frame2 
               Caption         =   "查询体检人员"
               Height          =   5055
               Left            =   -74880
               TabIndex        =   29
               Top             =   3240
               Width           =   6855
               Begin VB.CommandButton ccmdWork 
                  Caption         =   "单位定位"
                  Height          =   375
                  Left            =   3600
                  TabIndex        =   84
                  Top             =   960
                  Width           =   1185
               End
               Begin VB.CheckBox cchkSingleNo 
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "体检条码"
                  Height          =   255
                  Left            =   3600
                  TabIndex        =   83
                  Top             =   240
                  Width           =   1095
               End
               Begin VB.TextBox ctxtcchkNo 
                  Height          =   270
                  Left            =   4800
                  TabIndex        =   82
                  Top             =   240
                  Width           =   1695
               End
               Begin VB.CheckBox cchkCardNo 
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "身份证号"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   81
                  Top             =   600
                  Width           =   1095
               End
               Begin VB.CheckBox cchkWorkUnit 
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "单位名称"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   80
                  Top             =   960
                  Width           =   1095
               End
               Begin VB.TextBox ctxtcchkCardNo 
                  Height          =   270
                  Left            =   1560
                  TabIndex        =   79
                  Top             =   600
                  Width           =   1695
               End
               Begin VB.TextBox ctxtcchkWork 
                  Height          =   270
                  Left            =   1560
                  TabIndex        =   78
                  Top             =   960
                  Width           =   1695
               End
               Begin VB.CheckBox cchkSigResult 
                  Caption         =   "未填结果"
                  Height          =   255
                  Index           =   1
                  Left            =   1800
                  TabIndex        =   75
                  Top             =   1320
                  Value           =   1  'Checked
                  Width           =   1095
               End
               Begin VB.CheckBox cchkSigResult 
                  Caption         =   "已填结果"
                  Height          =   255
                  Index           =   0
                  Left            =   240
                  TabIndex        =   74
                  Top             =   1320
                  Width           =   1095
               End
               Begin VB.CheckBox cchkName 
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "姓名"
                  Height          =   255
                  Left            =   3600
                  TabIndex        =   33
                  Top             =   600
                  Width           =   735
               End
               Begin VB.TextBox ctxtCheckName 
                  Height          =   270
                  Left            =   4800
                  TabIndex        =   32
                  Top             =   600
                  Width           =   1695
               End
               Begin VB.CommandButton ccmdSingleQuery 
                  Caption         =   "查询(&Q)"
                  Height          =   375
                  Left            =   4920
                  Style           =   1  'Graphical
                  TabIndex        =   31
                  Top             =   960
                  Width           =   1185
               End
               Begin VB.CheckBox cchkDate 
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "体检日期"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   30
                  Top             =   240
                  Width           =   1095
               End
               Begin VSFlex8Ctl.VSFlexGrid cgrdSingleList 
                  Height          =   3255
                  Left            =   120
                  TabIndex        =   34
                  Top             =   1680
                  Width           =   6615
                  _cx             =   2088775060
                  _cy             =   2088769133
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
                  AllowBigSelection=   -1  'True
                  AllowUserResizing=   1
                  SelectionMode   =   3
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   1
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
                  Height          =   255
                  Left            =   1560
                  TabIndex        =   35
                  Top             =   240
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   450
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
                  Format          =   72679424
                  CurrentDate     =   36957
                  MaxDate         =   73050
                  MinDate         =   17899
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "人数："
                  Height          =   180
                  Left            =   5640
                  TabIndex        =   37
                  Top             =   1320
                  Width           =   540
               End
               Begin VB.Label TotalPeople 
                  AutoSize        =   -1  'True
                  Caption         =   "0"
                  Height          =   180
                  Left            =   6240
                  TabIndex        =   36
                  Top             =   1320
                  Width           =   90
               End
            End
            Begin VB.Frame Frame4 
               Caption         =   "体检人员基本信息   "
               ForeColor       =   &H000080FF&
               Height          =   2775
               Left            =   -74880
               TabIndex        =   15
               Top             =   360
               Width           =   6855
               Begin VB.ComboBox ccmbHistory 
                  Height          =   300
                  Left            =   1560
                  Style           =   2  'Dropdown List
                  TabIndex        =   85
                  Top             =   600
                  Width           =   2295
               End
               Begin VB.TextBox ctxtName 
                  Height          =   270
                  Left            =   1560
                  TabIndex        =   21
                  Top             =   1320
                  Width           =   2655
               End
               Begin VB.TextBox ctxtSex 
                  Height          =   270
                  Left            =   1560
                  TabIndex        =   20
                  Top             =   1680
                  Width           =   2655
               End
               Begin VB.TextBox ctxtAge 
                  Height          =   270
                  Left            =   1560
                  TabIndex        =   19
                  Top             =   2040
                  Width           =   2655
               End
               Begin VB.TextBox ctxtCompanyName 
                  Height          =   270
                  Left            =   1560
                  TabIndex        =   18
                  Top             =   2400
                  Width           =   3495
               End
               Begin VB.PictureBox cpicPhoto 
                  BeginProperty Font 
                     Name            =   "宋体"
                     Size            =   10.5
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1890
                  Index           =   0
                  Left            =   5040
                  ScaleHeight     =   1830
                  ScaleWidth      =   1515
                  TabIndex        =   17
                  Top             =   240
                  Width           =   1575
               End
               Begin VB.TextBox ctxtSingleNo 
                  Height          =   270
                  Left            =   1560
                  MaxLength       =   20
                  TabIndex        =   16
                  Top             =   960
                  Width           =   2655
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
                  Height          =   255
                  Left            =   1560
                  TabIndex        =   22
                  Top             =   240
                  Width           =   2295
                  _ExtentX        =   4048
                  _ExtentY        =   450
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
                  Format          =   72679424
                  CurrentDate     =   36951
                  MaxDate         =   73050
                  MinDate         =   17899
               End
               Begin VB.Label Label13 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "历年病历"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   86
                  Top             =   600
                  Width           =   975
               End
               Begin VB.Label Label5 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "单位名称"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   28
                  Top             =   2400
                  Width           =   975
               End
               Begin VB.Label Label3 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "性别"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   27
                  Top             =   1680
                  Width           =   975
               End
               Begin VB.Label Label4 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "年龄"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   26
                  Top             =   2040
                  Width           =   975
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "结论录入日期"
                  Height          =   255
                  Index           =   2
                  Left            =   240
                  TabIndex        =   25
                  Top             =   240
                  Width           =   1080
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "体检条码号"
                  Height          =   255
                  Index           =   3
                  Left            =   240
                  TabIndex        =   24
                  Top             =   960
                  Width           =   975
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "姓名"
                  Height          =   255
                  Index           =   4
                  Left            =   240
                  TabIndex        =   23
                  Top             =   1320
                  Width           =   975
               End
            End
            Begin VSFlex8Ctl.VSFlexGrid cgrdInfoBatch 
               Height          =   3015
               Left            =   240
               TabIndex        =   41
               Top             =   5280
               Width           =   6615
               _cx             =   2088775060
               _cy             =   2088768710
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
               AllowUserResizing=   1
               SelectionMode   =   3
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   1
               Cols            =   4
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   "体检条码编号"
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
            Begin MSComCtl2.DTPicker cdtpDateBatch 
               Height          =   300
               Left            =   1680
               TabIndex        =   64
               Top             =   3600
               Width           =   2415
               _ExtentX        =   4260
               _ExtentY        =   529
               _Version        =   393216
               Format          =   72679424
               CurrentDate     =   40969
            End
            Begin CCRProgressBar.ccrpProgressBar ccrp进度 
               Height          =   375
               Left            =   480
               Top             =   4800
               Visible         =   0   'False
               Width           =   6015
               _ExtentX        =   10610
               _ExtentY        =   661
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
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0E0FF&
               Caption         =   "人数："
               Height          =   180
               Left            =   360
               TabIndex        =   68
               Top             =   4440
               Width           =   540
            End
            Begin VB.Label TotalPeopleBatch 
               AutoSize        =   -1  'True
               Caption         =   "0"
               Height          =   180
               Left            =   960
               TabIndex        =   67
               Top             =   4440
               Width           =   90
            End
         End
         Begin MSComctlLib.StatusBar cstbMain 
            Height          =   375
            Left            =   120
            TabIndex        =   38
            Top             =   11475
            Width           =   13515
            _ExtentX        =   23839
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
               NumPanels       =   1
               BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
                  AutoSize        =   1
                  Object.Width           =   23786
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
            Height          =   540
            Left            =   120
            TabIndex        =   39
            Top             =   120
            Width           =   12915
            _ExtentX        =   22781
            _ExtentY        =   953
            ButtonWidth     =   820
            ButtonHeight    =   953
            Appearance      =   1
            Style           =   1
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
            Begin VB.CheckBox cchk刷条码 
               Caption         =   "刷条码"
               Height          =   255
               Left            =   9600
               TabIndex        =   40
               Top             =   240
               Width           =   1215
            End
         End
         Begin VB.Label LabelDoctor 
            BackColor       =   &H00C0FFFF&
            Caption         =   "医生："
            Height          =   255
            Left            =   5760
            TabIndex        =   65
            Top             =   960
            Width           =   1575
         End
      End
   End
End
Attribute VB_Name = "frmURT_ResultInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'2012-04-01 于登淼
'添加窗体，参照尿常规化验科结果录入的所有内容。

Option Explicit
Private WithEvents mobj界面通用对象 As cls界面通用对象    '提供工具栏初始化、热键处理。
Attribute mobj界面通用对象.VB_VarHelpID = -1
Private mobj体检医师  As Object   'clsMedicalExamer    获取当前体检医师可以作的指定属性（常规/化验）的体检项目
Private mstr体检表名称 As String  '批量处理时，当前一批体检记录所使用的体检表模板名称。
Private mstr操作名称  As String   '对应属性"操作名称"。
Private mstr体检项目属性 As String
Private mstr系统编号固定部分 As String

Private mobjQueryResult As Object
'2012-07-14 于登淼 ↓
'增加科室基本信息变量
Private priDeptName As String
Private priDeptNo As String
Private priDeptResultName As String
'2012-07-14 于登淼 ↑

'操作权限标志
'记录在第一次保存体检结果之后，如果再次修改结果，需要弹出“结果已修改，是否保存”之类的提示。
'-1，表示未获取该人数据库里体检结果的信息；
'0，表示该人的结果未录入过；
'1，表示数据库里已有该人的结果，但在界面上未被修改过；
'2，表示数据库里已有该人的结果，界面上已修改过。只有在为2的时候，才会弹出“结果已修改，是否保存”窗口
'3，表示没有权限进行修改操作。
Private ResultChanged As Integer

Private lcolResult As Collection    '体检结果集合，item:[体检项目名称，体检结果]。
Private lcolItem As Collection      '单个体检项目的体检结果：[体检项目名称，体检结果]。

Private mstrState As String     '记录当前体检状态
Private mblnSys As Boolean
Public mblnInUse As Boolean      '对应属性"pblnInUse"
Private mobj记忆 As cls用户操作记忆
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
       ' ccmdAdd_Click
    End If
End Sub

Private Sub cchkUnEnd_Click(Index As Integer)
    If Index = 0 Then
        ccmdSingleQuery_Click
    Else
        'ccmdAdd_Click
    End If
End Sub


'2012-07-14 于登淼
Private Sub cchkBchResult_Click(Index As Integer)
    Dim i, coptIndex As Integer
    For i = 0 To coptClasses.Count - 1
        If coptClasses(i).Value = True Then coptIndex = i
    Next
    sub查询列表显示 coptIndex
End Sub

'2012-07-14 于登淼
Private Sub cchkSigResult_Click(Index As Integer)
    Dim i, coptIndex As Integer
    For i = 0 To coptClasses.Count - 1
        If coptClasses(i).Value = True Then coptIndex = i
    Next
    sub查询列表显示 coptIndex
End Sub

'显示选中日期的病历信息
'翁乔
'2012-07-31
Private Sub ccmbHistory_Click()
    Dim lcolInfo As Collection '存放当前系统编号中当前医师可以做得指定属性的体检项目及其结果。
    Dim lobjItem As Variant    'clsFactTestItem,lcolInfo中的元素。
    Dim lstrEnum As String
    Dim i As Long
    Dim j As Long
'    On Error Resume Next
    
    '使用优化算法获取体检项目。
    Dim lobjRec As Object
    Dim lobjTemp As Object
    If ccmbHistory.Text <> "——" Then
        ctbMain.Buttons(2).Enabled = False
        Set lobjRec = mobj体检医师.func获取指定年份的体检病历(Trim(ctxtSingleNo.Text), ccmbHistory.Text, "尿常规化验科")

        If Not lobjRec Is Nothing Then
        
            '调整项目结果显示界面
'            Chk套用模版.Visible = False
            Cmd结论模板.Visible = False
'            Frame5.Visible = False
            Frame6.Caption = "体检人员历年病历："
'            Frame6.Height = Frame6.Height + 300
'            cgrdInput.Height = cgrdInput.Height - 300
            
            
            '更改当前录入状态
            If IsNull(lobjRec("单项结论")) Then
                ResultChanged = IIf(ResultChanged <> 3, 0, 3)
            Else
                ResultChanged = IIf(ResultChanged <> 3, 1, 3)
            End If
            
            cgrdInput.rows = lobjRec.recordcount + 1
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
    
                i = i + 1
                lobjRec.MoveNext
            Loop
            '添加控制列表显示函数
            With cgrdInput
                .col = 0
                .Sort = flexSortGenericAscending
                .AutoSize 0, .cols - 1, 0, 0
                .ExplorerBar = flexExSort
                .DataMode = flexDMFree
            End With
            
            Set lobjRec = mobj体检医师.func获取指定年份的体检病历结论(Trim(ctxtSingleNo.Text), "06", Trim(ccmbHistory.Text))
            If Not lobjRec Is Nothing Then
                ctxtConclun.Text = lobjRec("文字结论")
            End If
    
            cgrdInput.Select 1, 2, 1, 2
'            cgrdInput.Enabled = False
        Else
            cgrdInput.rows = 1
            ctbMain.Buttons(3).Enabled = False
            
        End If
        
    ElseIf ccmbHistory.Text = "——" Or ccmbHistory.Text = "" Then
        
'        Chk套用模版.Visible = True
        Cmd结论模板.Visible = True
        Frame5.Visible = True
        Frame6.Caption = "体检项目结果填写："
'        Frame6.Height = Frame6.Height - 300
'        cgrdInput.Height = cgrdInput.Height - 300
        cgrdInput.Enabled = True
        cgrdInput.rows = 1
        
        subShowInputGrid Trim(ctxtSingleNo.Text)
        
    End If
    
End Sub

Private Sub ccmdClear_Click()
    cgrdInfoBatch.rows = 1
    TotalPeopleBatch.Caption = 0
End Sub

'2012-07-13 于登淼
'修改之前翁乔添加的移除函数，允许按ctrl键批量移除
Private Sub ccmdRemove_Click()
'''    If cgrdInfoBatch.rows > 1 Then
'''        cgrdInfoBatch.RemoveItem
'''    End If
    Dim i As Integer
    With cgrdInfoBatch
        If .SelectedRows > 0 Then
            For i = .SelectedRows - 1 To 0 Step -1
                .RemoveItem (.SelectedRow(i))
            Next
        End If
    End With
    TotalPeopleBatch.Caption = cgrdInfoBatch.rows - 1
End Sub

Private Sub ccmdSelInfo_Click()
    On Error GoTo errHandler
    Dim lobjTmp, lobjRec As Object
    Dim i As Integer, j As Integer
    Dim lstrWhere As String
    Dim coptIndex As Integer
    
    '每次批量查询前把套用体检结果的标识去掉
    cchk套用体检结果.Value = 0
    For i = 0 To coptClasses.Count - 1
        If coptClasses(i).Value = True Then coptIndex = i
    Next
    'lstrWhere = " and 体检类型='" & coptClasses(coptIndex).Caption & "'"
        
    '组装查询条件
    If cchkDateBatch.Value = 1 Then
        lstrWhere = lstrWhere & " and 体检日期>='" & Format(cdtpDate.Value, "yyyy-mm-dd 00:00:00") & "' and 体检日期<='" & Format(cdtpDate.Value, "yyyy-mm-dd 23:59:59") & "'"
    End If
    If cchkCompanyBatch.Value = 1 Then
        lstrWhere = lstrWhere & " and 单位名称='" & Trim(ctxtQueyCompanyBatch.Text) & "'"
    End If
    
    '2012-07-14 于登淼 ↓
    '更改查询条件，加入8/48小时判断内容。超过修改时间的始终不列入查询结果中。
    '查询数据表和内容发生较大变化，若修改，请留意。

    '将该科室所有已有体检结果人员修改时间重新更新。体检基本信息表中“各科体检状态”由'2'改为'3'的，查询时忽略。
    sub更新可修改结果人员修改状态
    
    Set lobjTmp = CreateObject("职业病体检结果录入.clscommon")
    Set mobjQueryResult = lobjTmp.func获取可修改结论_本科室_体检人员信息(lstrWhere, priDeptName)
    
    sub查询列表显示 coptIndex
    '2012-07-14 于登淼 ↑
    
    Set lobjTmp = Nothing
    Set lobjRec = Nothing
    lstrWhere = ""
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病体检结果录入", "FrmENT_ResultInput", "ccmdQuery_Click", 6666, lstrError, False
End Sub

Private Sub ccmdWork_Click()
    On Error GoTo errHandler
    Dim lobjRec As Object  '单位定位返回的结果记录。
    Dim lobj单位 As Object
    Dim lobj单位信息 As Object
    Dim mstr单位申请编号 As String
    '启动单位定位界面。
    Set lobjRec = pobj业务对象.func单位定位
    '获取定位的单位，显示在“单位名称”录入框中。
    If Not lobjRec Is Nothing Then
        If lobjRec.recordcount > 0 Then
            ctxtcchkWork.Text = IIf(IsNull(lobjRec("单位名称")), "", lobjRec("单位名称"))
'            mstr单位申请编号 = lobjRec!申请编号
            'Set lobj单位 = CreateObject("职业病对象.class1")
            'lobj单位.单位信息申请 = lobjRec!申请编号
            'Set lobj单位信息申请 = lobj单位.单位信息
            
            
            
'            If mstr单位申请编号 <> "" Then
'                '修改：2001-8-23（显示单位属性）。
'                On Error Resume Next
'                'sub显示单位属性 ciptBase, mstr单位申请编号, mobjGUI
'                func获取单位信息 lobjRec!申请编号
'            End If
        End If
    End If
    
    '把焦点回到单位录入框。保存能保存新单位定位信息。
    ctxtcchkWork.SetFocus
    SendKeys vbTab
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病界面部件", "血常规录入", "ccmd单位定位_Click", 6666, lstrError, False
End Sub

Private Sub ccmd查询单位_Click()
    Dim lobjRec As Object                       '单位定位返回的结果记录。
    
    On Error GoTo errHandler
    Set lobjRec = pobj业务对象.func单位定位     '启动单位定位界面。
    
    '获取定位的单位，显示在“单位名称”录入框中。(暂时只显示“单位名称”)
    If Not lobjRec Is Nothing Then
        If lobjRec.recordcount > 0 Then
            ctxtQueyCompanyBatch.Text = IIf(IsNull(lobjRec("单位名称")), "", lobjRec("单位名称"))
        End If
    End If
    'flag名称.Value = 1
    Exit Sub
errHandler:
    'If Err.Number = 0 Then Exit Sub
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病界面部件", "FrmImportExcel", "ccmd单位定位_Click", 6666, lstrError, False
End Sub

'当下拉日历被关闭时发生。
Private Sub cdtpDate_CloseUp()
    ccmdSingleQuery_Click
End Sub

'Private Sub ccmbSheet_Click()
'    On Error Resume Next
'    cgrdPerson.rows = 1
'    cgrdInput.rows = 1
'End Sub

Private Sub ccmdAutoFull_Click(Index As Integer)
    Dim i As Integer
    Select Case Index
        Case 0
            For i = 1 To cgrdInput.rows - 1
                If cgrdInput.TextMatrix(i, 2) = "" Then
                    cgrdInput.TextMatrix(i, 2) = "正常"
                End If
            Next
        Case 1
            For i = 1 To cgrdInput.rows - 1
                cgrdInput.TextMatrix(i, 2) = "正常"
            Next
        Case 2
            subShowInputGrid Trim(ctxtSingleNo.Text)
        Case 3
            For i = 1 To cgrdInput.rows - 1
                cgrdInput.TextMatrix(i, 2) = ""
            Next
    End Select
End Sub

Private Sub ccmdSingleQuery_Click()
 On Error GoTo errHandler
    Dim lobjTmp, lobjRec As Object
    Dim i As Integer, j As Integer
    Dim lstrWhere As String
    Dim coptIndex As Integer
    
    '显示指定体检日期的未下结论的体检人员名单。
    'subShowSingleList
    '组装查询条件
    
    For i = 0 To coptClasses.Count - 1
        If coptClasses(i).Value = True Then coptIndex = i
    Next
    'lstrWhere = " and 体检类型='" & coptClasses(coptIndex).Caption & "'"
    
    If cchkDate.Value = 1 Then
        lstrWhere = lstrWhere & " and 体检日期>='" & Format(cdtpDate.Value, "yyyy-mm-dd 00:00:00") & "' and 体检日期<='" & Format(cdtpDate.Value, "yyyy-mm-dd 23:59:59") & "'"
    End If
    
    '2012-07-24 翁乔 修改：增加筛选条件↓
    '系统编号
    If cchkSingleNo.Value = 1 Then
        lstrWhere = lstrWhere & " and a.系统编号='" & Trim(ctxtcchkNo.Text) & "'"
    End If
    '身份证号
    If cchkCardNo.Value = 1 Then
        lstrWhere = lstrWhere & " and 公民身份号码='" & ctxtcchkCardNo.Text & "'"
    End If
    '名字
    If cchkName.Value = 1 Then
        lstrWhere = lstrWhere & " and 姓名='" & ctxtCheckName.Text & "'"
    End If
    '工作单位
    If cchkWorkUnit.Value = 1 Then
        lstrWhere = lstrWhere & " and 单位名称='" & ctxtcchkWork.Text & "'"
    End If
    '2012-07-24 翁乔 修改：增加筛选条件↑
    
   '2012-07-14 于登淼 ↓
    '更改查询条件，加入8/48小时判断内容。超过修改时间的始终不列入查询结果中。
    '查询数据表和内容发生较大变化，若修改，请留意。

    '将该科室所有已有体检结果人员修改时间重新更新。体检基本信息表中“各科体检状态”由'2'改为'3'的，查询时忽略。
    sub更新可修改结果人员修改状态
    
    Set lobjTmp = CreateObject("职业病体检结果录入.clscommon")
    Set mobjQueryResult = lobjTmp.func获取可修改结论_本科室_体检人员信息(lstrWhere, priDeptName)
    
    sub查询列表显示 coptIndex
    '2012-07-14 于登淼 ↑
    
    Set lobjTmp = Nothing
    Set lobjRec = Nothing
    lstrWhere = ""
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病体检结果录入", "frmURT_ResultInput", "ccmdSingleQuery_Click", 6666, lstrError, False
End Sub

'''Private Sub cgrdInfoBatch_Click()
'''    cgrdInfoBatch.SelectionMode = flexSelectionByRow
'''End Sub

Private Sub cgrdInfoBatch_DblClick()
    Dim indX, indY As Integer
    indX = cgrdInfoBatch.MouseRow
    indY = cgrdInfoBatch.MouseCol
    If indX < 0 Or indY < 0 Then
        Exit Sub
    ElseIf indX >= 0 And indX < cgrdInfoBatch.rows And indY >= 0 And indY < cgrdInfoBatch.cols Then
        ctxt体检条码.Text = cgrdInfoBatch.TextMatrix(indX, 0)
        ctxt体检条码_KeyDown 13, 0
    End If
End Sub

Private Sub cgrdInput_AfterEdit(ByVal row As Long, ByVal col As Long)
    Dim lstr单项结论 As String
    On Error GoTo errHandler
    If row > 0 Then
        lstr单项结论 = pobj业务对象.func获取单项结论(cgrdInput.TextMatrix(row, 0), cgrdInput.TextMatrix(row, 2))
        If lstr单项结论 = "不合格" Then
            '设置颜色。
            cgrdInput.Cell(flexcpBackColor, row, 2, row, 2) = &H8A5AFA
        Else
            '设置颜色。
            cgrdInput.Cell(flexcpBackColor, row, 2, row, 2) = vbWhite
        End If
        '2012-06-21 于登淼 ↓
        '设置当前录入状态(已经录入过，正在修改当前结果)
        If ResultChanged = 1 Then ResultChanged = 2
        '2012-06-21 于登淼 ↑
    End If
    Exit Sub
errHandler:
End Sub

Private Sub cgrdInput_DblClick()
    '修改颜色。
    On Error Resume Next
    If cgrdInput.row > 0 Then
        If cgrdInput.Cell(flexcpBackColor, cgrdInput.row, 2, cgrdInput.row, 2) = &H8A5AFA Then
            cgrdInput.Cell(flexcpBackColor, cgrdInput.row, 2, cgrdInput.row, 2) = vbWhite
        Else
            cgrdInput.Cell(flexcpBackColor, cgrdInput.row, 2, cgrdInput.row, 2) = &H8A5AFA
        End If
        '2012-06-21 于登淼 ↓
        '设置当前录入状态(已经录入过，正在修改当前结果)
        If ResultChanged = 1 Then ResultChanged = 2
        '2012-06-21 于登淼 ↑
    End If
End Sub

Private Sub cgrdInput_KeyDownEdit(ByVal row As Long, ByVal col As Long, KeyCode As Integer, ByVal Shift As Integer)
    On Error GoTo errHandler
    If col = 2 And KeyCode = 13 Then
        '换行。
        If row = cgrdInput.rows - 1 Then
            cgrdInput.row = 1
        Else
            cgrdInput.row = cgrdInput.row + 1
        End If
        cgrdInput.col = 2
    End If
    Exit Sub
errHandler:

End Sub

Private Sub cgrdInput_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If cgrdInput.col = 2 And Button = 1 Then
     '   cgrdInput.EditCell
    End If
End Sub
'(网格录入)
'Private Sub cgrdPerson_AfterEdit(ByVal row As Long, ByVal col As Long)
'    Dim lstr单项结论 As String
'    On Error GoTo errHandler
'    If row > 0 Then
'        lstr单项结论 = pobj业务对象.func获取单项结论(cgrdPerson.TextMatrix(row, 0), cgrdPerson.TextMatrix(row, col))
'        If lstr单项结论 = "不合格" Then
'            '设置颜色。
'            cgrdPerson.Cell(flexcpBackColor, row, col, row, col) = &H8A5AFA
'        Else
'            '设置颜色。
'            cgrdPerson.Cell(flexcpBackColor, row, col, row, col) = vbWhite
'        End If
'    End If
'    Exit Sub
'errHandler:
'End Sub

Private Sub cgrdPerson_BeforeEdit(ByVal row As Long, ByVal col As Long, Cancel As Boolean)
    If row < 1 Or col <= mintFixed Then
        Cancel = True
    End If
    
End Sub

'单击列表 表格内容，右侧表格显示体检项目详情
Private Sub cgrdSingleList_dblClick()
    'If cgrdInput.rows < 2 Then
        If cgrdSingleList.row > 0 Then
            ctxtSingleNo.Text = cgrdSingleList.Cell(flexcpText, cgrdSingleList.row, 0)
            
            '2012-07-15 于登淼 ↓
            '载入信息与界面控制功能不全，故直接调用ctxtsingleno_keydown
            ccmbHistory.Enabled = True
            Cmd结论模板.Visible = True
            Frame5.Visible = True
            Frame6.Caption = "体检项目结果填写："
    '        Frame6.Height = Frame6.Height - 300
    '        cgrdInput.Height = cgrdInput.Height - 300
            cgrdInput.Enabled = True
            cgrdInput.rows = 1
            
            ctxtSingleNo_KeyDown 13, 0
'''            '显示人员信息。
'''            subShowSinglePerson
            '2012-07-15 于登淼 ↑
        End If
    'Else
    '    MsgBox "请先保存当前体检人员信息！"
    'End If
End Sub

Private Sub clblInfo_Click(Index As Integer)
End Sub

'2012-05-11 陶露
'套用已有的体检结论模板 可进行选择
Private Sub Cmd结论模板_Click()
    frmConclusion.lobj科室 = priDeptName
    frmConclusion.lobj科室编号 = priDeptNo
    frmConclusion.lobj医生编号 = um用户编号
    frmConclusion.lobj时间 = Now
    frmConclusion.Show
End Sub
'2012-05-11 陶露

'2012-07-14 于登淼
Private Sub coptClasses_Click(Index As Integer)
    Dim coptIndex As Integer
    coptIndex = Index
    sub查询列表显示 coptIndex
End Sub

Private Sub ctxt体检条码_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lstrNo As String
    Dim i As Integer
    Dim str科室结论 As String
    Dim lcol职业病对象 As Object
    lstrNo = Trim(ctxt体检条码.Text)
    
'''    coptClasses(0).Enabled = False
'''    coptClasses(1).Enabled = False
'''    coptClasses(2).Enabled = False
    
    '检查条码号是否存在
    Dim mlobjRec As Object
    Dim lobjTmp As Object
    Set lobjTmp = CreateObject("职业病体检结果录入.clscommon")
    Set mlobjRec = lobjTmp.func获取体检人员基本信息(lstrNo)
    If mlobjRec.recordcount = 0 Then
        Set mlobjRec = Nothing
        
        '清空当前个人信息
        ctxt体检条码.Enabled = True
        ctxt姓名.Text = ""
        ctxt性别.Text = ""
        ctxt年龄.Text = ""
        ctxt单位名称.Text = ""
        Exit Sub
    End If
    
    LoadPersonalInfoBatch (lstrNo)
        
    Set lobjTmp = CreateObject("职业病体检结果录入.clscommon")
    If lobjTmp.func获取体检人员体检科室信息(lstrNo, priDeptName) Then
        Set lobjTmp = Nothing
       
        Set lcol职业病对象 = CreateObject("职业病对象.clsManageMedicalExam")
        str科室结论 = lcol职业病对象.func返回科室结论(ctxt体检条码.Text, priDeptName)
        ctxtConclun.Text = str科室结论
        
        '一旦确定当前体检人员编号，就不能更改。除非，清空界面内容。
        ctxt体检条码.Enabled = False
        ctxt姓名.Enabled = False
        ctxt性别.Enabled = False
        ctxt年龄.Enabled = False
        ctxt单位名称.Enabled = False '其实单位灰掉了之后，如果有“单位定位”按钮，还是可以改的。
        If ResultChanged <> 3 Then
            ccmdAutoFull(0).Enabled = True
            ccmdAutoFull(1).Enabled = True
            ccmdAutoFull(2).Enabled = False
            ctbMain.Buttons(2).Enabled = True
            ccmdAutoFull(3).Enabled = True
        End If
        
        ctbMain.Buttons(2).Enabled = False
        ctbMain.Buttons(3).Enabled = True
    Else
        Set lobjTmp = Nothing
        MsgBox ("该体检人员没有该科室的体检项目！")
        cgrdInfoBatch.RemoveItem
        subClear
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
        .Add "清空界面(&N)110"
        .Add "保存"
        .Add "批量保存(&S)"
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
    ctbMain.Buttons(2).Enabled = False
    ctbMain.Buttons(3).Enabled = False
    '录入结果时，为该医生的姓名
    LabelDoctor.Caption = LabelDoctor.Caption & " " & um用户名
    
    '初始化体检结果录入网格。
    With cgrdInput
        .rows = 1
        .cols = 6
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
    Set mobj体检医师 = CreateObject("职业病对象.clsMedicalExaminer")
    mobj体检医师.编号 = um用户编号
    
    '2012-06-21 于登淼 ↓
    '更改系统编号固定部分。省疾控新要求中改变系统编号规则。
    '获取系统编号固定部分。
'''    Dim lobj体检 As Object '职业病对象，获取系统编号的固定部分。
'''    Set lobj体检 = CreateObject("职业病对象.clsMedicalExam")
'''    mstr系统编号固定部分 = lobj体检.系统编号固定部分
'''    ctxtSingleNo.Text = mstr系统编号固定部分
    sub获取系统编号固定部分
    '2012-06-21 于登淼 ↑
    
    '体检医师"显示框显示当前用户名。
    ctxtDoctor.Text = um用户名
    cdtpInputDate.Value = Now
    cdtpDateBatch.Value = Now
    cdtpDate.Value = Now
    DTP录入日期.Value = Now
    
    cgrdInput.rows = 1
    ctxtSingleNo.TabIndex = 0
        
    '2012-04-11
    '界面按钮设置
    ccmdAutoFull(0).Enabled = False
    ccmdAutoFull(1).Enabled = False
    ccmdAutoFull(2).Enabled = False
    ccmdAutoFull(3).Enabled = False
    Frame5.Enabled = False
    
    '2012-04-12 陶露
    '界面权限设置
    Dim lobjTmp As Object
    Set lobjTmp = CreateObject("职业病设置.clsPermissionConfigure")
    If lobjTmp.func科室操作权限(um用户编号, "职业病体检_尿常规化验科结果录入_修改") = False Then
        ctbMain.Buttons(2).Visible = False
    End If
    
    '2012-05-22 翁乔 ↓↓↓
    '界面权限设置
    If lobjTmp.func科室操作权限(um用户编号, "职业病体检_尿常规化验科结果录入_批量修改") = False Then
        ctbMain.Buttons(3).Visible = False
    End If
    '2012-05-22 ↑↑↑
    Set lobjTmp = Nothing
    
    '2012-04-11

    '修改：2001-12-29（获取操作记忆值）。
    On Error Resume Next
    Set mobj记忆 = New cls用户操作记忆
    mobj记忆.用户编号 = um用户编号
    mobj记忆.业务名 = "体检管理"
    
    ctabPerson.Tab = 0
    
    '2012-06-21 于登淼 ↓
    '初始化当前录入状态(提前判断有无权限修改，若无，直接赋值为3)
    ResultChanged = IIf(ResultChanged <> 3, -1, 3)
    cchk刷条码_Click
    '2012-06-21 于登淼 ↑
    
    '2012-07-14 于登淼 ↓
    '初始化查询界面，调整查询列表格式。初始化科室基本信息。
    priDeptName = "尿常规化验科"
    priDeptNo = "06"
    priDeptResultName = "尿常规化验科"
    ccmdSingleQuery_Click
    ctabPerson.Tab = 1: ccmdSelInfo_Click: ctabPerson.Tab = 0
    coptClasses_Click (0)
    '2012-07-14 于登淼 ↑
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病体检结果录入", "frmURT_ResultInput", "Form_Load", 6666, lstrError, False
    Exit Sub
    Resume
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    '修改：2002-9-26（杨春）保存操作记忆值。
    'mobj记忆.sub覆盖记忆值 "录入结果时刷条码", IIf(cchk刷条码.Value = 1, "是", "否")
    
    '释放本界面的全局对象。
    Set mobj界面通用对象 = Nothing
    Set mobj体检医师 = Nothing
    Set mobj记忆 = Nothing
    
    '设置标志pblnInUse，表明窗体已不在使用。
    mblnInUse = False

End Sub

Private Sub cchk刷条码_Click()
    If Not cchk刷条码.Visible Then Exit Sub
    If ctxtSingleNo.Enabled = False Then Exit Sub
    
    If ctabPerson.Tab = 0 Then
        ctxtSingleNo.Text = ""
        If cchk刷条码.Value = 0 Then sub获取系统编号固定部分
        ctxtSingleNo.Enabled = True
        ctxtSingleNo.SetFocus
        ctxtSingleNo.SelStart = Len(ctxtSingleNo)
        ctxtSingleNo.SelLength = 0
    Else
        ctxt体检条码.Text = ""
        ctxt体检条码.SetFocus
    End If
End Sub

Private Sub cgrdInput_BeforeEdit(ByVal row As Long, ByVal col As Long, Cancel As Boolean)
    Dim lstrEnum As String   '当前体检结果的枚举来源（以英文逗号或中文逗号隔开）。
    Dim i As Long
    
    On Error GoTo errHandler
    '只有体检结果列可以录入。
    If col <> 2 Then
        Cancel = True
    Else
        '根据最后隐藏列存放的枚举来源设置当前单元的下拉列表。
        lstrEnum = cgrdInput.TextMatrix(cgrdInput.row, 3)
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
    sfsub错误处理 "职业病体检结果录入", "frmURT_ResultInput", "cgrdInput_BeforeEdit", 6666, lstrError, False
    Exit Sub
    Resume
End Sub


'Private Sub ctxtSingleNo_GotFocus()
'    On Error Resume Next
'    '若输入系统编号，设置系统编号的固定部分，方便录入。
'    If ctxtSingleNo.Text = "" Then
'        ctxtSingleNo.Text = mstr系统编号固定部分
'        ctxtSingleNo.SelLength = 0
'        ctxtSingleNo.SelStart = Len(mstr系统编号固定部分)
'        ctbMain.Buttons(1).Enabled = False
'
'    End If
'End Sub

Private Sub ctxtSingleNo_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errHandler
    Dim str科室结论 As String
    Dim lcol职业病对象 As Object
    If KeyCode = 13 And Trim(ctxtSingleNo.Text) <> "" Then
        '显示人员信息。
        subShowSinglePerson
        
        Set lcol职业病对象 = CreateObject("职业病对象.clsManageMedicalExam")
        str科室结论 = lcol职业病对象.func返回科室结论(ctxtSingleNo.Text, priDeptName)
        ctxtConclun.Text = str科室结论
        
        'ctbMain.Buttons(2).Enabled = True
        ctbMain.Buttons(3).Enabled = False
        If cgrdInput.rows > 1 Then ctbMain.Buttons(1).Enabled = True
        '2012-07-16 于登淼 ↓
        '添加控制列表显示函数
        With cgrdInput
            .col = 0
            .Sort = flexSortGenericAscending
            .AutoSize 0, .cols - 1, 0, 0
            .ExplorerBar = flexExSort
            .DataMode = flexDMFree
        End With
        '2012-07-16 于登淼 ↑
    End If
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病体检结果录入", "frmURT_ResultInput", "ctxtSingleNo_KeyDown", 6666, lstrError, False
    Exit Sub
    Resume
End Sub


'功能：根据系统编号，设置体检结果录入网格。

Private Sub subShowInputGrid(ByVal paraSysNo As String)
    Dim lcolInfo As Collection '存放当前系统编号中当前医师可以做得指定属性的体检项目及其结果。
    Dim lobjItem As Variant    'clsFactTestItem,lcolInfo中的元素。
    Dim lstrEnum As String
    Dim i As Long
    Dim j As Long
    On Error GoTo errHandler
    
    '使用优化算法获取体检项目。
    Dim lobjRec As Object
    Dim lobjTemp As Object
    
    '获取指定属性（常规/化验）的体检项目：clsFactTestItem(体检项目编码，体检项目名称，缺省值，枚举来源，体检结果)。
    '若选择所有体检表，则获取所有体检表上可作项目。
    Set lobjRec = mobj体检医师.Func优化的获取本人可作的体检项目(paraSysNo, mstr体检项目属性, priDeptName)
    
    '显示体检项目在cgrdInput中。
    cgrdInput.rows = 1
    
    Set mcol体检项目 = New Collection
    
    If lobjRec.recordcount > 0 Then
        '2012-06-21 于登淼 ↓
        '更改当前录入状态
        If IsNull(lobjRec("单项结论")) Then
            ResultChanged = IIf(ResultChanged <> 3, 0, 3)
        Else
            ResultChanged = IIf(ResultChanged <> 3, 1, 3)
        End If
        '2012-06-21 于登淼 ↑
        
        cgrdInput.rows = lobjRec.recordcount + 1
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

            '(为了进行网格录入，在cgrdperson后面增加列显示体检项目名称)。
'            If ctabPerson.Tab = 1 Then
'                For j = mintFixed + 1 To cgrdPerson.cols - 1
'                    If cgrdPerson.TextMatrix(0, j) = lobjRec!体检项目名称 Then Exit For
'                Next
'                If j = cgrdPerson.cols Then
'                    cgrdPerson.cols = cgrdPerson.cols + 1
'                    cgrdPerson.TextMatrix(0, j) = lobjRec!体检项目名称
'
'                    'cgrdPerson.ColHidden(j) = IIf(cchkGrid.Value = 0, True, False)
'                    If lstrEnum = "" Then
'                        cgrdPerson.ColComboList(j) = ""
'                    Else
'                        cgrdPerson.ColComboList(j) = "|" & lstrEnum
'                    End If
'                End If
'
'                mcol体检项目.Add lobjRec("体检项目编号").Value, lobjRec!体检项目名称
'            End If
            i = i + 1
            lobjRec.MoveNext
        Loop
        '2012-07-16 于登淼 ↓
        '添加控制列表显示函数
        With cgrdInput
            .col = 0
            .Sort = flexSortGenericAscending
            .AutoSize 0, .cols - 1, 0, 0
            .ExplorerBar = flexExSort
            .DataMode = flexDMFree
        End With
        '2012-07-16 于登淼 ↑
        cgrdInput.Select 1, 2, 1, 2
    Else
        cgrdInput.rows = 1
        
        Err.Raise 6666, , "对不起，该体检人员体检表上的所有" & mstr体检项目属性 & "体检项目，你都不可以操作。或许该体检人员所使用的体检表上有没有配置" & mstr体检项目属性 & "项目。请进入业务设置的“体检医师设置”检查你可操作的项目，并进入“体检表设置”检查体检表上配置的项目。"
    End If
    
    Exit Sub
errHandler:
    sfsub错误处理 "职业病体检结果录入", "frmURT_ResultInput", "subShowInputGrid", Err.Number, Err.Description, True
    Exit Sub
    Resume
End Sub

'功能：根据单个处理方式输入的编号显示体检人员信息和体检项目结果；或根据批量处理方式输入的单个编号获取体检人员信息并加入网格中，显示体检项目及其结果。
Private Sub subShowSinglePerson()
    On Error GoTo errHandler
    Dim lobj体检 As Object     '职业病对象。
    Dim lobj体检集 As Object   '体检集对象，用于根据试管编号+日期获取系统编号。
    Dim lobjRec As Object
    
    Dim lstrNo As String       '系统编号或试管编号。
    Dim llngNoType As Long     '编号类型：0 系统编号/1 试管编号。
    Dim lstrSysNo As String    '系统编号。
    Dim i As Long
    
    
    '获取输入的系统编号（或试管编号）。
    lstrNo = Trim(ctxtSingleNo.Text)
    
    If lstrNo <> "" Then
        '创建职业病对象。
        Set lobj体检 = CreateObject("职业病对象.clsMedicalExam")
        
        lobj体检.系统编号 = Trim(ctxtSingleNo.Text)
        
        lstrSysNo = lobj体检.系统编号
        'ctxtSingleNo.Text = lstrSysNo
        
        '清空界面。
        If ctabPerson.Tab = 0 Then
            ctxtName = ""
            ctxtSex = ""
            ctxtAge = ""
            ctxtCompanyName = ""
            cpicPhoto(0).Picture = Nothing
        End If
        
        '判断是否存在。
        If Not lobj体检.是否已存在 Then
            Err.Raise 6666, , "不存在你输入编号的体检人员。请重新输入。"
        End If
        
        '判断是否已下体检结论。
        'If lobj体检.体检状态 = P_ENDED_STATUS Then
        '    Err.Raise 6666, , "你输入编号的体检已被医师确定了体检结论，不允许修改已下了体检结论的体检结果。" & Chr(13) & Chr(10) & "若确实要修改，请下结论的医师进入“体检结论录入”操作界面先取消下结论，再回到此操作界面修改。"
        'End If
        
        '显示人员信息（包括相片）。
        If ctabPerson.Tab = 0 Then
            '单个处理方式。
            With lobj体检.体检人员
                .健康档案编号 = lstrSysNo
                ctxtName = .姓名
                ctxtSex = .性别
                ctxtAge = .年龄
                ctxtCompanyName = .单位名称
                                
                '2012-04-11
                '显示人员信息后不能进行修改
                ctxtSingleNo.Enabled = False
                ctxtName.Enabled = False
                ctxtSex.Enabled = False
                ctxtAge.Enabled = False
                ctxtCompanyName.Enabled = False
                '界面按钮能够进行操作
                ccmdAutoFull(0).Enabled = True
                ccmdAutoFull(1).Enabled = True
                ccmdAutoFull(2).Enabled = True
                ccmdAutoFull(3).Enabled = True
                Frame5.Enabled = True
                '2012-04-11
                
                If llngNoType = 1 Then '系统编号输入方式，需要显示试管编号。
                '    clblInfo(4) = lobj体检.体检单号
                    Label1(8).Caption = "体检单号："
                Else
                    'clblInfo(4) = lobj体检.试管编号
                    'Label1(8).Caption = "试管编号："
                End If
                
                '显示相片。
                If Not .像片 Is Nothing Then
                    cpicPhoto(0).Picture = .像片
                End If
            End With
            
            '显示病人的历年病历。（翁乔；2012-07-31）↓↓↓↓↓↓↓↓↓↓↓↓
            Dim lobjDatecobo As Object
            Set lobjDatecobo = mobj体检医师.func获取体检人员的体检病历(Trim(ctxtSingleNo.Text), "尿常规化验科")
            If Not lobjDatecobo Is Nothing Then
                Label3.Visible = True
                ccmbHistory.Visible = True
                ccmbHistory.Clear
                ccmbHistory.AddItem "——"
                For i = 1 To lobjDatecobo.recordcount
                    ccmbHistory.AddItem Format(lobjDatecobo("填写时间"), "yyyy-mm-dd")
'                    ccmbHistory.AddItem
                    lobjDatecobo.MoveNext
                Next i
            Else
                ccmbHistory.Clear
                ccmbHistory.Enabled = False
                
'                Chk套用模版.Visible = True
                Cmd结论模板.Visible = True
                Frame5.Visible = True
'                Frame6.Height = Frame6.Height - 300
'                cgrdInput.Height = cgrdInput.Height - 300
                cgrdInput.Enabled = True
                cgrdInput.rows = 1
                
            End If
'            ccmbHistory.ListIndex = 0
            
            '显示病人的历年病历。（翁乔；2012-07-31） ↑↑↑↑↑↑↑↑↑↑↑↑
            
            '设置体检结果录入网格。
            subShowInputGrid lstrSysNo
            
            cgrdSingleList.row = 0
            
        Else
            '修改：2001-11-2（杨春）因为只查询指定体检表的体检记录，所以不需判断体检表是否相同。
            
            '批量处理方式，把人员信息加入到cgrdPerson中（注意检查体检表是相同的）。
'            If cgrdPerson.rows = 1 Then
''                '若cgrdPerson中原没有记录，设置mstr体检表名称。
''                mstr体检表名称 = lobj体检.体检表.体检表名
''
''                '设置体检结果录入网格
'                subShowInputGrid lstrSysNo
'            'Else
''                '判断体检人员的体检表名是否一致。
'                '修改：2002-8-14（杨春）体检表可以选择所有。
'                'If ccmbSheet.Text <> "<所有>" Then
'                '    If ccmbSheet.Text <> lobj体检.体检表.体检表名 Then
'                '        Err.Raise 6666, , "你输入编号体检的体检表“" & lobj体检.体检表.体检表名 & "”与指定体检表不一致，不能批量录入体检表不相同的体检结果。"
'                '    End If
'                'End If
'            End If
'
'            '判断该人员是否已在网格中，若不在则可以加入网格。
'            For i = 1 To cgrdPerson.rows - 1
'                If cgrdPerson.TextMatrix(i, 0) = lstrSysNo Then
'                    '已在网格中存在，不再加入。
'                    Exit Sub
'                End If
'            Next
'
'            '把人员添加到体检人员网格中。
'            cgrdPerson.rows = cgrdPerson.rows + 1
'
'            i = cgrdPerson.rows - 1
'            cgrdPerson.TextMatrix(i, 0) = lstrSysNo
'
'            '修改：2002-10-11（杨春）增加显示试管编号。
'            cgrdPerson.TextMatrix(i, 1) = lobj体检.试管编号
'            With lobj体检.体检人员
'                cgrdPerson.TextMatrix(i, 2) = .姓名
'                cgrdPerson.TextMatrix(i, 3) = .性别
'                cgrdPerson.TextMatrix(i, 4) = .单位名称
'                cgrdPerson.TextMatrix(i, 5) = .年龄
'            End With
'
        End If
        
    End If

    ctbMain.Buttons(1).Enabled = True
    ctbMain.Buttons(2).Enabled = True
    cstbMain.Panels(1) = ""
    'cgrdInput.row = 1      '''''2012-07-04 于登淼 临时注释，不明原因，部分人员体检项目载入时出错
    cgrdInput.col = 2
    cgrdInput.SetFocus
    SendKeys ""
    
    '2012-07-03 于登淼 ↓
    '每次读入个人信息时，判断是否超过修改时间。
    '以此控制保存按钮是否可用。
    If pobj业务对象.sub是否在修改时间范围内(ctxtSingleNo.Text, priDeptName, 8) = False Then
        ctbMain.Buttons(2).Enabled = False
    End If
    '2012-07-03 于登淼 ↑
    Exit Sub
errHandler:
    If ctabPerson.Tab = 0 Then
        ctbMain.Buttons(1).Enabled = False
    End If
    sfsub错误处理 "职业病体检结果录入", "frmURT_ResultInput", "subShowSinglePerson", Err.Number, Err.Description, True
    Exit Sub
    Resume
End Sub

'功能：把一批（指定编号范围，或体检日期）的体检人员加入网格，并显示体检项目及其结果。
'Private Sub subShowBatchPerson()
'    Dim lobjRec As Object        '通过业务对象获取的指定范围内可以录入体检结果的体检记录。
'    Dim llngNoType  As Integer   '编号方式：0系统编号/1试管编号。
'    Dim llngStartRow As Long     '当前体检人员网格的最大行+1。
'    Dim llngRow As Long          '当前添加的行。
'    Dim i As Long
'    Dim lobjResult As Object
'
'    On Error GoTo errHandler
'    cstbMain.Panels(1) = "正在获取体检记录，请稍候..."
'
'    '获取编号类型。
'    llngNoType = 0 '系统编号。
'
'
'    '输入的是体检日期。
'    '修改：2001-11-2（增加查询参数：体检表名称）。
'    '修改：2002-8-14（杨春）体检表可以选择所有。
'    Set lobjRec = pobj业务对象.Func获取可修改的体检记录(IIf(coptBatchType(2).Value, ctxtBatchNo, ""), IIf(coptBatchType(2).Value, ctxtBatchNo, ""), IIf(coptBatchType(0).Value, Str(cdtpQueryDate.Value), ""), llngNoType, "", IIf(coptBatchType(1).Value, ctxt单位名称.Text, ""))
'
'    If lobjRec.RecordCount > 0 Then
'
'        lobjRec.Filter = ""
'        If cchkUnEnd(1).Value = 1 And cchkEnd(1).Value = 0 Then
'            lobjRec.Filter = "体检状态<>2"
'        ElseIf cchkUnEnd(1).Value = 0 And cchkEnd(1).Value = 1 Then
'            lobjRec.Filter = "体检状态=2"
'        ElseIf cchkUnEnd(1).Value = 0 And cchkEnd(1).Value = 0 Then
'            lobjRec.Filter = "体检状态=-1"
'        End If
'
'        cgrdPerson.Redraw = False
'        mblnSys = True
'        mintFixed = 6
'        If cgrdPerson.rows = 1 And lobjRec.RecordCount > 0 Then
'            '修改：2001-11-2（杨春）不需要判断体检表名称。
'            '设置体检结果录入网格中。
'            subShowInputGrid lobjRec!系统编号
'        End If
'
'        '显示人员信息到cgrdPerson中（注意检查体检表是相同的）。
'        llngStartRow = cgrdPerson.rows - 1
'        Do While Not lobjRec.EOF
'            '修改：2001-11-2（杨春）不需要判断体检表名称。
'    '        If lobjRec!体检表名称 = mstr体检表名称 Then
'                '判断该人员是否已在网格中，若不在则可以加入网格。
'                For i = 1 To llngStartRow
'                    If cgrdPerson.TextMatrix(i, 0) = lobjRec!系统编号 Then
'                        '已在网格中存在，不再加入。
'                        GoTo LabelContinue
'                    End If
'                Next
'                cgrdPerson.AddItem ""
'                llngRow = cgrdPerson.rows - 1
'                With cgrdPerson
'                    .TextMatrix(llngRow, 0) = lobjRec!系统编号
'
'                    .TextMatrix(llngRow, 1) = IIf(IsNull(lobjRec!试管编号), "", lobjRec!试管编号)
'                    .TextMatrix(llngRow, 2) = IIf(IsNull(lobjRec!姓名), "", lobjRec!姓名)
'                    .TextMatrix(llngRow, 3) = IIf(IsNull(lobjRec!性别), "", lobjRec!性别)
'                    .TextMatrix(llngRow, 4) = IIf(IsNull(lobjRec!单位名称), "", lobjRec!单位名称)
'                    .TextMatrix(llngRow, 5) = IIf(IsNull(lobjRec!年龄), "", lobjRec!年龄)
'                    .TextMatrix(llngRow, 6) = IIf(IsNull(lobjRec!体检单号), "", lobjRec!体检单号)
'
'                    If lobjRec!体检状态 = 2 Then
'                        .Cell(flexcpBackColor, llngRow, 0, llngRow, mintFixed) = cchkEnd(1).BackColor
'                    Else
'                        .Cell(flexcpBackColor, llngRow, 0, llngRow, mintFixed) = cchkUnEnd(1).BackColor
'                    End If
'
'                    '2006-6-19(网格录入）
'                    'If cchkGrid.Value = 1 Then
'                        '获取该人的所有体检结果。
'                        subShowPersonResult llngRow, lobjRec!系统编号
'                    'End If
'                End With
'
'LabelContinue:  lobjRec.MoveNext
'        Loop
'    End If
'
'    If cgrdPerson.rows > 1 Then
'        ccmdRemove.Enabled = True
'        ccmdClear.Enabled = True
'    Else
'        ccmdRemove.Enabled = False
'        ccmdClear.Enabled = False
'    End If
'    cgrdPerson.Redraw = True
'
'    On Error Resume Next
'    cgrdPerson.AutoSize 0, cgrdPerson.cols - 1
'    mblnSys = False
'    cstbMain.Panels(1) = ""
'    Exit Sub
'errHandler:
'    sfsub错误处理 "职业病体检结果录入", "frmURT_ResultInput", "subShowBatchPerson", Err.Number, Err.Description, True
'    mblnSys = False
'    Exit Sub
'    Resume
'End Sub

'Private Sub subShowPersonResult(ByVal paraRow As Long, ByVal para系统编号 As String)
'    Dim i As Long
'    Dim lobjResult As Object
'
'
'    Set lobjResult = pobj业务对象.func获取体检结果(para系统编号)
'    Do While Not lobjResult.EOF
'        For i = mintFixed + 1 To cgrdPerson.cols - 1
'            If cgrdPerson.TextMatrix(0, i) = lobjResult!体检项目名称 Then
'                cgrdPerson.TextMatrix(paraRow, i) = IIf(IIf(IsNull(lobjResult!体检结果), "", lobjResult!体检结果) = "", lobjResult!缺省值, lobjResult!体检结果)
'                '设置颜色。
'                If IIf(IsNull(lobjResult!单项结论), "", lobjResult!单项结论) = "不合格" Then
'                    cgrdPerson.Cell(flexcpBackColor, paraRow, i, paraRow, i) = &H8A5AFA
'                Else
'                    cgrdPerson.Cell(flexcpBackColor, paraRow, i, paraRow, i) = vbWhite
'                End If
'                Exit For
'            End If
'        Next
'        lobjResult.MoveNext
'    Loop
'
'End Sub


Private Sub mobj界面通用对象_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    On Error GoTo errHandler
    Dim lcolNo As Collection     '系统编号集合。
    Dim lcolResult As Collection '体检结果集合，item:[体检项目，体检结果]。
    Dim lcolItem As Collection   '单个体检项目的体检结果：[体检项目，体检结果]。
    Dim lcolDetail As Collection
    Dim lblnNotOver As Boolean
    Dim lcolConclusion As String '单个体检项目的体检结论
    Dim i As Long
    Dim j As Long
    
    Select Case Operate
    Case "清空界面"
        subClear
    Case "批量保存"
        '2012-07-13 于登淼 ↓
        '如果没有体检项目，则直接退出，不保存。
        If cgrdInfoBatch.rows <= 1 Then Exit Sub
        '2012-07-13 于登淼 ↑
        
        '2012-07-15 于登淼 ↓
        '没有录入体检结论时，提示且不保存。
        If Len(Trim(ctxtConclun.Text)) = 0 Then
            MsgBox "你还没有为病人下结论"
            GoTo errHandler
        End If
        '2012-07-15 于登淼 ↑
        
        sub批量保存
        
        '2012-07-15 于登淼 ↓
        '保存完之后，重新进行查询。
        ccmdSingleQuery_Click
        i = ctabPerson.Tab
        ctabPerson.Tab = 1: ccmdSelInfo_Click: ctabPerson.Tab = i
        '2012-07-15 于登淼 ↑
    Case "保存"
        '2012-07-03 于登淼 ↓
        '判断是否在修改时间范围内
        If pobj业务对象.sub是否在修改时间范围内(Trim(ctxtSingleNo.Text), priDeptName, 8) = False Then
            MsgBox ("距上次修改已经超过8小时。请与管理员联系获得修改权限后再继续。")
            Exit Sub
        End If
        '2012-07-03 于登淼 ↑
        
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
'            For i = 1 To cgrdPerson.rows - 1
'                lcolNo.Add cgrdPerson.TextMatrix(i, 0)
'
'                '2006-6-19(网格录入)
'                Set lcolDetail = New Collection
''                'If cchkGrid.Value = 1 Then
''
''                    For j = mintFixed + 1 To cgrdPerson.cols - 1
''                        Set lcolItem = New Collection
''                        lcolItem.Add mcol体检项目(cgrdPerson.TextMatrix(0, j)), "体检项目"
''                        lcolItem.Add cgrdPerson.TextMatrix(i, j), "体检结果"
''
''                        If cgrdPerson.TextMatrix(i, j) = "" Then
''                            lblnNotOver = True
''                            lcolItem.Add "", "单项结论"
''                        ElseIf cgrdPerson.Cell(flexcpBackColor, i, j, i, j) = &H8A5AFA Then
''                            lcolItem.Add "不合格", "单项结论"
''                        Else
''                            lcolItem.Add "合格", "单项结论"
''                        End If
''
''                        lcolDetail.Add lcolItem, lcolItem("体检项目")
''                    Next
''                    lcolResult.Add lcolDetail, cgrdPerson.TextMatrix(i, 0)
''                End If
'            Next
        End If
        
        If lcolNo.Count = 0 Then
            Err.Raise 6666, , "请选择体检人员，并录入体检结果后，再按“保存”。"
        End If
        
        If ctabPerson.Tab = 0 Then
            
            For i = 1 To cgrdInput.rows - 1
                Set lcolItem = New Collection
                lcolItem.Add cgrdInput.TextMatrix(i, 0), "体检项目"
                lcolItem.Add cgrdInput.TextMatrix(i, 2), "体检结果"
                
                '记录没有录完。
                If cgrdInput.TextMatrix(i, 2) = "" Then
                    lblnNotOver = True
                    lcolItem.Add "", "单项结论"
                ElseIf Trim(cgrdInput.TextMatrix(i, 2)) = "异常" Then
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
        
        If Len(Trim(ctxtConclun.Text)) > 0 Then
            lcolConclusion = ctxtConclun.Text
            '保存单个项目的医生结论
            pobj业务对象.sub单个填写体检结论 lcolNo.Item(1), priDeptName, lcolConclusion, um用户编号
            
            '2012-07-03 于登淼 ↓
            '增加一个字段"修改起始时间"的修改。同时修改该科室的体检结果录入状态。
            pobj业务对象.sub修改起始时间 Trim(ctxtSingleNo.Text), priDeptName
            pobj业务对象.sub修改结果录入状态 Trim(ctxtSingleNo.Text), priDeptNo, "2"
            pobj业务对象.sub结果录入修改体检状态 Trim(ctxtSingleNo.Text), "4"
            '2012-07-03 于登淼 ↑
        Else
            MsgBox "你还没有为病人下结论"
            GoTo errHandler
        End If
        
        '使用优化的算法保存体检结果。
        If ctabPerson.Tab = 0 Then
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
        subClear
        
        '2012-07-15 于登淼 ↓
        '保存完之后，重新进行查询。
        ccmdSingleQuery_Click
        i = ctabPerson.Tab
        ctabPerson.Tab = 1: ccmdSelInfo_Click: ctabPerson.Tab = i
        '2012-07-15 于登淼 ↑
        
        MousePointer = 0
        cstbMain.Panels(1) = "保存成功！"
        Cancel = True
    '2012-06-21 于登淼 ↓
    '退出时增加判断是否保存
    Case "退出"
        ctxtSingleNo.Enabled = True
        ctxtSingleNo.SetFocus
        ctxtSingleNo.Enabled = False
        Dim isSave As Integer
        If ResultChanged = 2 Or ResultChanged = 0 Then
            '修改：如果处于病历查看、则退出不提醒。（翁乔，2012-08-01）
            If Trim(Frame6.Caption) <> "体检项目结果填写：" Then
                Unload Me
                Exit Sub
            End If
            isSave = MsgBox("是否保存已修改结果？", vbYesNoCancel)
            If isSave = vbCancel Then Exit Sub
            If isSave = vbYes Then mobj界面通用对象_BeforeOperate "保存", False
        End If
        Unload Me
        Set frmURT_ResultInput = Nothing
    '2012-06-21 于登淼 ↑
    End Select
    Exit Sub
    
errHandler:
    If Err.Number <> 0 Then
        Dim lstrError As String
        lstrError = func错误处理(Err.Number, Err.Description)
        sfsub错误处理 "职业病体检结果录入", "frmURT_ResultInput", "mobj界面通用对象_BeforeOperate", 6666, lstrError, False
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
'弹出手写结论窗口
Private Sub WriteConclun_Click()
    frmConclunInput_Write.Show 1
End Sub

Sub LoadPersonalInfoBatch(ByVal paraSysNo As String)
    On Error GoTo errHandler
    Dim lobj体检 As Object     '职业病对象。
    Dim lobj体检集 As Object   '体检集对象，用于根据试管编号+日期获取系统编号。
    Dim lobjRec As Object
    
    Dim lstrNo As String       '系统编号或试管编号。
    Dim llngNoType As Long     '编号类型：0 系统编号/1 试管编号。
    Dim lstrSysNo As String    '系统编号。
    Dim i As Long
    
    
    '获取输入的系统编号（或试管编号）。
    lstrNo = paraSysNo
    
    If lstrNo <> "" Then
        '创建职业病对象。
        Set lobj体检 = CreateObject("职业病对象.clsMedicalExam")
        
        lobj体检.系统编号 = lstrNo
        
        lstrSysNo = lobj体检.系统编号
        'ctxtSingleNo.Text = lstrSysNo
        
        '清空界面。
        If ctabPerson.Tab = 1 Then
            ctxt姓名 = ""
            ctxt性别 = ""
            ctxt年龄 = ""
            ctxt单位名称 = ""
            cpicPhoto(0).Picture = Nothing
        End If
        
        '判断是否存在。
        If Not lobj体检.是否已存在 Then
            Err.Raise 6666, , "不存在你输入编号的体检人员。请重新输入。"
        End If
        
        '判断是否已下体检结论。
        'If lobj体检.体检状态 = P_ENDED_STATUS Then
        '    Err.Raise 6666, , "你输入编号的体检已被医师确定了体检结论，不允许修改已下了体检结论的体检结果。" & Chr(13) & Chr(10) & "若确实要修改，请下结论的医师进入“体检结论录入”操作界面先取消下结论，再回到此操作界面修改。"
        'End If
        
        '显示人员信息（包括相片）。
        With lobj体检.体检人员
            .健康档案编号 = lstrSysNo
            ctxt姓名 = .姓名
            ctxt性别 = .性别
            ctxt年龄 = .年龄
            ctxt单位名称 = .单位名称
            Picture4.Enabled = True
            Picture4.Visible = True
            Picture4.Picture = .像片
                
            '2012-04-11
            '显示人员信息后不能进行修改
            ctxt体检条码.Enabled = False
            ctxt姓名.Enabled = False
            ctxt性别.Enabled = False
            ctxt年龄.Enabled = False
            ctxt单位名称.Enabled = False
            '界面按钮能够进行操作
            ccmdAutoFull(0).Enabled = True
            ccmdAutoFull(1).Enabled = True
            ccmdAutoFull(2).Enabled = True
            ccmdAutoFull(3).Enabled = True
            Frame5.Enabled = True
            '2012-04-11
                
            If llngNoType = 1 Then '系统编号输入方式，需要显示试管编号。
            '    clblInfo(4) = lobj体检.体检单号
                    Label1(8).Caption = "体检单号："
            Else
                'clblInfo(4) = lobj体检.试管编号
                'Label1(8).Caption = "试管编号："
            End If
                
            '显示相片。
            If Not .像片 Is Nothing Then
                cpicPhoto(0).Picture = .像片
            End If
        End With
            
        '设置体检结果录入网格。
        If cchk套用体检结果.Value = 1 Then
            Exit Sub
        End If
        subShowInputGrid lstrSysNo
            
        cgrdSingleList.row = 0
        
    End If

    ctbMain.Buttons(1).Enabled = True
    'cstbMain.Panels(1) = ""
    cgrdInput.row = 1
    cgrdInput.col = 2
    cgrdInput.SetFocus
    SendKeys ""
    Exit Sub
errHandler:
    If ctabPerson.Tab = 0 Then
        ctbMain.Buttons(1).Enabled = False
    End If
    sfsub错误处理 "职业病体检结果录入", "FrmInMedi_ResultInput", "subShowSinglePerson", Err.Number, Err.Description, True
    Exit Sub
    Resume
End Sub

'功能：清空界面数据
'作者：翁乔
'时间：2012-05-08
Sub subClear()
    TotalPeople.Caption = 0
    TotalPeopleBatch.Caption = 0
    
    '当前界面不可操作
    cgrdInput.rows = 1
            
    cdtpInputDate.Value = Now
    '清空当前个人信息
    ctxtSingleNo.Text = ""
    ctxtSingleNo.Enabled = True
    ctxtName.Text = ""
    ctxtSex.Text = ""
    ctxtAge.Text = ""
    ctxtCompanyName.Text = ""
    cgrdSingleList.rows = 1
    
    '批量信息清除
    DTP录入日期.Value = Now
    ctxt体检条码.Text = ""
    ctxt体检条码.Enabled = True
    ctxt姓名.Text = ""
    ctxt性别.Text = ""
    ctxt年龄.Text = ""
    ctxt单位名称.Text = ""
    cgrdInfoBatch.rows = 1
    '套用信息标志清空
    cchk套用体检结果.Value = 0
    ctxtConclun.Text = ""
'清空照片
    Set cpicPhoto(0).Picture = Nothing
    Set Picture4.Picture = Nothing
'    '清空查询结果（不一定要有的,也没写全）
'    cchkDate.Value = 0
'    cdtpDate.Value = Now
'    cgrdInfo.Clear

    '恢复为form_load时的状态。
    If ctabPerson.Tab = 0 Then
        ctxtSingleNo.Enabled = True
        ctxtSingleNo.SetFocus
        ctxtName.Enabled = True
        ctxtSex.Enabled = True
        ctxtAge.Enabled = True
        ctxtCompanyName.Enabled = True
        ccmdAutoFull(0).Enabled = False
        ccmdAutoFull(1).Enabled = False
        ccmdAutoFull(2).Enabled = False
        ccmdAutoFull(3).Enabled = False
    Else
        ctxt体检条码.Enabled = True
        ctxt姓名.Enabled = True
        ctxt性别.Enabled = True
        ctxt年龄.Enabled = True
        ctxt单位名称.Enabled = True
    End If
    
'''    coptClasses(0).Enabled = True
'''    coptClasses(1).Enabled = True
'''    coptClasses(2).Enabled = True
    ctbMain.Enabled = True
    ctabPerson.Enabled = True
    Frame1.Enabled = True
    Frame2.Enabled = True
    Frame3.Enabled = True
    Frame4.Enabled = True
    Frame5.Enabled = True
    Frame6.Enabled = True
    
    ctbMain.Buttons(1).Enabled = True
    ctbMain.Buttons(2).Enabled = False
    ctbMain.Buttons(3).Enabled = False
    
    '2012-06-21 于登淼 ↓
    '初始化当前录入状态(提前判断有无权限修改，若无，直接赋值为3)
    ResultChanged = IIf(ResultChanged <> 3, -1, 3)
    cchk刷条码_Click
    '2012-06-21 于登淼 ↑
    
End Sub

'功能：批量提交体检人员的体检结果
'时间：2012-04-26
'作者：翁乔
Private Sub sub批量保存()
    MousePointer = 11
    Dim lblnNotOver As Boolean
    Dim i As Integer
    Dim barCode As Collection '批量保存体检条码
        'cstbMain.Panels(1) = "正在保存，请稍候..."
        
        
        '暂时界面不能操作。
        ctbMain.Enabled = False
        ctabPerson.Enabled = False
        Frame1.Enabled = False
        cgrdInput.Select 1, 0, 1, 0
'''        coptClasses(0).Enabled = False
'''        coptClasses(1).Enabled = False
'''        coptClasses(2).Enabled = False

        lblnNotOver = False
        
        
        Set barCode = New Collection
        Set lcolItem = New Collection
        Set lcolResult = New Collection
        '读取批量体检人员的体检条码号
        For i = 1 To cgrdInfoBatch.rows - 1
            barCode.Add cgrdInfoBatch.TextMatrix(i, 0)
        Next i
        '翁乔
        '时间：2012-05-23 ↓↓↓↓
        For i = 1 To cgrdInput.rows - 1
            lcolItem.Add cgrdInput.TextMatrix(i, 1)
            lcolResult.Add cgrdInput.TextMatrix(i, 2)
        Next
        '时间：2012-05-23 ↑↑↑↑
        '若没有录完，进行提示。
        If lblnNotOver Then
            If Not sffuncMsg("你没有录完所有体检项目的体检结果，是否坚持要保存？", sf询问) Then
                '用户选择不保存。
                GoTo errHandler
            End If
        End If

        subSaveBatch
        MousePointer = 0
        'cstbMain.Panels(1) = "保存成功！"
        'Cancel = True
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病体检结果录入", "FrmENT_ResultInput", "ccmdSave_Click", 6666, lstrError, False
    
End Sub

Private Sub subSaveBatch()
    On Error GoTo errHandler
    
    If cgrdInfoBatch.rows < 1 Then
        MsgBox ("请确认录入人员数目是否正确！")
        Exit Sub
    End If
    Dim ccrpValue As Integer
    Dim ccrpI As Integer
    Dim isOk As Boolean
    Dim lstrTmp As String
    Dim lobjTmp As Object
    Dim barCode As Collection
    Dim lcolConclusion As String '内科的体检结论
    Dim i As Integer
    Set barCode = New Collection
    For i = 1 To cgrdInfoBatch.rows - 1
        barCode.Add cgrdInfoBatch.TextMatrix(i, 0)
    Next i
    
    '显示保存进度条
    ccrpI = barCode.Count
    ccrp进度.Max = ccrpI * 2
    ccrp进度.Visible = True
    ccrp进度.Caption = "0%"
    ccrp进度.Value = 0
    
    
    Set lobjTmp = CreateObject("职业病体检结果录入.clsCommon")
    For i = 1 To barCode.Count
        isOk = lobjTmp.func保存单人体检结果(barCode(i), um用户名, DTP录入日期.Value, lcolItem, lcolResult, "职业病体检_结果信息_尿常规化验科")
        ccrp进度.Caption = Int(i / ccrp进度.Max * 100) & "%"
        ccrp进度.Value = ccrp进度.Value + 1
        If i = barCode.Count Then ccrpValue = Int(i / ccrp进度.Max * 100)
    Next i
    
    If ResultChanged <> 3 Then ResultChanged = 1
    If isOk = True Then
        For i = 1 To barCode.Count
            '保存单个项目的医生结论
            lcolConclusion = ctxtConclun.Text
            pobj业务对象.sub单个填写体检结论 barCode(i), priDeptName, lcolConclusion, um用户编号
            ccrp进度.Caption = Int(i / ccrp进度.Max * 100) + ccrpValue & "%"
            ccrp进度.Value = ccrp进度.Value + 1
            
            '2012-07-03 于登淼 ↓
            '增加一个字段"修改起始时间"的修改。同时修改该科室的体检结果录入状态。
            pobj业务对象.sub修改起始时间 barCode(i), priDeptName
            pobj业务对象.sub修改结果录入状态 barCode(i), priDeptNo, "2"
            pobj业务对象.sub结果录入修改体检状态 barCode(i), "4"
            '2012-07-03 于登淼 ↑
        Next i
        MsgBox ("批量保存成功！")
    Else
        subClear
    End If
    
    subClear
    
    ccrp进度.Visible = False
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病体检结果录入", "frmENT_ResultInput", "subSave", 6666, lstrError, False
End Sub

'2012-06-21 于登淼
Sub sub获取系统编号固定部分()
    '获取服务器日期
    Dim lobjRec As Object
    Set lobjRec = dafuncGetData("select getdate()")
    ctxtSingleNo.Text = um防疫站编号 & um服务器代号 & Format(lobjRec(0), "yyyy")
    Set lobjRec = Nothing
End Sub

'2012-07-14 于登淼
Sub sub更新可修改结果人员修改状态()
    Dim lobjRec As Object
    Dim strSQL As String
    Dim canModify As Boolean
    
    strSQL = "select 系统编号,各科体检状态 from 职业病体检_体检基本数据库 where substring(各科体检状态," & priDeptNo & ",1)='1' or substring(各科体检状态," & priDeptNo & ",1)='2'"
    Set lobjRec = dafuncGetData(strSQL)
    If lobjRec.recordcount = 0 Then Exit Sub
    lobjRec.movefirst
    While lobjRec.EOF <> True
        canModify = pobj业务对象.sub是否在修改时间范围内(lobjRec("系统编号"), priDeptName, 8)
        If canModify = False Then Call pobj业务对象.sub修改结果录入状态(lobjRec("系统编号"), priDeptNo, "3")
        lobjRec.MoveNext
    Wend
End Sub

'2012-07-14 于登淼
Sub sub查询列表显示(ByVal coptIndex As Integer)
    mobjQueryResult.Filter = ""
    
    If mobjQueryResult.recordcount > 0 Then
    
        If ctabPerson.Tab = 0 Then
            If cchkSigResult(0).Value = 1 And cchkSigResult(1).Value = 0 Then
                mobjQueryResult.Filter = "填写时间<>null"
            ElseIf cchkSigResult(0).Value = 0 And cchkSigResult(1).Value = 1 Then
                mobjQueryResult.Filter = "填写时间=null"
            ElseIf cchkSigResult(0).Value = 0 And cchkSigResult(1).Value = 0 Then
                mobjQueryResult.Filter = "系统编号='xxx'"
            Else
                mobjQueryResult.Filter = ""
            End If
        ElseIf ctabPerson.Tab = 1 Then
            If cchkBchResult(0).Value = 1 And cchkBchResult(1).Value = 0 Then
                mobjQueryResult.Filter = "填写时间<>null"
            ElseIf cchkBchResult(0).Value = 0 And cchkBchResult(1).Value = 1 Then
                mobjQueryResult.Filter = "填写时间=null"
            ElseIf cchkBchResult(0).Value = 0 And cchkBchResult(1).Value = 0 Then
                mobjQueryResult.Filter = "系统编号='xxx'"
            Else
                mobjQueryResult.Filter = ""
            End If
        End If
        
        If mobjQueryResult.Filter <> "" And mobjQueryResult.Filter <> 0 And mobjQueryResult.Filter <> "系统编号='xxx'" Then
            mobjQueryResult.Filter = mobjQueryResult.Filter & " and 体检类型='" & coptClasses(coptIndex).Caption & "'"
        Else
            mobjQueryResult.Filter = "体检类型='" & coptClasses(coptIndex).Caption & "'"
        End If
        
    End If 'mobjQueryResult.recordcount = 0
    
    If ctabPerson.Tab = 0 Then
        With cgrdSingleList
            Set .DataSource = mobjQueryResult
            .col = 0
            .Sort = flexSortGenericDescending
            .AutoSize 0, .cols - 1, 0, 0
            .ExplorerBar = flexExSort
            .DataMode = flexDMFree
            .AllowSelection = True
            .AllowBigSelection = False
            .SelectionMode = flexSelectionByRow
        End With
        TotalPeople.Caption = IIf(mobjQueryResult.recordcount = 0, "0", mobjQueryResult.recordcount)
    Else
        With cgrdInfoBatch
            Set .DataSource = mobjQueryResult
            .col = 0
            .Sort = flexSortGenericDescending
            .AutoSize 0, .cols - 1, 0, 0
            .ExplorerBar = flexExSort
            .DataMode = flexDMFree
            .AllowSelection = True
            .AllowBigSelection = True
            .SelectionMode = flexSelectionListBox
        End With
        TotalPeopleBatch.Caption = IIf(mobjQueryResult.recordcount = 0, "0", mobjQueryResult.recordcount)
    End If
    cgrdInput.rows = 1

End Sub
