VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C8530F8A-C19C-11D2-99D6-9419F37DBB29}#1.1#0"; "ccrpprg6.ocx"
Begin VB.Form frmHEENT_ResultInput 
   Caption         =   "五官科结果录入窗口"
   ClientHeight    =   9420
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14430
   LinkTopic       =   "Form1"
   ScaleHeight     =   9420
   ScaleWidth      =   14430
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox Picture1 
      Height          =   9255
      Left            =   0
      ScaleHeight     =   9195
      ScaleWidth      =   14235
      TabIndex        =   0
      Top             =   0
      Width           =   14295
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   9135
         Left            =   0
         TabIndex        =   1
         Top             =   120
         Width           =   14175
         Begin VB.CheckBox cchk刷条码 
            BackColor       =   &H00C0FFFF&
            Caption         =   "刷条码"
            Height          =   180
            Left            =   2160
            TabIndex        =   86
            Top             =   1200
            Width           =   975
         End
         Begin VB.OptionButton coptClasses 
            BackColor       =   &H00C0FFFF&
            Caption         =   "放射健康"
            Height          =   255
            Index           =   2
            Left            =   2400
            TabIndex        =   84
            Top             =   840
            Width           =   1095
         End
         Begin VB.OptionButton coptClasses 
            BackColor       =   &H00C0FFFF&
            Caption         =   "职业健康"
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   83
            Top             =   840
            Width           =   1095
         End
         Begin VB.OptionButton coptClasses 
            BackColor       =   &H00C0FFFF&
            Caption         =   "普通体检"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   82
            Top             =   840
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton coptClasses 
            BackColor       =   &H00C0FFFF&
            Caption         =   "涉核部队"
            Height          =   255
            Index           =   3
            Left            =   3480
            TabIndex        =   81
            Top             =   840
            Width           =   1095
         End
         Begin VB.OptionButton coptClasses 
            BackColor       =   &H00C0FFFF&
            Caption         =   "8023部队"
            Height          =   255
            Index           =   4
            Left            =   4560
            TabIndex        =   80
            Top             =   840
            Width           =   1095
         End
         Begin TabDlg.SSTab SSTResultIn 
            Height          =   8175
            Left            =   6000
            TabIndex        =   2
            Top             =   840
            Width           =   8055
            _ExtentX        =   14208
            _ExtentY        =   14420
            _Version        =   393216
            Tabs            =   2
            TabHeight       =   520
            TabCaption(0)   =   "眼科"
            TabPicture(0)   =   "frmHEENT_ResultInput.frx":0000
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "fraDraw"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "Frame3"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).ControlCount=   2
            TabCaption(1)   =   "耳鼻喉科"
            TabPicture(1)   =   "frmHEENT_ResultInput.frx":001C
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Frame2"
            Tab(1).Control(1)=   "ctxtConclun"
            Tab(1).Control(2)=   "Cmd结论模板"
            Tab(1).Control(3)=   "Chk套用模版"
            Tab(1).ControlCount=   4
            Begin VB.CheckBox Chk套用模版 
               Caption         =   "套用模版"
               Height          =   255
               Left            =   -69360
               TabIndex        =   55
               Top             =   5280
               Value           =   1  'Checked
               Width           =   1095
            End
            Begin VB.CommandButton Cmd结论模板 
               Caption         =   "结论模板"
               Height          =   495
               Left            =   -68160
               TabIndex        =   54
               Top             =   5040
               Width           =   1095
            End
            Begin VB.TextBox ctxtConclun 
               Height          =   2415
               Left            =   -74880
               MultiLine       =   -1  'True
               TabIndex        =   22
               Top             =   5640
               Width           =   7815
            End
            Begin VB.Frame Frame2 
               Height          =   3855
               Left            =   -74880
               TabIndex        =   20
               Top             =   480
               Width           =   7935
               Begin VSFlex8Ctl.VSFlexGrid ResultEar 
                  Height          =   3615
                  Left            =   120
                  TabIndex        =   21
                  Top             =   120
                  Width           =   7335
                  _cx             =   2088776330
                  _cy             =   2088769768
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
                  SelectionMode   =   0
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   19
                  Cols            =   3
                  FixedRows       =   1
                  FixedCols       =   2
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
            Begin VB.Frame Frame3 
               Height          =   3975
               Left            =   120
               TabIndex        =   17
               Top             =   360
               Width           =   7815
               Begin VSFlex8Ctl.VSFlexGrid ResultEye 
                  Height          =   2055
                  Left            =   120
                  TabIndex        =   18
                  ToolTipText     =   "单击可修改"
                  Top             =   240
                  Width           =   3735
                  _cx             =   2088769980
                  _cy             =   2088767017
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
                  SelectionMode   =   0
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   16
                  Cols            =   11
                  FixedRows       =   2
                  FixedCols       =   2
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   ""
                  ScrollTrack     =   -1  'True
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
               Begin VSFlex8Ctl.VSFlexGrid ResultEyeArmy 
                  Height          =   1815
                  Left            =   120
                  TabIndex        =   19
                  Top             =   1920
                  Width           =   3735
                  _cx             =   2088769980
                  _cy             =   2088766593
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
                  SelectionMode   =   0
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   13
                  Cols            =   9
                  FixedRows       =   2
                  FixedCols       =   2
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
            Begin VB.Frame fraDraw 
               Caption         =   "晶体环面及正面图(画笔颜色默认红色)"
               Height          =   3615
               Left            =   120
               TabIndex        =   3
               Top             =   4440
               Width           =   7815
               Begin VB.CommandButton ccmdLoadOriginalPicture 
                  Caption         =   "载入原图"
                  Height          =   375
                  Left            =   5280
                  TabIndex        =   9
                  Top             =   720
                  Width           =   975
               End
               Begin VB.PictureBox Picture3 
                  AutoRedraw      =   -1  'True
                  AutoSize        =   -1  'True
                  Height          =   2220
                  Left            =   120
                  ScaleHeight     =   2160
                  ScaleWidth      =   7515
                  TabIndex        =   8
                  Top             =   1200
                  Width           =   7575
               End
               Begin VB.CommandButton ccmdSavePicture 
                  Caption         =   "保存图像"
                  Height          =   255
                  Left            =   2160
                  TabIndex        =   7
                  Top             =   840
                  Visible         =   0   'False
                  Width           =   1215
               End
               Begin VB.CommandButton ccmdClearPicture 
                  Caption         =   "清空此次修改"
                  Height          =   375
                  Left            =   5280
                  TabIndex        =   6
                  Top             =   240
                  Width           =   1335
               End
               Begin VB.CommandButton ccmdEraser 
                  Caption         =   "橡皮擦"
                  Height          =   375
                  Left            =   4200
                  TabIndex        =   5
                  Top             =   720
                  Width           =   855
               End
               Begin VB.CommandButton ccmdDraw 
                  Caption         =   "画图"
                  Height          =   375
                  Left            =   4200
                  TabIndex        =   4
                  Top             =   240
                  Width           =   855
               End
               Begin VB.Label Label9 
                  BackColor       =   &H00FFC0C0&
                  Caption         =   "最粗"
                  Height          =   375
                  Index           =   3
                  Left            =   3600
                  TabIndex        =   16
                  Top             =   360
                  Width           =   495
               End
               Begin VB.Label Label9 
                  BackColor       =   &H00FFC0C0&
                  Caption         =   "粗"
                  Height          =   375
                  Index           =   2
                  Left            =   3000
                  TabIndex        =   15
                  Top             =   360
                  Width           =   495
               End
               Begin VB.Label Label9 
                  BackColor       =   &H00FFC0C0&
                  Caption         =   "中"
                  Height          =   375
                  Index           =   1
                  Left            =   2400
                  TabIndex        =   14
                  Top             =   360
                  Width           =   495
               End
               Begin VB.Label Label13 
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "左眼"
                  Height          =   255
                  Left            =   6360
                  TabIndex        =   13
                  Top             =   840
                  Width           =   375
               End
               Begin VB.Label Label12 
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "右眼"
                  Height          =   255
                  Left            =   960
                  TabIndex        =   12
                  Top             =   840
                  Width           =   375
               End
               Begin VB.Label Label9 
                  BackColor       =   &H00FFC0C0&
                  Caption         =   "细"
                  Height          =   375
                  Index           =   0
                  Left            =   1800
                  TabIndex        =   11
                  Top             =   360
                  Width           =   495
               End
               Begin VB.Label Label8 
                  BackColor       =   &H00FFC0C0&
                  Caption         =   "线条/橡皮擦粗细："
                  Height          =   375
                  Left            =   120
                  TabIndex        =   10
                  Top             =   360
                  Width           =   1575
               End
            End
            Begin VB.Label Label10 
               Caption         =   "五官科结论："
               Height          =   255
               Left            =   -74760
               TabIndex        =   23
               Top             =   5040
               Width           =   1335
            End
         End
         Begin TabDlg.SSTab SSTPersonalInfo 
            Height          =   7455
            Left            =   240
            TabIndex        =   24
            Top             =   1440
            Width           =   5775
            _ExtentX        =   10186
            _ExtentY        =   13150
            _Version        =   393216
            Tabs            =   2
            Tab             =   1
            TabHeight       =   520
            TabCaption(0)   =   "单个录入"
            TabPicture(0)   =   "frmHEENT_ResultInput.frx":0038
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "fraInfo"
            Tab(0).Control(1)=   "fraQuery"
            Tab(0).ControlCount=   2
            TabCaption(1)   =   "批量录入"
            TabPicture(1)   =   "frmHEENT_ResultInput.frx":0054
            Tab(1).ControlEnabled=   -1  'True
            Tab(1).Control(0)=   "TotalPeopleBatch"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).Control(1)=   "Label6"
            Tab(1).Control(1).Enabled=   0   'False
            Tab(1).Control(2)=   "cdtpDateBatch"
            Tab(1).Control(2).Enabled=   0   'False
            Tab(1).Control(3)=   "ccmdRemove"
            Tab(1).Control(3).Enabled=   0   'False
            Tab(1).Control(4)=   "ccmdClear"
            Tab(1).Control(4).Enabled=   0   'False
            Tab(1).Control(5)=   "cchkDateBatch"
            Tab(1).Control(5).Enabled=   0   'False
            Tab(1).Control(6)=   "fraQueryBatch"
            Tab(1).Control(6).Enabled=   0   'False
            Tab(1).Control(7)=   "ccmd查询单位"
            Tab(1).Control(7).Enabled=   0   'False
            Tab(1).Control(8)=   "ctxtQueyCompanyBatch"
            Tab(1).Control(8).Enabled=   0   'False
            Tab(1).Control(9)=   "cchkCompanyBatch"
            Tab(1).Control(9).Enabled=   0   'False
            Tab(1).Control(10)=   "ccmdSelInfo"
            Tab(1).Control(10).Enabled=   0   'False
            Tab(1).Control(11)=   "cgrdInfoBatch"
            Tab(1).Control(11).Enabled=   0   'False
            Tab(1).Control(12)=   "Timerccrp"
            Tab(1).Control(12).Enabled=   0   'False
            Tab(1).Control(13)=   "ccrp进度"
            Tab(1).Control(13).Enabled=   0   'False
            Tab(1).Control(14)=   "cchkBchResult(0)"
            Tab(1).Control(14).Enabled=   0   'False
            Tab(1).Control(15)=   "cchkBchResult(1)"
            Tab(1).Control(15).Enabled=   0   'False
            Tab(1).ControlCount=   16
            Begin VB.CheckBox cchkBchResult 
               Caption         =   "未填结果"
               Height          =   180
               Index           =   1
               Left            =   2760
               TabIndex        =   92
               Top             =   4320
               Value           =   1  'Checked
               Width           =   1095
            End
            Begin VB.CheckBox cchkBchResult 
               Caption         =   "已填结果"
               Height          =   180
               Index           =   0
               Left            =   1440
               TabIndex        =   91
               Top             =   4320
               Width           =   1095
            End
            Begin CCRProgressBar6.ccrpProgressBar ccrp进度 
               Height          =   375
               Left            =   120
               Top             =   4560
               Visible         =   0   'False
               Width           =   5535
               _ExtentX        =   9763
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
               ForeColor       =   255
            End
            Begin VB.Timer Timerccrp 
               Left            =   4560
               Top             =   4560
            End
            Begin VSFlex8Ctl.VSFlexGrid cgrdInfoBatch 
               Height          =   2295
               Left            =   120
               TabIndex        =   76
               Top             =   5160
               Width           =   5535
               _cx             =   2088773155
               _cy             =   2088767440
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
            Begin VB.CommandButton ccmdSelInfo 
               Caption         =   "查 询"
               Height          =   375
               Left            =   4920
               TabIndex        =   75
               Top             =   3960
               Width           =   735
            End
            Begin VB.CheckBox cchkCompanyBatch 
               BackColor       =   &H00C0FFC0&
               Caption         =   "单位名称"
               Height          =   255
               Left            =   120
               TabIndex        =   74
               Top             =   3960
               Width           =   1215
            End
            Begin VB.TextBox ctxtQueyCompanyBatch 
               Height          =   300
               Left            =   1440
               TabIndex        =   73
               Top             =   3960
               Width           =   2415
            End
            Begin VB.CommandButton ccmd查询单位 
               Caption         =   "单位定位"
               Height          =   375
               Left            =   3960
               TabIndex        =   72
               Top             =   3960
               Width           =   855
            End
            Begin VB.Frame fraQueryBatch 
               Caption         =   "批量查询体检人员"
               Height          =   2895
               Left            =   120
               TabIndex        =   58
               Top             =   360
               Width           =   5535
               Begin VB.CheckBox cchk套用体检结果 
                  BackColor       =   &H008080FF&
                  Caption         =   "该体检人员结果作为批量体检结果录入"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   77
                  Top             =   2520
                  Value           =   1  'Checked
                  Width           =   3615
               End
               Begin VB.TextBox ctxt体检条码 
                  Height          =   300
                  Left            =   1320
                  TabIndex        =   64
                  Top             =   720
                  Width           =   2415
               End
               Begin VB.TextBox ctxt姓名 
                  Height          =   300
                  Left            =   1320
                  TabIndex        =   63
                  Top             =   1080
                  Width           =   2415
               End
               Begin VB.TextBox ctxt性别 
                  Height          =   300
                  Left            =   1320
                  TabIndex        =   62
                  Top             =   1440
                  Width           =   2415
               End
               Begin VB.TextBox ctxt年龄 
                  Height          =   300
                  Left            =   1320
                  TabIndex        =   61
                  Top             =   1800
                  Width           =   2415
               End
               Begin VB.TextBox ctxt单位名称 
                  Height          =   300
                  Left            =   1320
                  TabIndex        =   60
                  Top             =   2160
                  Width           =   2415
               End
               Begin VB.PictureBox Picture4 
                  Height          =   1935
                  Left            =   3840
                  ScaleHeight     =   1875
                  ScaleWidth      =   1515
                  TabIndex        =   59
                  Top             =   240
                  Width           =   1575
               End
               Begin MSComCtl2.DTPicker DTP录入日期 
                  Height          =   300
                  Left            =   1320
                  TabIndex        =   65
                  Top             =   360
                  Width           =   2415
                  _ExtentX        =   4260
                  _ExtentY        =   529
                  _Version        =   393216
                  Enabled         =   0   'False
                  Format          =   64225280
                  CurrentDate     =   40969
               End
               Begin VB.Label Label18 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "体检条码号"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   71
                  Top             =   720
                  Width           =   975
               End
               Begin VB.Label Label17 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "姓名"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   70
                  Top             =   1080
                  Width           =   975
               End
               Begin VB.Label Label16 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "性别"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   69
                  Top             =   1440
                  Width           =   975
               End
               Begin VB.Label Label15 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "年龄"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   68
                  Top             =   1800
                  Width           =   975
               End
               Begin VB.Label Label14 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "单位名称"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   67
                  Top             =   2160
                  Width           =   975
               End
               Begin VB.Label Label11 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "结论录入日期"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   66
                  Top             =   360
                  Width           =   1095
               End
            End
            Begin VB.CheckBox cchkDateBatch 
               BackColor       =   &H00C0FFC0&
               Caption         =   "体检日期"
               Height          =   255
               Left            =   120
               TabIndex        =   56
               Top             =   3480
               Width           =   1215
            End
            Begin VB.CommandButton ccmdClear 
               Caption         =   "清 空"
               Height          =   375
               Left            =   3960
               TabIndex        =   48
               Top             =   3360
               Width           =   855
            End
            Begin VB.CommandButton ccmdRemove 
               Caption         =   "移 除"
               Height          =   375
               Left            =   4920
               TabIndex        =   47
               Top             =   3360
               Width           =   735
            End
            Begin VB.Frame fraQuery 
               Caption         =   "查询体检人员"
               Height          =   4455
               Left            =   -74880
               TabIndex        =   40
               Top             =   3000
               Width           =   5535
               Begin VB.CheckBox cchkSigResult 
                  Caption         =   "未填结果"
                  Height          =   255
                  Index           =   1
                  Left            =   1680
                  TabIndex        =   90
                  Top             =   720
                  Value           =   1  'Checked
                  Width           =   1095
               End
               Begin VB.CheckBox cchkSigResult 
                  Caption         =   "已填结果"
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   89
                  Top             =   720
                  Width           =   1095
               End
               Begin VB.CommandButton ccmdQuery 
                  Caption         =   "查   询"
                  Height          =   375
                  Left            =   3120
                  TabIndex        =   44
                  Top             =   720
                  Width           =   1095
               End
               Begin VB.TextBox ctxtCheckName 
                  Height          =   270
                  Left            =   3840
                  TabIndex        =   43
                  Top             =   240
                  Width           =   1335
               End
               Begin VB.CheckBox cchkName 
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "姓名"
                  Height          =   255
                  Left            =   3120
                  TabIndex        =   42
                  Top             =   240
                  Width           =   735
               End
               Begin VB.CheckBox cchkDate 
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "体检日期"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   41
                  Top             =   240
                  Width           =   1095
               End
               Begin MSComCtl2.DTPicker cdtpDate 
                  Height          =   255
                  Left            =   1200
                  TabIndex        =   45
                  Top             =   240
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   450
                  _Version        =   393216
                  Format          =   64225280
                  CurrentDate     =   40969
               End
               Begin VSFlex8Ctl.VSFlexGrid cgrdinfo 
                  Height          =   3135
                  Left            =   120
                  TabIndex        =   46
                  ToolTipText     =   "双击自动填入个人信息和已有体检结果"
                  Top             =   1200
                  Width           =   5295
                  _cx             =   2088772732
                  _cy             =   2088768922
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
                  SelectionMode   =   0
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
               Begin VB.Label TotalPeople 
                  AutoSize        =   -1  'True
                  Caption         =   "0"
                  Height          =   180
                  Left            =   5160
                  TabIndex        =   88
                  Top             =   840
                  Width           =   90
               End
               Begin VB.Label Label19 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "人数："
                  Height          =   180
                  Left            =   4560
                  TabIndex        =   87
                  Top             =   840
                  Width           =   540
               End
            End
            Begin VB.Frame fraInfo 
               Caption         =   "个人信息"
               Height          =   2535
               Left            =   -74880
               TabIndex        =   25
               Top             =   360
               Width           =   5535
               Begin VB.CommandButton ccmdLocate 
                  Caption         =   "单位定位"
                  Height          =   255
                  Left            =   4080
                  TabIndex        =   32
                  Top             =   2160
                  Visible         =   0   'False
                  Width           =   975
               End
               Begin VB.TextBox ctxtCompanyName 
                  Height          =   270
                  Left            =   1320
                  TabIndex        =   31
                  Top             =   2160
                  Width           =   2415
               End
               Begin VB.TextBox ctxtAge 
                  Height          =   270
                  Left            =   1320
                  TabIndex        =   30
                  Top             =   1800
                  Width           =   2415
               End
               Begin VB.TextBox ctxtSex 
                  Height          =   270
                  Left            =   1320
                  TabIndex        =   29
                  Top             =   1440
                  Width           =   2415
               End
               Begin VB.TextBox ctxtName 
                  Height          =   270
                  Left            =   1320
                  TabIndex        =   28
                  Top             =   1080
                  Width           =   2415
               End
               Begin VB.TextBox ctxtBarCode 
                  Height          =   270
                  Left            =   1320
                  TabIndex        =   27
                  Top             =   720
                  Width           =   2415
               End
               Begin VB.PictureBox Picture2 
                  Height          =   1935
                  Left            =   3840
                  ScaleHeight     =   1875
                  ScaleWidth      =   1515
                  TabIndex        =   26
                  Top             =   120
                  Width           =   1575
               End
               Begin MSComCtl2.DTPicker cdtpConclusionDate 
                  Height          =   255
                  Left            =   1320
                  TabIndex        =   33
                  Top             =   360
                  Width           =   2415
                  _ExtentX        =   4260
                  _ExtentY        =   450
                  _Version        =   393216
                  Enabled         =   0   'False
                  Format          =   64225280
                  CurrentDate     =   40969
               End
               Begin VB.Label Label7 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "结论录入日期"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   39
                  Top             =   360
                  Width           =   1095
               End
               Begin VB.Label Label5 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "单位名称"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   38
                  Top             =   2160
                  Width           =   975
               End
               Begin VB.Label Label4 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "年龄"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   37
                  Top             =   1800
                  Width           =   975
               End
               Begin VB.Label Label3 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "性别"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   36
                  Top             =   1440
                  Width           =   975
               End
               Begin VB.Label Label2 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "姓名"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   35
                  Top             =   1080
                  Width           =   975
               End
               Begin VB.Label Label1 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "体检条码号"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   34
                  Top             =   720
                  Width           =   975
               End
            End
            Begin MSComCtl2.DTPicker cdtpDateBatch 
               Height          =   300
               Left            =   1440
               TabIndex        =   57
               Top             =   3480
               Width           =   2415
               _ExtentX        =   4260
               _ExtentY        =   529
               _Version        =   393216
               Format          =   64225280
               CurrentDate     =   40969
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0E0FF&
               Caption         =   "人数："
               Height          =   180
               Left            =   120
               TabIndex        =   79
               Top             =   4320
               Width           =   540
            End
            Begin VB.Label TotalPeopleBatch 
               AutoSize        =   -1  'True
               Caption         =   "0"
               Height          =   180
               Left            =   720
               TabIndex        =   78
               Top             =   4320
               Width           =   90
            End
         End
         Begin MSComctlLib.Toolbar ctlb工具栏 
            Height          =   540
            Left            =   120
            TabIndex        =   49
            Top             =   0
            Width           =   13950
            _ExtentX        =   24606
            _ExtentY        =   953
            ButtonWidth     =   1455
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
            Begin VB.CommandButton ccmdUnfilledAllPass 
               Caption         =   "当前未填写项全部正常"
               Height          =   375
               Left            =   6000
               TabIndex        =   53
               Top             =   120
               Width           =   2055
            End
            Begin VB.CommandButton ccmdAutoFill 
               Caption         =   "全部正常"
               Height          =   375
               Left            =   12720
               TabIndex        =   52
               Top             =   120
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.CommandButton ccmdClearResult 
               Caption         =   "清空当前表格"
               Height          =   375
               Left            =   10200
               TabIndex        =   51
               Top             =   120
               Width           =   1455
            End
            Begin VB.CommandButton ccmdSave 
               Caption         =   "保存全部结果"
               Height          =   375
               Left            =   8280
               TabIndex        =   50
               Top             =   120
               Width           =   1695
            End
         End
         Begin MSComDlg.CommonDialog ccmdFile 
            Left            =   7920
            Top             =   720
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            Flags           =   6148
         End
         Begin VB.Label LabelDoctor 
            BackColor       =   &H00C0FFFF&
            Caption         =   "医生："
            Height          =   255
            Left            =   4080
            TabIndex        =   85
            Top             =   1200
            Width           =   1575
         End
      End
   End
End
Attribute VB_Name = "frmHEENT_ResultInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'2012-03-01 于登淼
'增加 五官科结果录入窗体，及相应部件功能

Option Explicit
Public mblnInUse As Boolean
Private WithEvents mobjGUI As cls界面通用对象
Attribute mobjGUI.VB_VarHelpID = -1
Private mstr体检单号 As String
Private mstr系统编号 As String
Private mlobjRec As Object

'查询结果
Private mstrDoctorName As String
Private mobjQueryResult As Object
Private mcolIndex As New Collection
Private indX, indY As Integer       '记录鼠标点击vsflexgrid的坐标。
Private lcolResult As Collection    '体检结果集合，item:[体检项目名称，体检结果]。
Private lcolItem As Collection      '单个体检项目的体检结果：[体检项目名称，体检结果]。

'2012-07-14 于登淼 ↓
'增加科室基本信息变量
Private priDeptName As String
Private priDeptNo As String
Private priDeptResultName As String
'2012-07-14 于登淼 ↑

'记录在第一次保存体检结果之后，如果再次修改结果，需要弹出“结果已修改，是否保存”之类的提示。
'-1，表示未获取该人数据库里体检结果的信息；
'0，表示该人的结果未录入过；
'1，表示数据库里已有该人的结果，但在界面上未被修改过；
'2，表示数据库里已有该人的结果，界面上已修改过。只有在为2的时候，才会弹出“结果已修改，是否保存”窗口
'3，表示没有权限进行修改操作。
Private ResultChanged As Integer

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private DrawLineWidth As Integer        '画图线条粗细的标记
Private DrawState As Integer            '记录当前画图状态，未做修改不保存 -1、未画图0、画图1、橡皮擦2
Private mstr结果图片项目编号 As String  '标记五官科结果图片项目编号，数据库中记录的值为”01069“

Private lobj批量操作对象 As Object    '为批量操作提供对象函数
Private EyeMapCheck(6050, 2) As Integer
Private pointCnt As Long

'功能：返回当前窗体是否已经加载标志。这是系统平台所要求的。
Public Property Get pblnInUse() As Boolean
    On Error Resume Next
    pblnInUse = mblnInUse
End Property


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

Private Sub cchk刷条码_Click()
    If Not cchk刷条码.Visible Then Exit Sub
    
    If SSTPersonalInfo.Tab = 0 Then
        ctxtBarCode.Text = ""
        If cchk刷条码.Value = 0 Then sub获取系统编号固定部分
        ctxtBarCode.Enabled = True
        ctxtBarCode.SetFocus
        ctxtBarCode.SelStart = Len(ctxtBarCode)
        ctxtBarCode.SelLength = 0
    Else
        ctxt体检条码.Text = ""
        ctxt体检条码.Enabled = True
        ctxt体检条码.SetFocus
    End If
End Sub

Private Sub ccmdAutoFill_Click()
    On Error GoTo errHandler
    ccmdUnfilledAllPass_Click
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病体检结果录入", "frmHEENT_ResultInput", "ccmdAutoFill_Click", 6666, lstrError, False
End Sub

Private Sub ccmdClear_Click()
    cgrdInfoBatch.Clear
    cgrdInfoBatch.rows = 1
    cgrdInfoBatch.FormatString = "体检条码编号"
    TotalPeopleBatch.Caption = 0
End Sub

Private Sub ccmdClearResult_Click()
    If SSTResultIn.Tab = 0 Then
        If coptClasses(2).Value = False Then
            ResultEye.Clear
            sub调整结果表头格式_眼科
        Else
            ResultEyeArmy.Clear
            sub调整结果表头格式_眼科_涉核部队
        End If
    ElseIf SSTResultIn.Tab = 1 Then
        ResultEar.Clear
        sub调整结果表头格式_耳鼻喉科
    End If
End Sub

'ccmdLocate暂时隐藏掉了
Private Sub ccmdLocate_Click()
    On Error GoTo errHandler
    Dim lobjRec As Object                       '单位定位返回的结果记录。
    Set lobjRec = pobj业务对象.func单位定位     '启动单位定位界面。
    
    '获取定位的单位，显示在“单位名称”录入框中。(暂时只显示“单位名称”)
    '-----不知道这里需不需要在其他模块里面设定涉核部队。
    If Not lobjRec Is Nothing Then
        If lobjRec.recordcount > 0 Then
            ctxtCompanyName.Text = IIf(IsNull(lobjRec("单位名称")), "", lobjRec("单位名称"))
        End If
    End If
    Set lobjRec = Nothing
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病体检结果录入", "frmHEENT_ResultInput", "ccmdLocate_Click", 6666, lstrError, False
End Sub

Private Sub ccmdLocateBatch_Click()

End Sub

Private Sub ccmdQuery_Click()
    On Error GoTo errHandler
    
    Dim lobjTmp, lobjRec As Object
    Dim i As Integer, j As Integer
    Dim lstrWhere As String
    Dim coptIndex As Integer
    
    For i = 0 To coptClasses.Count - 1
        If coptClasses(i).Value = True Then coptIndex = i
    Next
    lstrWhere = " and 体检类型='" & coptClasses(coptIndex).Caption & "'"
        
    '组装查询条件
    If cchkDate.Value = 1 Then
        lstrWhere = lstrWhere & " and 体检日期>='" & Format(cdtpDate.Value, "yyyy-mm-dd 00:00:00") & "' and 体检日期<='" & Format(cdtpDate.Value, "yyyy-mm-dd 23:59:59") & "'"
    End If
    
    If cchkName.Value = 1 Then
        If ctxtCheckName.Text = "" Then
            MsgBox ("若要查询姓名，则姓名不能为空。")
            Exit Sub
        End If
        lstrWhere = lstrWhere & " and 姓名='" & Trim(ctxtCheckName.Text) & "'"
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
    sfsub错误处理 "职业病体检结果录入", "frmHEENT_ResultInput", "ccmdQuery_Click", 6666, lstrError, False
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

'总的保存函数。暂时的设定是，单击这个按钮，眼科(包括图片)和耳鼻喉科的结果一起保存和提示。
Private Sub ccmdSave_Click()
    On Error GoTo errHandler
    
    Dim lstrCheck, lstrItem, lstrResult As String
    Dim i, j, isOk As Integer
    Dim lobjTmp As Object
    Dim lcolConclusion As String '五官科的体检结论
    
    '录入结果界面暂时不能操作
    ccmdUnfilledAllPass.Enabled = False
    ccmdAutoFill.Enabled = False
    ccmdSave.Enabled = False
    ctlb工具栏.Buttons(2).Enabled = False
    ccmdClear.Enabled = False
    SSTResultIn.Enabled = False
    
    Set lcolResult = New Collection
    Set lcolItem = New Collection
    
    '再保存表格中的结果数据
    If SSTPersonalInfo.Tab = 0 Then                 '此时为单个录入
    
        '保存单个项目的医生结论
        lcolConclusion = ctxtConclun.Text
        pobj业务对象.sub单个填写体检结论 ctxtBarCode.Text, priDeptName, lcolConclusion, um用户编号
        
        '先保存结果图片。仅放射工作和涉核部队的眼科有图片需保存
        If coptClasses(0).Value = False And DrawState <> -1 Then ccmdSavePicture_Click
        'If SSTResultIn.Tab = 1 Then GoTo ENTFill    '跳转到耳鼻喉科结果添加部分
        If coptClasses(2).Value = False Then        '接下来是眼科结果添加
            For i = 2 To 15
                If ResultEye.RowHidden(i) = False Then
                    If Not (ResultEye.TextMatrix(i, 1) = "裸眼" Or ResultEye.TextMatrix(i, 1) = "矫正" Or ResultEye.TextMatrix(i, 1) = "眼科其它") Then
                        lstrCheck = sub添加单项结果(ResultEye.TextMatrix(i, 2), (ResultEye.TextMatrix(i, 1) & "-右"), lstrCheck)
                        lstrCheck = sub添加单项结果(ResultEye.TextMatrix(i, 7), (ResultEye.TextMatrix(i, 1) & "-左"), lstrCheck)
                    ElseIf ResultEye.TextMatrix(i, 1) = "裸眼" Then
                        lstrCheck = sub添加单项结果(ResultEye.TextMatrix(i, 3), "裸眼-远视力-右", lstrCheck)
                        lstrCheck = sub添加单项结果(ResultEye.TextMatrix(i, 5), "裸眼-近视力-右", lstrCheck)
                        lstrCheck = sub添加单项结果(ResultEye.TextMatrix(i, 8), "裸眼-远视力-左", lstrCheck)
                        lstrCheck = sub添加单项结果(ResultEye.TextMatrix(i, 10), "裸眼-近视力-左", lstrCheck)
                    ElseIf ResultEye.TextMatrix(i, 1) = "矫正" Then
                        lstrCheck = sub添加单项结果(ResultEye.TextMatrix(i, 3), "矫正-远视力-右", lstrCheck)
                        lstrCheck = sub添加单项结果(ResultEye.TextMatrix(i, 5), "矫正-近视力-右", lstrCheck)
                        lstrCheck = sub添加单项结果(ResultEye.TextMatrix(i, 8), "矫正-远视力-左", lstrCheck)
                        lstrCheck = sub添加单项结果(ResultEye.TextMatrix(i, 10), "矫正-近视力-左", lstrCheck)
                    ElseIf ResultEye.TextMatrix(i, 1) = "眼科其它" Then
                        lstrCheck = sub添加单项结果(ResultEye.TextMatrix(i, 2), "眼科其它", lstrCheck)
                    End If
                End If
            Next i
        Else
            '添加涉核部队的单项结果，可以用公共函数 sub添加单项结果 和 subSave
            
            '仅有左右眼别的项目结果
            For i = 2 To 5
                lstrCheck = sub添加单项结果(ResultEyeArmy.TextMatrix(i, 2), (ResultEyeArmy.TextMatrix(i, 1) & "-右"), lstrCheck)
                lstrCheck = sub添加单项结果(ResultEyeArmy.TextMatrix(i, 6), (ResultEyeArmy.TextMatrix(i, 1) & "-左"), lstrCheck)
            Next i
            
            '左右眼别下，分别都有3项检查项目的结果填写
            For i = 7 To 11
                For j = 2 To 8
                    If j <= 4 Then lstrCheck = sub添加单项结果(ResultEyeArmy.TextMatrix(i, j), (ResultEyeArmy.TextMatrix(i, 1) & "-" & ResultEyeArmy.TextMatrix(6, j) & "-右"), lstrCheck)
                    If j >= 6 Then lstrCheck = sub添加单项结果(ResultEyeArmy.TextMatrix(i, j), (ResultEyeArmy.TextMatrix(i, 1) & "-" & ResultEyeArmy.TextMatrix(6, j) & "-左"), lstrCheck)
                Next j
            Next i
            
            '"诊断"的结果填写
            lstrCheck = sub添加单项结果(ResultEyeArmy.TextMatrix(12, 2), "眼科诊断", lstrCheck)
        End If
        
'ENTFill:
        'If SSTResultIn.Tab = 1 Then
            For i = 1 To 18
                If ResultEar.RowHidden(i) = False Then
                    If i <= 5 Or i >= 16 Then
                        lstrCheck = sub添加单项结果(ResultEar.TextMatrix(i, ResultEar.cols - 1), ResultEar.TextMatrix(i, 0), lstrCheck)
                    Else
                        lstrCheck = sub添加单项结果(ResultEar.TextMatrix(i, ResultEar.cols - 1), ResultEar.TextMatrix(i, 0) & "-" & ResultEar.TextMatrix(i, 1), lstrCheck)
                    End If
                End If
            Next
        'End If
        
    Else '此时为批量录入
         '---------
         
         
         '---------
    End If
    
    'lstrcheck字符串检查
    If (Not lstrCheck = "") And (Not ResultChanged = 2) Then
        isOk = MsgBox("以下项目未填写结果，确定保存吗？" & Chr(10) & "未填写项将不会记录到数据库！" & Chr(10) & Chr(10) & Trim(lstrCheck), vbOKCancel)
        If isOk = 2 Then
            Set lcolResult = Nothing
            Set lcolItem = Nothing
            ccmdUnfilledAllPass.Enabled = True
            ccmdAutoFill.Enabled = True
            ccmdSave.Enabled = True
            ctlb工具栏.Buttons(2).Enabled = True
            ccmdClearResult.Enabled = True
            Exit Sub
        End If
    End If
        
    If ResultChanged = 2 Then
        isOk = MsgBox("是否保存该体检人员的修改结果？", vbOKCancel)
        If isOk = 1 Then
            subSave         '里面包含保存成功提示
        Else
            LoadPersonalInfo (ctxtBarCode)
        End If
    Else
        subSave
    End If
    
    Set lcolResult = Nothing
    Set lcolItem = Nothing
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病体检结果录入", "frmHEENT_ResultInput", "ccmdSave_Click", 6666, lstrError, False
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
    lstrWhere = " and 体检类型='" & coptClasses(coptIndex).Caption & "'"
        
    '组装查询条件
    If cchkDateBatch.Value = 1 Then
        lstrWhere = lstrWhere & " and 体检日期>='" & Format(cdtpDateBatch.Value, "yyyy-mm-dd 00:00:00") & "' and 体检日期<='" & Format(cdtpDateBatch.Value, "yyyy-mm-dd 23:59:59") & "'"
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
    sfsub错误处理 "职业病体检结果录入", "frmHEENT_ResultInput", "ccmdQuery_Click", 6666, lstrError, False
End Sub

Private Sub ccmdUnfilledAllPass_Click()
    On Error GoTo errHandler
    
    Dim hasFilled As Boolean
    Dim i, j As Integer
    
    If SSTResultIn.Tab = 1 Then GoTo ENTFill
    If coptClasses(2).Value = False Then
        '职业健康和放射人员的未填写项自动填满“正常”部分
        For i = 2 To 15
            If ResultEye.RowHidden(i) = False Then  '区分体检类别
                hasFilled = True
                For j = 2 To 5
                    If i <> 5 And i <> 6 And ResultEye.TextMatrix(i, j) = "" Then hasFilled = False: Exit For
                Next j
                If hasFilled = False Then For j = 2 To 5: ResultEye.TextMatrix(i, j) = "正常": Next
                
                hasFilled = True
                For j = 7 To 10
                    If i <> 5 And i <> 6 And ResultEye.TextMatrix(i, j) = "" Then hasFilled = False: Exit For
                Next j
                If hasFilled = False Then For j = 7 To 10: ResultEye.TextMatrix(i, j) = "正常": Next j
                
                If ResultEye.TextMatrix(5, 3) = "" Then ResultEye.TextMatrix(5, 3) = "正常"
                If ResultEye.TextMatrix(5, 5) = "" Then ResultEye.TextMatrix(5, 5) = "正常"
                If ResultEye.TextMatrix(5, 8) = "" Then ResultEye.TextMatrix(5, 8) = "正常"
                If ResultEye.TextMatrix(5, 10) = "" Then ResultEye.TextMatrix(5, 10) = "正常"
                If ResultEye.TextMatrix(6, 3) = "" Then ResultEye.TextMatrix(6, 3) = "正常"
                If ResultEye.TextMatrix(6, 5) = "" Then ResultEye.TextMatrix(6, 5) = "正常"
                If ResultEye.TextMatrix(6, 8) = "" Then ResultEye.TextMatrix(6, 8) = "正常"
                If ResultEye.TextMatrix(6, 10) = "" Then ResultEye.TextMatrix(6, 10) = "正常"
                
                If i = 15 And ResultEye.TextMatrix(i, 2) = "正常" And ResultEye.TextMatrix(i, 6) = "" Then
                    ResultEye.TextMatrix(15, 6) = "正常"
                End If
            End If
        Next i
        sub调整结果表头格式_眼科
    Else
        '涉核部队的未填写项自动填满“正常”部分
        For i = 2 To 5
            If ResultEyeArmy.TextMatrix(i, 2) = "" Then
                For j = 2 To 4: ResultEyeArmy.TextMatrix(i, j) = "正常": Next
            End If
            If ResultEyeArmy.TextMatrix(i, 6) = "" Then
                For j = 6 To 8: ResultEyeArmy.TextMatrix(i, j) = "正常": Next
            End If
        Next i
        
        For i = 7 To 11
            For j = 2 To 8
                If j <> 5 And ResultEyeArmy.TextMatrix(i, j) = "" Then ResultEyeArmy.TextMatrix(i, j) = "正常"
            Next j
        Next i
    
        If ResultEyeArmy.TextMatrix(12, 2) = "" Then
            For j = 2 To 8: ResultEyeArmy.TextMatrix(12, j) = "正常": Next
        End If
        
        sub调整结果表头格式_眼科_涉核部队
    End If
    Exit Sub
    
ENTFill:
    For i = 1 To 18
        If ResultEar.RowHidden(i) = False And ResultEar.TextMatrix(i, ResultEar.cols - 1) = "" Then ResultEar.TextMatrix(i, ResultEar.cols - 1) = "正常"
    Next
    sub调整结果表头格式_耳鼻喉科
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病体检结果录入", "frmHEENT_ResultInput", "ccmdUnfilledAllPass_Click", 6666, lstrError, False
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

Private Sub cgrdInfo_DblClick()
    If coptClasses(2).Value = False Then
        ResultEye.Clear
        sub调整结果表头格式_眼科
    Else
        ResultEyeArmy.Clear
        sub调整结果表头格式_眼科_涉核部队
    End If
    ResultEar.Clear
    sub调整结果表头格式_耳鼻喉科
    indX = cgrdinfo.MouseRow
    indY = cgrdinfo.MouseCol
    If indX < 0 Or indY < 0 Then
        Exit Sub
    ElseIf indX >= 0 And indX < cgrdinfo.rows And indY >= 0 And indY < cgrdinfo.cols Then
        ctxtBarCode.Text = cgrdinfo.TextMatrix(indX, 0)
        ctxtBarCode_LostFocus
        
        '2012-07-03 于登淼 ↓
        '每次读入个人信息时，判断是否超过修改时间。
        '以此控制保存按钮是否可用。
        If pobj业务对象.sub是否在修改时间范围内(ctxtBarCode.Text, priDeptName, 8) = False Then
            ctlb工具栏.Buttons(2).Enabled = False
        End If
        '2012-07-03 于登淼 ↑
    End If
End Sub

'''Private Sub cgrdInfoBatch_Click()
'''    cgrdInfoBatch.SelectionMode = flexSelectionByRow
'''End Sub

Private Sub cgrdInfoBatch_DblClick()
    If cchk套用体检结果.Value = 0 Then
        If coptClasses(2).Value = False Then
            ResultEye.Clear
            sub调整结果表头格式_眼科
        Else
            ResultEyeArmy.Clear
            sub调整结果表头格式_眼科_涉核部队
        End If
        ResultEar.Clear
        sub调整结果表头格式_耳鼻喉科
    End If
    indX = cgrdInfoBatch.MouseRow
    indY = cgrdInfoBatch.MouseCol
    If indX <= 0 Or indY < 0 Then
        Exit Sub
    ElseIf indX > 0 And indX < cgrdInfoBatch.rows And indY >= 0 And indY < cgrdInfoBatch.cols Then
        ctxt体检条码.Text = cgrdInfoBatch.TextMatrix(indX, 0)
        ctxt体检条码_lostfocus
    End If
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

Private Sub coptClasses_Click(Index As Integer)
    If coptClasses(0).Value = True Or coptClasses(1).Value = True Then
        ResultEye.Visible = True
        ResultEyeArmy.Visible = False
        sub调整结果表头格式_眼科
    End If
    If coptClasses(2).Value = True Then
        ResultEye.Visible = False
        ResultEyeArmy.Visible = True
        sub调整结果表头格式_眼科_涉核部队
    End If
    sub调整结果表头格式_耳鼻喉科
    If coptClasses(0).Value = True Then fraQuery.Caption = "查询体检人员(职业健康)"
    If coptClasses(1).Value = True Then fraQuery.Caption = "查询体检人员(放射工作)"
    If coptClasses(2).Value = True Then fraQuery.Caption = "查询体检人员(涉核部队)"
    
    Dim coptIndex As Integer
    coptIndex = Index
    sub查询列表显示 coptIndex
End Sub
Private Sub ctxtAge_Change()
    If ctxtAge.Text = "" Then Exit Sub
    If IsNumeric(CLng(ctxtAge.Text)) = False Then
        MsgBox ("年龄必须为小于150的数字！")
        Exit Sub
    ElseIf CLng(ctxtAge.Text) >= 150 Then
        MsgBox ("年龄要小于150！")
    End If
    ctxtAge.Text = CLng(ctxtAge.Text)
End Sub

Private Sub ctxtAge_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then: ctxtCompanyName.SetFocus
End Sub

Private Sub ctxtBarCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then ctxtBarCode_LostFocus
End Sub

Private Sub ctxtBarCode_LostFocus()
    Dim lstrNo As String
    Dim i As Integer
    Dim str科室结论 As String
    Dim lcol职业病对象 As Object
    lstrNo = Trim(ctxtBarCode.Text)
    
    '检查条码号是否存在
    Dim lobjTmp As Object
    Set lobjTmp = CreateObject("职业病体检结果录入.clscommon")
    Set mlobjRec = lobjTmp.func获取体检人员基本信息(lstrNo)
    If mlobjRec.recordcount = 0 Then
        Set mlobjRec = Nothing
        
        '清空当前个人信息
        ctxtBarCode.Enabled = True
        ctxtName.Text = ""
        ctxtSex.Text = ""
        ctxtAge.Text = ""
        ctxtCompanyName.Text = ""
        Exit Sub
    End If
    
    '判断是否该科室有此人员的体检权限
    Set lobjTmp = CreateObject("职业病体检结果录入.clscommon")
    If lobjTmp.func获取体检人员体检科室信息(lstrNo, priDeptName) Then
        Set lobjTmp = Nothing
        '载入已有的个人信息和现有的体检结果
        LoadPersonalInfo (lstrNo)
        
        Set lcol职业病对象 = CreateObject("职业病对象.clsManageMedicalExam")
        str科室结论 = lcol职业病对象.func返回科室结论(ctxtBarCode.Text, priDeptName)
        ctxtConclun.Text = str科室结论
        
        '一旦确定当前体检人员编号，就不能更改。除非，清空界面内容。
        ctxtBarCode.Enabled = False
        ctxtName.Enabled = False
        ctxtSex.Enabled = False
        ctxtAge.Enabled = False
        ctxtCompanyName.Enabled = False '其实单位灰掉了之后，如果有“单位定位”按钮，还是可以改的。
        If ResultChanged <> 3 Then
            ccmdUnfilledAllPass.Enabled = True
            ccmdAutoFill.Enabled = True
            ccmdSave.Enabled = True
            ctlb工具栏.Buttons(2).Enabled = True
            ccmdClearResult.Enabled = True
            '2012-06-27 于登淼 ↓
            '每次读入个人信息时，判断是否超过修改时间。
            '以此控制保存按钮是否可用。
            If pobj业务对象.sub是否在修改时间范围内(ctxtBarCode.Text, priDeptName, 8) = False Then
                ctlb工具栏.Buttons(2).Enabled = False
                ccmdSave.Enabled = False
            End If
            '2012-06-27 于登淼 ↑
        End If
        SSTResultIn.Enabled = True
        For i = 0 To 2
            If coptClasses(i).Value = False Then coptClasses(i).Enabled = False
        Next i
        If coptClasses(0).Value = False Then
            fraDraw.Enabled = True
            DrawLineWidth = Pow_2(2)
            If DrawState <> -1 Then ccmdDraw_Click
        End If
        If ResultChanged = 3 Then fraDraw.Enabled = False
    Else
        Set lobjTmp = Nothing
        MsgBox ("没有该条码对应的体检人员信息！")
        If cgrdinfo.rows > 0 Then cgrdinfo.RemoveItem
        subClear   '''2012-07-04 于登淼 临时注释，setfocus在体检条码号错误时陷入死循环。查询功能失效
    End If
End Sub

Private Sub ctxtCheckName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then ccmdQuery_Click
End Sub

Private Sub ctxtCompanyName_Change()
    ctxtCompanyName.Text = Trim(ctxtCompanyName.Text)
End Sub

Private Sub ctxtName_Change()
    ctxtName.Text = Trim(ctxtName.Text)
End Sub

Private Sub ctxtName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then: ctxtSex.SetFocus
End Sub

Private Sub ctxtSex_Change()
    ctxtSex.Text = Trim(ctxtSex.Text)
End Sub

Private Sub ctxtSex_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then: ctxtAge.SetFocus
End Sub

Private Sub ctxt体检条码_lostfocus()
    Dim lstrNo As String
    Dim i As Integer
    Dim str科室结论 As String
    Dim lcol职业病对象 As Object
    lstrNo = Trim(ctxt体检条码.Text)
    
    '检查条码号是否存在
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
    
    Set lobjTmp = CreateObject("职业病体检结果录入.clscommon")
    If lobjTmp.func获取体检人员体检科室信息(lstrNo, priDeptName) Then
        Set lobjTmp = Nothing
       
        LoadPersonalInfoBatch (lstrNo)
        If cchk套用体检结果.Value = 0 Then
            Set lcol职业病对象 = CreateObject("职业病对象.clsManageMedicalExam")
            str科室结论 = lcol职业病对象.func返回科室结论(ctxt体检条码.Text, priDeptName)
            ctxtConclun.Text = str科室结论
        End If
        '一旦确定当前体检人员编号，就不能更改。除非，清空界面内容。
        ctxt体检条码.Enabled = False
        ctxt姓名.Enabled = False
        ctxt性别.Enabled = False
        ctxt年龄.Enabled = False
        ctxt单位名称.Enabled = False '其实单位灰掉了之后，如果有“单位定位”按钮，还是可以改的。
        If ResultChanged <> 3 Then
            ccmdUnfilledAllPass.Enabled = True
            ccmdAutoFill.Enabled = True
            ccmdSave.Enabled = True '''''''''''''''''''''''''''不懂了，为啥之前写的是false
            ctlb工具栏.Buttons(3).Enabled = True
            ccmdClearResult.Enabled = True
        End If
        SSTResultIn.Enabled = True
        For i = 0 To 2
            If coptClasses(i).Value = False Then coptClasses(i).Enabled = False
        Next i
        If coptClasses(0).Value = False Then
            fraDraw.Enabled = True
            DrawLineWidth = Pow_2(2)
            If DrawState <> -1 Then ccmdDraw_Click
        End If
        If ResultChanged = 3 Then fraDraw.Enabled = False
        ctlb工具栏.Buttons(2).Enabled = False
        ctlb工具栏.Buttons(3).Enabled = True
        ccmdSave.Enabled = False
    Else
        Set lobjTmp = Nothing
        MsgBox ("该体检人员没有该科室的体检项目！")
        cgrdInfoBatch.RemoveItem
        subClear
    End If
End Sub

Private Sub ResultEye_AfterEdit(ByVal row As Long, ByVal col As Long)
    sub调整结果表头格式_眼科
End Sub

Private Sub ResultEye_DblClick()
    If ResultChanged = 3 Then Exit Sub
    indX = ResultEye.MouseRow
    indY = ResultEye.MouseCol
    ResultEye.Editable = flexEDKbdMouse
    If indX < 0 Or indY < 0 Then
        Exit Sub
    ElseIf indX >= 0 And indX < ResultEye.rows And indY >= 0 And indY < ResultEye.cols Then
        sub录入修改内容 indX, indY
        If ResultChanged = 1 Then ResultChanged = 2
    End If
End Sub


'只是测试保存功能。这部分写得相当不完满。。。
'还有啊，这个 mouseCol 和 mouseRow 的变量是不是只能用一次，之后，系统自动删掉啊？
Private Sub ResultEye_Click()
    If ResultChanged = 3 Then Exit Sub
    indX = ResultEye.MouseRow
    indY = ResultEye.MouseCol
    'MsgBox ("x = " & indX & ",y = " & indY)
    Dim i As Integer
    ResultEye.Editable = flexEDKbdMouse
    If indX < 0 Or indY < 0 Then
        Exit Sub
    ElseIf indX = 15 And indY >= 0 And indY <= 10 Then
        For i = 2 To 10: ResultEye.TextMatrix(indX, i) = "正常": Next
        If ResultChanged = 1 Then ResultChanged = 2
    ElseIf indY >= 2 And indY <= 5 And ResultEye.TextMatrix(indX, indY) = "" Then
        For i = 2 To 5: ResultEye.TextMatrix(indX, i) = "正常": Next
        If ResultChanged = 1 Then ResultChanged = 2
    ElseIf indY >= 7 And indY <= 10 And ResultEye.TextMatrix(indX, indY) = "" Then
        For i = 7 To 10: ResultEye.TextMatrix(indX, i) = "正常": Next
        If ResultChanged = 1 Then ResultChanged = 2
    End If
    sub调整结果表头格式_眼科
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    ctxtBarCode.SetFocus
    ctxtBarCode.SelStart = Len(ctxtBarCode)
    ctxtBarCode.SelLength = 0
    
End Sub

Private Sub Form_Load()
    Dim lcol工具栏按钮 As New Collection '工具栏上的按钮初始化集合。
    
    On Error GoTo errHandler
    
    '如果窗体已经初始化过则不再进行初始化。
    If mblnInUse Then Exit Sub
    
    '设置窗体正在使用的标志。
    mblnInUse = True
    
    '初始化界面通用对象（每个界面对应一个界面通用对象实例，不可混用，切记切记）。
    Set mobjGUI = New cls界面通用对象
    
    '设置工具栏上所需要的各种按钮。
    With lcol工具栏按钮
        .Add "清空界面(&N)110"
        .Add "保存"
        .Add "批量保存(&S)"
        .Add "|"
        .Add "退出"
    End With
    
    '为需要通过界面通用对象控制的各种控件赋初始值。
    With mobjGUI
        Set .Form = Me
        Set .c工具栏 = ctlb工具栏
    End With
    
    '调用界面通用对象提供的方法，对界面控件进行初始化。
    mobjGUI.subInitialize lcol工具栏按钮, ""

    '保存结果时，医师姓名为当前用户名。
    mstrDoctorName = um用户名
    LabelDoctor.Caption = LabelDoctor.Caption & " " & mstrDoctorName
    
    '界面权限设置。仅限于该界面上各个按钮和其它控件的使用。
    '大致功能暂时有：查看、修改与保存，这两种。不允许删除操作（也没有删除按钮）
    Dim lobjTmp As Object
    Set lobjTmp = CreateObject("职业病设置.clsPermissionConfigure")
    If lobjTmp.func科室操作权限(um用户编号, "职业病体检_五官科结果录入_修改") = False Then
        ResultChanged = 3           '限制vsFlexGrid.Enabled
        ccmdUnfilledAllPass.Visible = False
        ccmdAutoFill.Visible = False
        ccmdSave.Visible = False
        ccmdClearResult.Visible = False
        ctlb工具栏.Buttons(2).Visible = False
    End If
    
    '2012-05-22 翁乔 ↓↓↓
    '界面权限设置
    If lobjTmp.func科室操作权限(um用户编号, "职业病体检_五官科结果录入_批量修改") = False Then
        ctlb工具栏.Buttons(3).Visible = False
    End If
    '2012-05-22 ↑↑↑
    Set lobjTmp = Nothing
    
    '加载内容较少，所以先执行了。而且，默认加载职业健康部分
    sub调整结果表头格式_眼科
    sub调整结果表头格式_耳鼻喉科
    
    '界面按钮设定
    cdtpConclusionDate = Now       '录入结论日期设为当前日期
    cdtpDate.Value = Now           '查询体检日期设为当前日期
    ctxtBarCode.Enabled = True
    ccmdUnfilledAllPass.Enabled = False
    ccmdAutoFill.Enabled = False
    ccmdSave.Enabled = False
    ctlb工具栏.Buttons(2).Enabled = False
    ctlb工具栏.Buttons(3).Enabled = False
    ccmdClearResult.Enabled = False
    SSTResultIn.Enabled = False
    If ResultChanged <> 3 Then ResultChanged = -1
    coptClasses(0).Value = True
    DTP录入日期.Value = Now
    cdtpDateBatch.Value = Now
    '批量查询功能查询体检表读取
    
    
    'lobj批量操作对象
        
    '画图部分设定
    DrawState = -1       'form_load时，画图状态为“未做修改不保存”
    If coptClasses(0).Value = True And SSTResultIn.Tab = 0 Then Picture3.Picture = Nothing '仅放射工作和涉核部队有图片保存
    sub原图解析
    
    '询问frame初始设定
    fraQuery.Caption = "查询体检人员(职业健康)"

    '2012-07-03 于登淼 ↓
    '更改系统编号固定部分。省疾控新要求中改变系统编号规则。
    '获取系统编号固定部分。
    sub获取系统编号固定部分
    '2012-07-03 于登淼 ↑
    
    '2012-07-14 于登淼 ↓
    '初始化查询界面，调整查询列表格式。初始化科室基本信息。
    priDeptName = "五官科"
    priDeptNo = "01"
    priDeptResultName = "五官科"
    ccmdQuery_Click
    SSTPersonalInfo.Tab = 1: ccmdSelInfo_Click: SSTPersonalInfo.Tab = 0
    '2012-07-14 于登淼 ↑
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病体检结果录入", "frmHEENT_ResultInput", "Form_Load", 6666, lstrError, False
    Exit Sub
    Resume

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    mblnInUse = False
    Set mobjGUI = Nothing
    Set mobjQueryResult = Nothing
    Unload frmHEENT_ResultInput
    Set frmHEENT_ResultInput = Nothing
End Sub

Private Sub SSTPersonalInfo_Click(PreviousTab As Integer)
    If SSTPersonalInfo.Tab = 0 Then
        ctlb工具栏.Buttons(3).Enabled = False
    ElseIf SSTPersonalInfo.Tab = 1 Then
        ccmdSave.Enabled = False
        ctlb工具栏.Buttons(2).Enabled = False
    End If
End Sub

Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    On Error GoTo errHandler
    Dim i As Integer
    
    Cancel = True
    Select Case Operate
    Case "清空界面"
        subClear
    Case "保存"
        '2012-07-03 于登淼 ↓
        '判断是否在修改时间范围内
        If pobj业务对象.sub是否在修改时间范围内(Trim(ctxtBarCode.Text), priDeptName, 8) = False Then
            MsgBox ("距上次修改已经超过8小时。请与管理员联系获得修改权限后再继续。")
            Exit Sub
        End If
        '2012-07-03 于登淼 ↑
        
        '2012-07-15 于登淼 ↓
        '没有录入体检结论时，提示且不保存。
        If Len(Trim(ctxtConclun.Text)) > 0 Then
            MsgBox "你还没有为病人下结论"
            GoTo errHandler
        End If
        '2012-07-15 于登淼 ↑
        
        '2012-07-03 于登淼 ↓
        '增加一个字段"修改起始时间"的修改。同时修改该科室的体检结果录入状态。
        pobj业务对象.sub修改起始时间 Trim(ctxtBarCode.Text), priDeptName
        pobj业务对象.sub修改结果录入状态 Trim(ctxtBarCode.Text), priDeptNo, "2"  '   01表示五官科
        pobj业务对象.sub结果录入修改体检状态 Trim(ctxtBarCode.Text), "4"
        '2012-07-03 于登淼 ↑
        
        ccmdSave_Click
        
        '2012-07-15 于登淼 ↓
        '保存完之后，重新进行查询。
        ccmdQuery_Click
        i = SSTPersonalInfo.Tab
        SSTPersonalInfo.Tab = 1: ccmdSelInfo_Click: SSTPersonalInfo.Tab = i
        '2012-07-15 于登淼 ↑
    Case "批量保存"
        '2012-07-15 于登淼 ↓
        '没有录入体检结论时，提示且不保存。
        If Len(Trim(ctxtConclun.Text)) > 0 Then
            MsgBox "你还没有为病人下结论"
            GoTo errHandler
        End If
        '2012-07-15 于登淼 ↑
        
        sub批量保存
        
        '2012-07-15 于登淼 ↓
        '保存完之后，重新进行查询。
        ccmdQuery_Click
        i = SSTPersonalInfo.Tab
        SSTPersonalInfo.Tab = 1: ccmdSelInfo_Click: SSTPersonalInfo.Tab = i
        '2012-07-15 于登淼 ↑
    Case "退出"
        Unload frmHEENT_ResultInput
        Set frmHEENT_ResultInput = Nothing
    End Select
    
    Exit Sub
errHandler:
    sfsub错误处理 "职业病体检结果录入", "frmHEENT_ResultInput", "mobjGUI_BeforeOperate", Err.Number, Err.Description, False
    Cancel = True
    Exit Sub
    Resume
End Sub

Sub sub调整结果表头格式_眼科()
    Dim i As Integer
    
    '调整表格大小和位置(手动调整)
    ResultEye.Height = 3500
    ResultEye.Left = 120
    ResultEye.Top = 240
    ResultEye.Width = 8500
    
    '调整列宽和字体位置
    'ResultEye.AutoSize 2, ResultEye.Cols - 1, 1, 650 '如果修改了内容，表格会由于不明原因拓宽。暂不用这句。
    ResultEye.ColWidth(0) = 960
    ResultEye.ColWidth(1) = 640
    For i = 0 To 10: ResultEye.ColAlignment(i) = flexAlignCenterCenter: Next
    For i = 2 To 10: ResultEye.ColWidth(i) = 800: Next
    
    '区分职业健康和放射工作，某些项需要隐藏
    If coptClasses(0) = True Then
        ResultEye.RowHidden(7) = True
        ResultEye.RowHidden(9) = False
        ResultEye.RowHidden(13) = True
        ResultEye.RowHidden(15) = False
    Else
        ResultEye.RowHidden(7) = False
        ResultEye.RowHidden(9) = True
        ResultEye.RowHidden(13) = False
        ResultEye.RowHidden(15) = True
    End If
    
    '0~15行，0~1列范围内，值相同的要合并单元格
    ResultEye.MergeCompare = flexMCIncludeNulls
    ResultEye.MergeCells = flexMergeFree
    For i = 0 To 15: ResultEye.MergeRow(i) = True: Next
    For i = 0 To 1: ResultEye.MergeCol(i) = True: Next
    
    '0、1列，前两列格式设定
    ResultEye.TextMatrix(0, 0) = "项目": ResultEye.TextMatrix(0, 1) = "项目"
    ResultEye.TextMatrix(1, 0) = "眼别": ResultEye.TextMatrix(1, 1) = "眼别"
    ResultEye.TextMatrix(2, 0) = "色觉": ResultEye.TextMatrix(2, 1) = "色觉"
    ResultEye.TextMatrix(3, 0) = "暗适应": ResultEye.TextMatrix(3, 1) = "暗适应"
    ResultEye.TextMatrix(4, 0) = "视野": ResultEye.TextMatrix(4, 1) = "视野"
    ResultEye.TextMatrix(5, 0) = "视力": ResultEye.TextMatrix(6, 0) = "视力"
    ResultEye.TextMatrix(5, 1) = "裸眼": ResultEye.TextMatrix(6, 1) = "矫正"
    ResultEye.TextMatrix(7, 0) = "眼前部": ResultEye.TextMatrix(7, 1) = "眼前部"
    ResultEye.TextMatrix(8, 0) = "晶体裂隙灯" & Chr(10) & "检查所见": ResultEye.TextMatrix(9, 0) = "晶体裂隙灯" & Chr(10) & "检查所见": ResultEye.TextMatrix(10, 0) = "晶体裂隙灯" & Chr(10) & "检查所见": ResultEye.TextMatrix(11, 0) = "晶体裂隙灯" & Chr(10) & "检查所见": ResultEye.TextMatrix(12, 0) = "晶体裂隙灯" & Chr(10) & "检查所见"
    ResultEye.TextMatrix(8, 1) = "角膜": ResultEye.TextMatrix(9, 1) = "结膜": ResultEye.TextMatrix(10, 1) = "前房": ResultEye.TextMatrix(11, 1) = "虹膜": ResultEye.TextMatrix(12, 1) = "晶状体"
    ResultEye.TextMatrix(13, 0) = "玻璃体": ResultEye.TextMatrix(13, 1) = "玻璃体"
    ResultEye.TextMatrix(14, 0) = "眼底": ResultEye.TextMatrix(14, 1) = "眼底"
    ResultEye.TextMatrix(15, 0) = "眼科其它": ResultEye.TextMatrix(15, 1) = "眼科其它"
    
    '2~10列格式设定，中间列为第6列，值全为***，为了两列单元格不合并，默认隐藏。
    ResultEye.ColHidden(6) = True
    For i = 1 To 14: ResultEye.TextMatrix(i, 6) = "***": Next
    For i = 2 To 10: ResultEye.TextMatrix(0, i) = "检查结果": Next
    For i = 2 To 5: ResultEye.TextMatrix(1, i) = "右": Next
    For i = 7 To 10: ResultEye.TextMatrix(1, i) = "左": Next
    ResultEye.TextMatrix(5, 2) = "远视力": ResultEye.TextMatrix(5, 7) = "远视力"
    ResultEye.TextMatrix(6, 2) = "远视力": ResultEye.TextMatrix(6, 7) = "远视力"
    ResultEye.TextMatrix(5, 4) = "近视力": ResultEye.TextMatrix(5, 9) = "近视力"
    ResultEye.TextMatrix(6, 4) = "近视力": ResultEye.TextMatrix(6, 9) = "近视力"
    
    '某些单元格不可编辑
    '不知道该如何解决，暂时：每次编辑后刷新表头格式
    
    '表头颜色可以不太一样，也可以隔行一个颜色。
End Sub

Sub sub录入修改内容(ByVal paraRow As Integer, ByVal paraCol As Integer)
    Dim i As Integer
    If SSTResultIn.Tab = 0 Then
        '眼科结果填写。
        If coptClasses(2).Value = False Then
            ResultEye.Select paraRow, paraCol
            ResultEye.EditCell
        Else
            ResultEyeArmy.Select paraRow, paraCol
            ResultEyeArmy.EditCell
        End If
    Else
        '耳鼻喉科结果填写。
        ResultEar.Select paraRow, paraCol
        ResultEar.EditCell
    End If
    
End Sub

Sub LoadPersonalInfoBatch(ByVal paraSysNo As String)
    On Error GoTo errHandler
    
    Dim lobjTmp, lobjRec As Object
    Set lobjTmp = CreateObject("职业病体检结果录入.clscommon")
    Set mlobjRec = lobjTmp.func获取体检人员基本信息(paraSysNo)
    If mlobjRec.recordcount > 0 Then
        ctxt姓名 = mlobjRec("姓名")
        ctxt性别 = mlobjRec("性别")
        ctxt年龄 = mlobjRec("年龄")
        ctxt单位名称 = mlobjRec("单位名称")
        
        '设置体检类型
        If mlobjRec("体检类型") = "职业健康" Then coptClasses(0).Value = True
        If mlobjRec("体检类型") = "放射工作" Then coptClasses(1).Value = True
        If mlobjRec("体检类型") = "涉核部队" Then coptClasses(2).Value = True
        
        
         '载入已有的个人信息和现有的体检结果
        If cchk套用体检结果.Value = 0 Then
            '显示照片
            Set lobjRec = CreateObject("职业病对象.clspersonexamed")
            lobjRec.系统编号 = ctxt体检条码.Text
            Picture4.Enabled = True
            Picture4.Visible = True
            Picture4.Picture = lobjRec.像片
        
        
            mstr结果图片项目编号 = lobjTmp.func获取体检项目编号("晶状体环面及正面图")
            DrawState = -1
            
            Set lobjRec = lobjTmp.func是否已经体检过(ctxt体检条码.Text, priDeptName)
            If lobjRec.recordcount = 0 Then
                If ResultChanged <> 3 Then ResultChanged = 0
            ElseIf lobjRec.recordcount > 0 Then
                If ResultChanged <> 3 Then ResultChanged = 1
                If coptClasses(2).Value = False Then
                    sub填写已有的体检结果_眼科 lobjRec
                Else
                    sub填写已有的体检结果_眼科_涉核部队 lobjRec
                End If
                sub填写已有的体检结果_耳鼻喉科 lobjRec
            End If
            
            '显示眼睛检查结果图
            If coptClasses(0).Value = False Then Picture3.Picture = lobjTmp.func获取结果图片(ctxt体检条码.Text, mstr结果图片项目编号, "晶状体环面及正面图.bmp")  '01069是眼睛检查结果图的项目编号。
        End If
    Else
        MsgBox ("没有该条码对应的体检人员信息！")
        Exit Sub
    End If
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病体检结果录入", "frmHEENT_ResultInput", "LoadPersonalInfo", 6666, lstrError, False
End Sub

Sub LoadPersonalInfo(ByVal paraSysNo As String)
    On Error GoTo errHandler
    
    Dim lobjTmp, lobjRec As Object
    Set lobjTmp = CreateObject("职业病体检结果录入.clscommon")
    Set mlobjRec = lobjTmp.func获取体检人员基本信息(paraSysNo)
    If mlobjRec.recordcount > 0 Then
        ctxtName = mlobjRec("姓名")
        ctxtSex = mlobjRec("性别")
        ctxtAge = mlobjRec("年龄")
        ctxtCompanyName = mlobjRec("单位名称")
            
        '显示照片
        Set lobjRec = CreateObject("职业病对象.clspersonexamed")
        lobjRec.系统编号 = ctxtBarCode.Text
        Picture2.Enabled = True
        Picture2.Visible = True
        Picture2.Picture = lobjRec.像片
            
        '设置体检类型
        If mlobjRec("体检类型") = "职业健康" Then coptClasses(0).Value = True
        If mlobjRec("体检类型") = "放射工作" Then coptClasses(1).Value = True
        If mlobjRec("体检类型") = "涉核部队" Then coptClasses(2).Value = True
            
        mstr结果图片项目编号 = lobjTmp.func获取体检项目编号("晶状体环面及正面图")
        DrawState = -1
            
        Set lobjRec = lobjTmp.func是否已经体检过(ctxtBarCode.Text, priDeptName)
        If lobjRec.recordcount = 0 Then
            If ResultChanged <> 3 Then ResultChanged = 0
        ElseIf lobjRec.recordcount > 0 Then
            If ResultChanged <> 3 Then ResultChanged = 1
            If coptClasses(2).Value = False Then
                sub填写已有的体检结果_眼科 lobjRec
            Else
                sub填写已有的体检结果_眼科_涉核部队 lobjRec
            End If
            sub填写已有的体检结果_耳鼻喉科 lobjRec
        End If
    
        '显示眼睛检查结果图
        If coptClasses(0).Value = False Then Picture3.Picture = lobjTmp.func获取结果图片(ctxtBarCode.Text, mstr结果图片项目编号, "晶状体环面及正面图.bmp")  '01069是眼睛检查结果图的项目编号。
    Else
        MsgBox ("没有该条码对应的体检人员信息！")
        Exit Sub
    End If
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病体检结果录入", "frmHEENT_ResultInput", "LoadPersonalInfo", 6666, lstrError, False
End Sub

'清除界面上体检人员的所有信息(查询框内容暂不清除)，恢复界面设置为form_load时设置
Sub subClear()
    TotalPeople.Caption = 0
    TotalPeopleBatch.Caption = 0
    
    '清空当前个人信息
    cdtpConclusionDate = Now
    ctxtBarCode.Text = ""
    ctxtBarCode.Enabled = True
    ctxtName.Text = ""
    ctxtSex.Text = ""
    ctxtAge.Text = ""
    ctxtCompanyName.Text = ""
    cgrdinfo.rows = 1
    
    '批量信息清除
    ctxt体检条码.Text = ""
    ctxt体检条码.Enabled = True
    ctxt姓名.Text = ""
    ctxt性别.Text = ""
    ctxt年龄.Text = ""
    ctxt单位名称.Text = ""
    cgrdInfoBatch.rows = 1
    
    '套用信息标志清空
    cchk套用体检结果.Value = 0
    
    Picture1.Picture = Nothing
    Picture2.Picture = Nothing
    Picture3.Picture = Nothing
    Picture4.Picture = Nothing
    ctxtConclun.Text = ""

'    '清空查询结果（不一定要有的,也没写全）
'    cchkDate.Value = 0
'    cdtpDate.Value = Now
'    cgrdInfo.Clear

    '清空结果栏
    ResultEye.Clear
    sub调整结果表头格式_眼科
    ResultEyeArmy.Clear
    sub调整结果表头格式_眼科_涉核部队
    ResultEar.Clear
    sub调整结果表头格式_耳鼻喉科

    '恢复为form_load时的状态。
    If SSTPersonalInfo.Tab = 0 Then
        ctxtBarCode.Enabled = True
        ctxtBarCode.SetFocus
        ctxtName.Enabled = True
        ctxtSex.Enabled = True
        ctxtAge.Enabled = True
        ctxtCompanyName.Enabled = True
        ccmdUnfilledAllPass.Enabled = False
        ccmdAutoFill.Enabled = False
    Else
        ctxt体检条码.Enabled = True
        ctxt姓名.Enabled = True
        ctxt性别.Enabled = True
        ctxt年龄.Enabled = True
        ctxt单位名称.Enabled = True
    End If
    
    ccmdSave.Enabled = False
    ctlb工具栏.Buttons(2).Enabled = False
    ctlb工具栏.Buttons(3).Enabled = False
    ccmdClearResult.Enabled = False
    SSTResultIn.Enabled = False
    If ResultChanged <> 3 Then ResultChanged = -1
    coptClasses(0).Value = True
    coptClasses(0).Enabled = True: coptClasses(1).Enabled = True: coptClasses(2).Enabled = True
    '画图部分设定
    If coptClasses(0).Value = True Then Picture3.Picture = Nothing '仅放射工作和涉核部队有图片保存
    
    '2012-06-21 于登淼 ↓
    '初始化当前录入状态(提前判断有无权限修改，若无，直接赋值为3)
    ResultChanged = IIf(ResultChanged <> 3, -1, 3)
    cchk刷条码_Click
    '2012-06-21 于登淼 ↑
End Sub

Private Function sub添加单项结果(ByVal paraResult As String, ByVal paraItem As String, ByVal paraCheck As String) As String
    If paraResult = "" Then
        paraCheck = paraCheck & IIf(paraCheck = "", "", Chr(10)) & paraItem
    Else
        lcolItem.Add paraItem
        lcolResult.Add paraResult
    End If
    sub添加单项结果 = paraCheck
End Function

Sub subSave()
    On Error GoTo errHandler
    
    Dim isOk As Boolean
    Dim lstrTmp As String
    Dim lobjTmp As Object
    Set lobjTmp = CreateObject("职业病体检结果录入.clscommon")
    isOk = lobjTmp.func保存单人体检结果(ctxtBarCode.Text, mstrDoctorName, cdtpConclusionDate.Value, lcolItem, lcolResult, "职业病体检_结果信息_五官科")
    subClear
    If ResultChanged <> 3 Then ResultChanged = 1
    If isOk = True Then MsgBox ("保存成功！")
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病体检结果录入", "frmHEENT_ResultInput", "subSave", 6666, lstrError, False
End Sub

'没有包括涉核部队的结果填写，因为结果表格差别很大。
Sub sub填写已有的体检结果_眼科(ByVal paraRec As Object)
    Dim strArray
    Dim i, j, k, L, intTmp As Integer
    Dim hasSaved As Boolean
      
    For i = 1 To paraRec.recordcount
CONTINUE:
        If IsNull(paraRec("体检结果")) = True And paraRec.EOF = False Then
            paraRec.MoveNext
            GoTo CONTINUE
        ElseIf paraRec.EOF = True Then
            Exit Sub
        End If
        
        hasSaved = False
        strArray = Split(paraRec("项目名称"), "-", -1, vbBinaryCompare)
        L = UBound(strArray)
        If L = 0 Then   '只有普通职业病、放射的“其它”
            If paraRec("项目名称") = "眼科其它" Then
                For j = 2 To 10: ResultEye.TextMatrix(15, j) = paraRec("体检结果"): Next
                hasSaved = True
            End If
        End If
        
        If L = 2 And hasSaved = False Then   '只有普通职业病、放射的“裸眼”、“矫正”
            If strArray(0) = "裸眼" Then
                If (strArray(1) = "远视力" And strArray(2) = "右") Then ResultEye.TextMatrix(5, 3) = paraRec("体检结果"): hasSaved = True
                If (strArray(1) = "远视力" And strArray(2) = "左") Then ResultEye.TextMatrix(5, 8) = paraRec("体检结果"): hasSaved = True
                If (strArray(1) = "近视力" And strArray(2) = "右") Then ResultEye.TextMatrix(5, 5) = paraRec("体检结果"): hasSaved = True
                If (strArray(1) = "近视力" And strArray(2) = "左") Then ResultEye.TextMatrix(5, 10) = paraRec("体检结果"): hasSaved = True
            End If
            If strArray(0) = "矫正" Then
                If (strArray(1) = "远视力" And strArray(2) = "右") Then ResultEye.TextMatrix(6, 3) = paraRec("体检结果"): hasSaved = True
                If (strArray(1) = "远视力" And strArray(2) = "左") Then ResultEye.TextMatrix(6, 8) = paraRec("体检结果"): hasSaved = True
                If (strArray(1) = "近视力" And strArray(2) = "右") Then ResultEye.TextMatrix(6, 5) = paraRec("体检结果"): hasSaved = True
                If (strArray(1) = "近视力" And strArray(2) = "左") Then ResultEye.TextMatrix(6, 10) = paraRec("体检结果"): hasSaved = True
            End If
        End If
            
        If L = 1 And hasSaved = False Then
            For j = 2 To 14
                If ResultEye.TextMatrix(j, 1) = strArray(0) Then
                    If strArray(1) = "右" Then
                        For k = 2 To 5: ResultEye.TextMatrix(j, k) = paraRec("体检结果"): Next k
                    Else
                        For k = 7 To 10: ResultEye.TextMatrix(j, k) = paraRec("体检结果"): Next k
                    End If
                    hasSaved = True
                End If
            Next j
        End If
        paraRec.MoveNext
    Next i
End Sub



'-------以下为涉核部队所用的函数，有些操作与其他两个类似--------

Private Sub ResultEyeArmy_AfterEdit(ByVal row As Long, ByVal col As Long)
    sub调整结果表头格式_眼科_涉核部队
End Sub

Private Sub ResultEyeArmy_Click()
    If ResultChanged = 3 Then Exit Sub
    indX = ResultEyeArmy.MouseRow
    indY = ResultEyeArmy.MouseCol
    'MsgBox ("x = " & indX & ",y = " & indY)
    Dim i As Integer
    ResultEyeArmy.Editable = flexEDKbdMouse
    If indX < 0 Or indY < 0 Then
        Exit Sub
    ElseIf indX = 12 And indY >= 0 And indY <= 8 Then
        For i = 2 To 8: ResultEyeArmy.TextMatrix(indX, i) = "正常": Next
        If ResultChanged = 1 Then ResultChanged = 2
    ElseIf indX >= 2 And indX <= 5 And indY >= 2 And indY <= 4 And ResultEyeArmy.TextMatrix(indX, indY) = "" Then
        For i = 2 To 4: ResultEyeArmy.TextMatrix(indX, i) = "正常": Next
        If ResultChanged = 1 Then ResultChanged = 2
    ElseIf indX >= 2 And indX <= 5 And indY >= 6 And indY <= 8 And ResultEyeArmy.TextMatrix(indX, indY) = "" Then
        For i = 6 To 8: ResultEyeArmy.TextMatrix(indX, i) = "正常": Next
        If ResultChanged = 1 Then ResultChanged = 2
    ElseIf indX >= 7 And indX <= 11 Then
        ResultEyeArmy.TextMatrix(indX, indY) = "正常"
        If ResultChanged = 1 Then ResultChanged = 2
    End If
    sub调整结果表头格式_眼科_涉核部队
End Sub

Private Sub ResultEyeArmy_DblClick()
    If ResultChanged = 3 Then Exit Sub
    indX = ResultEyeArmy.MouseRow
    indY = ResultEyeArmy.MouseCol
    ResultEyeArmy.Editable = flexEDKbdMouse
    If indX < 0 Or indY < 0 Then
        Exit Sub
    ElseIf indX >= 0 And indX < ResultEyeArmy.rows And indY >= 0 And indY < ResultEyeArmy.cols Then
        sub录入修改内容 indX, indY
        If ResultChanged = 1 Then ResultChanged = 2
    End If
End Sub

Sub sub调整结果表头格式_眼科_涉核部队()
    Dim i As Integer
    
    '调整表格位置和大小，与职业健康、放射工作一样
    ResultEyeArmy.Height = ResultEye.Height
    ResultEyeArmy.Left = ResultEye.Left
    ResultEyeArmy.Top = ResultEye.Top
    ResultEyeArmy.Width = ResultEye.Width
    
    '调整列宽和字体位置
    'ResultEyeArmy.AutoSize 2, ResultEyeArmy.Cols - 1, 1, 650 '如果修改了内容，表格会由于不明原因拓宽。暂不用这句。
    ResultEyeArmy.ColWidth(0) = 960
    ResultEyeArmy.ColWidth(1) = 640
    For i = 0 To 8: ResultEyeArmy.ColAlignment(i) = flexAlignCenterCenter: Next
    
    '0~12行，0~1列范围内，值相同的要合并单元格
    ResultEyeArmy.MergeCompare = flexMCIncludeNulls
    ResultEyeArmy.MergeCells = flexMergeFree
    ResultEyeArmy.MergeRow(12) = True
    For i = 0 To 5: ResultEyeArmy.MergeRow(i) = True: Next
    For i = 0 To 1: ResultEyeArmy.MergeCol(i) = True: Next
    
    '0、1列，前两列格式设定
    ResultEyeArmy.TextMatrix(0, 0) = "项目": ResultEyeArmy.TextMatrix(0, 1) = "项目"
    ResultEyeArmy.TextMatrix(1, 0) = "眼别": ResultEyeArmy.TextMatrix(1, 1) = "眼别"
    ResultEyeArmy.TextMatrix(2, 0) = "视力": ResultEyeArmy.TextMatrix(3, 0) = "视力"
    ResultEyeArmy.TextMatrix(2, 1) = "裸眼": ResultEyeArmy.TextMatrix(3, 1) = "矫正"
    ResultEyeArmy.TextMatrix(4, 0) = "角膜": ResultEyeArmy.TextMatrix(4, 1) = "角膜"
    ResultEyeArmy.TextMatrix(5, 0) = "结膜": ResultEyeArmy.TextMatrix(5, 1) = "结膜"
    ResultEyeArmy.TextMatrix(6, 0) = "晶体裂隙灯" & Chr(10) & "检查所见": ResultEyeArmy.TextMatrix(7, 0) = "晶体裂隙灯" & Chr(10) & "检查所见": ResultEyeArmy.TextMatrix(8, 0) = "晶体裂隙灯" & Chr(10) & "检查所见": ResultEyeArmy.TextMatrix(9, 0) = "晶体裂隙灯" & Chr(10) & "检查所见": ResultEyeArmy.TextMatrix(10, 0) = "晶体裂隙灯" & Chr(10) & "检查所见": ResultEyeArmy.TextMatrix(11, 0) = "晶体裂隙灯" & Chr(10) & "检查所见"
    ResultEyeArmy.TextMatrix(7, 1) = "粉尘状": ResultEyeArmy.TextMatrix(8, 1) = "点状": ResultEyeArmy.TextMatrix(9, 1) = "片状": ResultEyeArmy.TextMatrix(10, 1) = "空泡": ResultEyeArmy.TextMatrix(11, 1) = "其它"
    ResultEyeArmy.TextMatrix(12, 0) = "眼科诊断": ResultEyeArmy.TextMatrix(12, 1) = "眼科诊断"
    
    '2~8列格式设定，中间列为第5列，值全为***，为了两列单元格不合并，默认隐藏。
    ResultEyeArmy.ColHidden(5) = True
    For i = 1 To 11: ResultEyeArmy.TextMatrix(i, 5) = "***": Next
    For i = 2 To 8: ResultEyeArmy.TextMatrix(0, i) = "检查结果": Next
    For i = 2 To 4: ResultEyeArmy.TextMatrix(1, i) = "右": Next
    For i = 6 To 8: ResultEyeArmy.TextMatrix(1, i) = "左": Next
    ResultEyeArmy.TextMatrix(6, 2) = "后囊下": ResultEyeArmy.TextMatrix(6, 3) = "前囊下": ResultEyeArmy.TextMatrix(6, 4) = "赤道"
    ResultEyeArmy.TextMatrix(6, 6) = "后囊下": ResultEyeArmy.TextMatrix(6, 7) = "前囊下": ResultEyeArmy.TextMatrix(6, 8) = "赤道"
End Sub

Sub sub填写已有的体检结果_眼科_涉核部队(ByVal paraRec As Object)
    Dim strArray
    Dim i, j, k, Upd, intTmp As Integer
    Dim hasSaved As Boolean
    
    For i = 1 To paraRec.recordcount
CONTINUE:
        If IsNull(paraRec("体检结果")) = True And paraRec.EOF = False Then
            paraRec.MoveNext
            GoTo CONTINUE
        ElseIf paraRec.EOF = True Then
            Exit Sub
        End If
        
        hasSaved = False
        strArray = Split(paraRec("项目名称"), "-", -1, vbBinaryCompare)
        Upd = UBound(strArray)
        If Upd = 0 Then   '只有普通职业病、放射的“其它”
            If paraRec("项目名称") = "眼科诊断" Then
                For j = 2 To 8: ResultEyeArmy.TextMatrix(12, j) = paraRec("体检结果"): Next
                hasSaved = True
            End If
        End If
        
        If Upd = 2 And hasSaved = False Then   '只有普通职业病、放射的“裸眼”、“矫正”
            For j = 7 To 11
                For k = 2 To 4
                    If ResultEyeArmy.TextMatrix(j, 1) = strArray(0) And ResultEyeArmy.TextMatrix(6, k) = strArray(1) Then
                        If strArray(2) = "右" Then
                            ResultEyeArmy.TextMatrix(j, k) = paraRec("体检结果")
                        Else
                            ResultEyeArmy.TextMatrix(j, k + 4) = paraRec("体检结果")
                        End If
                        hasSaved = True
                    End If
                Next k
            Next j
        End If
            
        If Upd = 1 And hasSaved = False Then
            For j = 2 To 5
                If ResultEyeArmy.TextMatrix(j, 1) = strArray(0) Then
                    If strArray(1) = "右" Then
                        For k = 2 To 4: ResultEyeArmy.TextMatrix(j, k) = paraRec("体检结果"): Next k
                    Else
                        For k = 6 To 8: ResultEyeArmy.TextMatrix(j, k) = paraRec("体检结果"): Next k
                    End If
                    hasSaved = True
                End If
            Next j
        End If
        paraRec.MoveNext
    Next i
End Sub


'-----------耳鼻喉科单独用函数
Private Sub ResultEar_AfterEdit(ByVal row As Long, ByVal col As Long)
    sub调整结果表头格式_耳鼻喉科
End Sub

Private Sub ResultEar_Click()
    If ResultChanged = 3 Then Exit Sub
    indX = ResultEar.MouseRow
    indY = ResultEar.MouseCol
    ResultEar.Editable = flexEDKbdMouse
    If indX < 0 Or indY < 0 Then Exit Sub
    ResultEar.TextMatrix(indX, ResultEar.cols - 1) = "正常"
    sub调整结果表头格式_耳鼻喉科
End Sub

Private Sub ResultEar_DblClick()
    If ResultChanged = 3 Then Exit Sub
    indX = ResultEar.MouseRow
    indY = ResultEar.MouseCol
    ResultEar.Editable = flexEDKbdMouse
    If indX < 0 Or indY < 0 Then
        Exit Sub
    ElseIf indX >= 0 And indX < ResultEar.rows And indY = ResultEar.cols - 1 Then
        sub录入修改内容 indX, indY
        If ResultChanged = 1 Then ResultChanged = 2
    End If
End Sub

Sub sub填写已有的体检结果_耳鼻喉科(ByVal paraRec As Object)
    Dim strArray
    Dim i, j, L As Integer
    
    paraRec.movefirst
    For j = 1 To paraRec.recordcount
CONTINUE:
        If IsNull(paraRec("体检结果")) = True And paraRec.EOF = False Then
            paraRec.MoveNext
            GoTo CONTINUE
        ElseIf paraRec.EOF = True Then
            Exit Sub
        End If
        
        strArray = Split(paraRec("项目名称"), "-", -1, vbBinaryCompare)
        L = UBound(strArray)
        If L = 0 Then
            For i = 1 To ResultEar.rows - 1
                If ResultEar.TextMatrix(i, 0) = strArray(0) Then ResultEar.TextMatrix(i, ResultEar.cols - 1) = paraRec("体检结果")
            Next i
        Else
            For i = 2 To ResultEar.rows - 1
                If ResultEar.TextMatrix(i, 0) = strArray(0) And ResultEar.TextMatrix(i, 1) = strArray(1) Then
                    ResultEar.TextMatrix(i, ResultEar.cols - 1) = paraRec("体检结果")
                End If
            Next i
        End If
        paraRec.MoveNext
    Next j
End Sub

Sub sub调整结果表头格式_耳鼻喉科()
    Dim i As Integer
    
    '调整表格位置和大小。
    ResultEar.Height = ResultEye.Height     '不知何种原因，显示时可以同时看到eye和ear的结果表格。所以干脆大小设为一样。
    ResultEar.Left = ResultEye.Left
    ResultEar.Top = ResultEye.Top
    ResultEar.Width = ResultEye.Width
    
    '调整列宽和字体位置
    'ResultEar.AutoSize 2, ResultEar.Cols - 1, 0, 0
    ResultEar.ColWidth(0) = 800
    ResultEar.ColWidth(1) = 500
    ResultEar.ColWidth(2) = ResultEar.Width - ResultEar.ColWidth(0) - ResultEar.ColWidth(1)
    ResultEar.AllowUserResizing = flexResizeColumns
    For i = 0 To 2
        ResultEar.ColAlignment(i) = flexAlignCenterCenter
    Next
    
    '设置需要合并的单元格
    ResultEar.MergeCompare = flexMCIncludeNulls
    ResultEar.MergeCells = flexMergeFree
    ResultEar.MergeCol(0) = True: ResultEar.MergeCol(1) = True
    For i = 0 To ResultEar.rows - 1: ResultEar.MergeRow(i) = True: Next
    
    '单独设置表头内容：
    '职业健康表头部分
    ResultEar.TextMatrix(4, 0) = "外耳": ResultEar.TextMatrix(4, 1) = "外耳"
    ResultEar.TextMatrix(5, 0) = "中耳": ResultEar.TextMatrix(5, 1) = "中耳"
    ResultEar.TextMatrix(6, 0) = "听力": ResultEar.TextMatrix(7, 0) = "听力"    '与涉核部队共用
    ResultEar.TextMatrix(6, 1) = "左": ResultEar.TextMatrix(7, 1) = "右"        '与涉核部队共用
    ResultEar.TextMatrix(3, 0) = "鼻": ResultEar.TextMatrix(3, 1) = "鼻"
    ResultEar.TextMatrix(14, 0) = "口腔": ResultEar.TextMatrix(15, 0) = "口腔"
    ResultEar.TextMatrix(14, 1) = "粘膜": ResultEar.TextMatrix(15, 1) = "牙齿"
    ResultEar.TextMatrix(16, 0) = "咽喉": ResultEar.TextMatrix(16, 1) = "咽喉"  '与涉核部队共用
    
    '放射工作表头部分
    ResultEar.TextMatrix(1, 0) = "听力": ResultEar.TextMatrix(1, 1) = "听力"
    ResultEar.TextMatrix(2, 0) = "嗅觉": ResultEar.TextMatrix(2, 1) = "嗅觉"
    
    '涉核部队表头部分
    ResultEar.TextMatrix(8, 0) = "外耳道": ResultEar.TextMatrix(9, 0) = "外耳道"
    ResultEar.TextMatrix(8, 1) = "左": ResultEar.TextMatrix(9, 1) = "右"
    ResultEar.TextMatrix(10, 0) = "乳突": ResultEar.TextMatrix(11, 0) = "乳突"
    ResultEar.TextMatrix(10, 1) = "左": ResultEar.TextMatrix(11, 1) = "右"
    ResultEar.TextMatrix(12, 0) = "鼻": ResultEar.TextMatrix(13, 0) = "鼻"
    ResultEar.TextMatrix(12, 1) = "粘膜": ResultEar.TextMatrix(13, 1) = "出血"
    ResultEar.TextMatrix(17, 0) = "口腔": ResultEar.TextMatrix(17, 1) = "口腔"
    
    '其它表头
    ResultEar.TextMatrix(0, 0) = "项目": ResultEar.TextMatrix(0, 1) = "项目"
    ResultEar.TextMatrix(0, 2) = "检查结果"
    ResultEar.TextMatrix(18, 0) = "耳鼻喉科其它": ResultEar.TextMatrix(18, 1) = "耳鼻喉科其它"
    
    '体检项目修正，不是该体检类别的行需要隐藏
    For i = 1 To ResultEar.rows - 1: ResultEar.RowHidden(i) = False: Next
    If coptClasses(0).Value = True Then
        ResultEar.RowHidden(1) = True
        ResultEar.RowHidden(2) = True
        ResultEar.RowHidden(8) = True
        ResultEar.RowHidden(9) = True
        ResultEar.RowHidden(10) = True
        ResultEar.RowHidden(11) = True
        ResultEar.RowHidden(12) = True
        ResultEar.RowHidden(13) = True
        ResultEar.RowHidden(17) = True
    ElseIf coptClasses(1).Value = True Then
        For i = 3 To 17: ResultEar.RowHidden(i) = True: Next
    Else
        For i = 1 To 5: ResultEar.RowHidden(i) = True: Next
        ResultEar.RowHidden(14) = True
        ResultEar.RowHidden(15) = True
        ResultEar.RowHidden(18) = True
    End If
End Sub


'-----------------------以下为画图控制部分-------------------------
Private Sub ccmdClearPicture_Click()
    Picture3.Cls
    Picture3.ForeColor = vbRed
    Picture3.DrawWidth = DrawLineWidth
    DrawState = -1
End Sub

Private Sub ccmdDraw_Click()
    Picture3.ForeColor = vbRed
    Picture3.DrawWidth = DrawLineWidth
    DrawState = 1
End Sub

Private Sub ccmdEraser_Click()
    Picture3.DrawWidth = DrawLineWidth
    DrawState = 2
End Sub

Private Sub Label9_Click(Index As Integer)
    DrawLineWidth = Pow_2(Index)
    Picture3.DrawWidth = DrawLineWidth
End Sub

Private Function Pow_2(ByVal paraExp As Integer) As Integer
    Dim i, resultTmp As Integer     'paraExp最好小于10,否则会溢出
    resultTmp = 1
    If paraExp > 10 Then paraExp = 10
    For i = 1 To paraExp: resultTmp = resultTmp * 2: Next
    Pow_2 = resultTmp
End Function

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CurrentX = X: CurrentY = Y
    If DrawState = -1 Then ccmdDraw_Click
End Sub

Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And DrawState = 1 Then
        Picture3.Line (CurrentX, CurrentY)-(X, Y)
    End If
    If Button = 1 And DrawState = 2 Then
        Picture3.ForeColor = vbWhite
        Picture3.MousePointer = 4
        Picture3.Line (CurrentX, CurrentY)-(X, Y)
    End If
    CurrentX = X
    CurrentY = Y
End Sub

'复杂度略高啊~~~~~~~6040像素
Private Sub Picture3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If DrawState <> 2 Then Exit Sub
    
    Dim i, xx, yy As Integer
    For i = 1 To pointCnt
        xx = EyeMapCheck(i, 1)
        yy = EyeMapCheck(i, 2)
        Call SetPixel(Picture3.hdc, xx, yy, 0)
    Next
    Picture3.Refresh
End Sub


Private Sub ccmdSavePicture_Click()
    Dim lobjTmp As Object
    Set lobjTmp = CreateObject("职业病体检结果录入.clscommon")
    lobjTmp.func保存结果图片 Picture3.Image, ctxtBarCode.Text, mstr结果图片项目编号, cdtpConclusionDate.Value    '01069 是眼睛体检结果图的项目编号。在“体检项目设置”的表里有记录。
    If DrawState = 0 Then
        Call sub添加单项结果("正常", "晶状体环面及正面图", "")
    ElseIf DrawState <> -1 Then
        Call sub添加单项结果("不正常", "晶状体环面及正面图", "")
    End If
    MsgBox ("结果图片保存成功！")
    DrawState = 0
    Exit Sub
End Sub

Private Sub ccmdLoadOriginalPicture_Click()
    Dim lobjTmp, lobjRec As Object
    Dim isOk As Integer
    Set lobjTmp = CreateObject("职业病体检结果录入.clscommon")
    Set lobjRec = lobjTmp.func查找结果图片(ctxtBarCode.Text, mstr结果图片项目编号)
    If lobjRec.recordcount > 0 Then
        isOk = MsgBox("载入原图会删除现有的结果，确定继续吗？", vbOKCancel)
        If isOk = 2 Then Exit Sub
        lobjTmp.func删除结果图片 ctxtBarCode.Text, mstr结果图片项目编号
        DrawState = 0
    End If
    Picture3.Picture = lobjTmp.func获取结果图片(ctxtBarCode.Text, mstr结果图片项目编号, "晶状体环面及正面图.bmp")
End Sub

Sub sub原图解析()
    Dim i, j, rows, cols As Integer
    Picture3.ScaleMode = 3
    
    Set Picture3.Picture = LoadPicture(App.Path & "\晶状体环面及正面图.bmp")
    cols = Picture3.ScaleWidth - 1
    rows = Picture3.ScaleHeight - 1

    pointCnt = 0
    For i = 1 To cols
        For j = 1 To rows
            'true时像素为黑色，false时像素为白色
            If Hex(GetPixel(Picture3.hdc, i, j)) = Hex(&H0) Then
                pointCnt = pointCnt + 1
                EyeMapCheck(pointCnt, 1) = i
                EyeMapCheck(pointCnt, 2) = j
            End If
        Next
    Next
    Set Picture3.Picture = Nothing

End Sub

'功能：批量提交体检人员的体检结果
'时间：2012-04-26
'作者：翁乔
Private Sub sub批量保存()
    Dim lstrCheck, lstrItem, lstrResult As String
    Dim i, j, isOk As Integer
    Dim lobjTmp As Object
    
On Error GoTo errHandler
    '录入结果界面暂时不能操作
    ccmdUnfilledAllPass.Enabled = False
    ccmdAutoFill.Enabled = False
    ccmdSave.Enabled = False
    ctlb工具栏.Buttons(3).Enabled = False
    ccmdClear.Enabled = False
    SSTResultIn.Enabled = False
    
    Set lcolResult = New Collection
    Set lcolItem = New Collection

    If coptClasses(2).Value = False Then        '接下来是眼科结果添加
        For i = 2 To 15
            If ResultEye.RowHidden(i) = False Then
                If Not (ResultEye.TextMatrix(i, 1) = "裸眼" Or ResultEye.TextMatrix(i, 1) = "矫正" Or ResultEye.TextMatrix(i, 1) = "眼科其它") Then
                    lstrCheck = sub添加单项结果(ResultEye.TextMatrix(i, 2), (ResultEye.TextMatrix(i, 1) & "-右"), lstrCheck)
                    lstrCheck = sub添加单项结果(ResultEye.TextMatrix(i, 7), (ResultEye.TextMatrix(i, 1) & "-左"), lstrCheck)
                ElseIf ResultEye.TextMatrix(i, 1) = "裸眼" Then
                    lstrCheck = sub添加单项结果(ResultEye.TextMatrix(i, 3), "裸眼-远视力-右", lstrCheck)
                    lstrCheck = sub添加单项结果(ResultEye.TextMatrix(i, 5), "裸眼-近视力-右", lstrCheck)
                    lstrCheck = sub添加单项结果(ResultEye.TextMatrix(i, 8), "裸眼-远视力-左", lstrCheck)
                    lstrCheck = sub添加单项结果(ResultEye.TextMatrix(i, 10), "裸眼-近视力-左", lstrCheck)
                ElseIf ResultEye.TextMatrix(i, 1) = "矫正" Then
                    lstrCheck = sub添加单项结果(ResultEye.TextMatrix(i, 3), "矫正-远视力-右", lstrCheck)
                    lstrCheck = sub添加单项结果(ResultEye.TextMatrix(i, 5), "矫正-近视力-右", lstrCheck)
                    lstrCheck = sub添加单项结果(ResultEye.TextMatrix(i, 8), "矫正-远视力-左", lstrCheck)
                    lstrCheck = sub添加单项结果(ResultEye.TextMatrix(i, 10), "矫正-近视力-左", lstrCheck)
                ElseIf ResultEye.TextMatrix(i, 1) = "眼科其它" Then
                    lstrCheck = sub添加单项结果(ResultEye.TextMatrix(i, 2), "眼科其它", lstrCheck)
                End If
            End If
        Next i
    Else
        '添加涉核部队的单项结果，可以用公共函数 sub添加单项结果 和 subSave
            
        '仅有左右眼别的项目结果
        For i = 2 To 5
            lstrCheck = sub添加单项结果(ResultEyeArmy.TextMatrix(i, 2), (ResultEyeArmy.TextMatrix(i, 1) & "-右"), lstrCheck)
            lstrCheck = sub添加单项结果(ResultEyeArmy.TextMatrix(i, 6), (ResultEyeArmy.TextMatrix(i, 1) & "-左"), lstrCheck)
        Next i
            
        '左右眼别下，分别都有3项检查项目的结果填写
        For i = 7 To 11
            For j = 2 To 8
                If j <= 4 Then lstrCheck = sub添加单项结果(ResultEyeArmy.TextMatrix(i, j), (ResultEyeArmy.TextMatrix(i, 1) & "-" & ResultEyeArmy.TextMatrix(6, j) & "-右"), lstrCheck)
                If j >= 6 Then lstrCheck = sub添加单项结果(ResultEyeArmy.TextMatrix(i, j), (ResultEyeArmy.TextMatrix(i, 1) & "-" & ResultEyeArmy.TextMatrix(6, j) & "-左"), lstrCheck)
            Next j
        Next i
            
        '"诊断"的结果填写
        lstrCheck = sub添加单项结果(ResultEyeArmy.TextMatrix(12, 2), "眼科诊断", lstrCheck)
    End If
        
    For i = 1 To 18
        If ResultEar.RowHidden(i) = False Then
            If i <= 5 Or i >= 16 Then
                lstrCheck = sub添加单项结果(ResultEar.TextMatrix(i, ResultEar.cols - 1), ResultEar.TextMatrix(i, 0), lstrCheck)
            Else
                lstrCheck = sub添加单项结果(ResultEar.TextMatrix(i, ResultEar.cols - 1), ResultEar.TextMatrix(i, 0) & "-" & ResultEar.TextMatrix(i, 1), lstrCheck)
            End If
        End If
    Next
    
     'lstrcheck字符串检查
    If (Not lstrCheck = "") And (Not ResultChanged = 2) Then
        isOk = MsgBox("以下项目未填写结果，确定保存吗？" & Chr(10) & "未填写项将不会记录到数据库！" & Chr(10) & Chr(10) & Trim(lstrCheck), vbOKCancel)
        If isOk = 2 Then
            Set lcolResult = Nothing
            Set lcolItem = Nothing
            ccmdUnfilledAllPass.Enabled = True
            ccmdAutoFill.Enabled = True
            ccmdSave.Enabled = True
            ctlb工具栏.Buttons(2).Enabled = True
            ccmdClearResult.Enabled = True
            Exit Sub
        End If
    End If
        
    If ResultChanged = 2 Then
        isOk = MsgBox("是否保存该体检人员的修改结果？", vbOKCancel)
        If isOk = 1 Then
            subSaveBatch         '里面包含保存成功提示
        Else
            LoadPersonalInfoBatch (ctxt体检条码)
        End If
    Else
        subSaveBatch
    End If
    
    Set lcolResult = Nothing
    Set lcolItem = Nothing
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病体检结果录入", "frmHEENT_ResultInput", "ccmdSave_Click", 6666, lstrError, False
    
End Sub
'批量保存到数据库
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
    Dim lcolConclusion As String '五官科的体检结论
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
        isOk = lobjTmp.func保存单人体检结果(barCode(i), mstrDoctorName, DTP录入日期.Value, lcolItem, lcolResult, "职业病体检_结果信息_五官科")
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
            '01069 是眼睛体检结果图的项目编号。在“体检项目设置”的表里有记录。
            lobjTmp.func保存结果图片 Picture3.Image, barCode(i), mstr结果图片项目编号, cdtpConclusionDate.Value
            If DrawState = 0 Then
                Call sub添加单项结果("正常", "晶状体环面及正面图", "")
            ElseIf DrawState <> -1 Then
                Call sub添加单项结果("不正常", "晶状体环面及正面图", "")
            End If
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
        subClear
    Else
        subClear
    End If
        ccrp进度.Visible = False
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病体检结果录入", "frmHEENT_ResultInput", "subSave", 6666, lstrError, False
End Sub

'2012-07-03 于登淼
Sub sub获取系统编号固定部分()
    '获取服务器日期
    Dim lobjRec As Object
    Set lobjRec = dafuncGetData("select getDate()")
    ctxtBarCode.Text = um防疫站编号 & um服务器代号 & Format(lobjRec(0), "yyyy")
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
    
        If SSTPersonalInfo.Tab = 0 Then
            If cchkSigResult(0).Value = 1 And cchkSigResult(1).Value = 0 Then
                mobjQueryResult.Filter = "填写时间<>null"
            ElseIf cchkSigResult(0).Value = 0 And cchkSigResult(1).Value = 1 Then
                mobjQueryResult.Filter = "填写时间=null"
            ElseIf cchkSigResult(0).Value = 0 And cchkSigResult(1).Value = 0 Then
                mobjQueryResult.Filter = "系统编号='xxx'"
            Else
                mobjQueryResult.Filter = ""
            End If
        ElseIf SSTPersonalInfo.Tab = 1 Then
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
    
    If SSTPersonalInfo.Tab = 0 Then
        With cgrdinfo
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

End Sub
