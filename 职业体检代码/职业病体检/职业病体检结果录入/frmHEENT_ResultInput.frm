VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C8530F8A-C19C-11D2-99D6-9419F37DBB29}#1.1#0"; "ccrpprg6.ocx"
Begin VB.Form frmHEENT_ResultInput 
   Caption         =   "��ٿƽ��¼�봰��"
   ClientHeight    =   9420
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14430
   LinkTopic       =   "Form1"
   ScaleHeight     =   9420
   ScaleWidth      =   14430
   StartUpPosition =   2  '��Ļ����
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
         Begin VB.CheckBox cchkˢ���� 
            BackColor       =   &H00C0FFFF&
            Caption         =   "ˢ����"
            Height          =   180
            Left            =   2160
            TabIndex        =   86
            Top             =   1200
            Width           =   975
         End
         Begin VB.OptionButton coptClasses 
            BackColor       =   &H00C0FFFF&
            Caption         =   "���佡��"
            Height          =   255
            Index           =   2
            Left            =   2400
            TabIndex        =   84
            Top             =   840
            Width           =   1095
         End
         Begin VB.OptionButton coptClasses 
            BackColor       =   &H00C0FFFF&
            Caption         =   "ְҵ����"
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   83
            Top             =   840
            Width           =   1095
         End
         Begin VB.OptionButton coptClasses 
            BackColor       =   &H00C0FFFF&
            Caption         =   "��ͨ���"
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
            Caption         =   "��˲���"
            Height          =   255
            Index           =   3
            Left            =   3480
            TabIndex        =   81
            Top             =   840
            Width           =   1095
         End
         Begin VB.OptionButton coptClasses 
            BackColor       =   &H00C0FFFF&
            Caption         =   "8023����"
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
            TabCaption(0)   =   "�ۿ�"
            TabPicture(0)   =   "frmHEENT_ResultInput.frx":0000
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "fraDraw"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "Frame3"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).ControlCount=   2
            TabCaption(1)   =   "���Ǻ��"
            TabPicture(1)   =   "frmHEENT_ResultInput.frx":001C
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Frame2"
            Tab(1).Control(1)=   "ctxtConclun"
            Tab(1).Control(2)=   "Cmd����ģ��"
            Tab(1).Control(3)=   "Chk����ģ��"
            Tab(1).ControlCount=   4
            Begin VB.CheckBox Chk����ģ�� 
               Caption         =   "����ģ��"
               Height          =   255
               Left            =   -69360
               TabIndex        =   55
               Top             =   5280
               Value           =   1  'Checked
               Width           =   1095
            End
            Begin VB.CommandButton Cmd����ģ�� 
               Caption         =   "����ģ��"
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
                  ToolTipText     =   "�������޸�"
                  Top             =   240
                  Width           =   3735
                  _cx             =   2088769980
                  _cy             =   2088767017
                  Appearance      =   1
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
               Caption         =   "���廷�漰����ͼ(������ɫĬ�Ϻ�ɫ)"
               Height          =   3615
               Left            =   120
               TabIndex        =   3
               Top             =   4440
               Width           =   7815
               Begin VB.CommandButton ccmdLoadOriginalPicture 
                  Caption         =   "����ԭͼ"
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
                  Caption         =   "����ͼ��"
                  Height          =   255
                  Left            =   2160
                  TabIndex        =   7
                  Top             =   840
                  Visible         =   0   'False
                  Width           =   1215
               End
               Begin VB.CommandButton ccmdClearPicture 
                  Caption         =   "��մ˴��޸�"
                  Height          =   375
                  Left            =   5280
                  TabIndex        =   6
                  Top             =   240
                  Width           =   1335
               End
               Begin VB.CommandButton ccmdEraser 
                  Caption         =   "��Ƥ��"
                  Height          =   375
                  Left            =   4200
                  TabIndex        =   5
                  Top             =   720
                  Width           =   855
               End
               Begin VB.CommandButton ccmdDraw 
                  Caption         =   "��ͼ"
                  Height          =   375
                  Left            =   4200
                  TabIndex        =   4
                  Top             =   240
                  Width           =   855
               End
               Begin VB.Label Label9 
                  BackColor       =   &H00FFC0C0&
                  Caption         =   "���"
                  Height          =   375
                  Index           =   3
                  Left            =   3600
                  TabIndex        =   16
                  Top             =   360
                  Width           =   495
               End
               Begin VB.Label Label9 
                  BackColor       =   &H00FFC0C0&
                  Caption         =   "��"
                  Height          =   375
                  Index           =   2
                  Left            =   3000
                  TabIndex        =   15
                  Top             =   360
                  Width           =   495
               End
               Begin VB.Label Label9 
                  BackColor       =   &H00FFC0C0&
                  Caption         =   "��"
                  Height          =   375
                  Index           =   1
                  Left            =   2400
                  TabIndex        =   14
                  Top             =   360
                  Width           =   495
               End
               Begin VB.Label Label13 
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "����"
                  Height          =   255
                  Left            =   6360
                  TabIndex        =   13
                  Top             =   840
                  Width           =   375
               End
               Begin VB.Label Label12 
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "����"
                  Height          =   255
                  Left            =   960
                  TabIndex        =   12
                  Top             =   840
                  Width           =   375
               End
               Begin VB.Label Label9 
                  BackColor       =   &H00FFC0C0&
                  Caption         =   "ϸ"
                  Height          =   375
                  Index           =   0
                  Left            =   1800
                  TabIndex        =   11
                  Top             =   360
                  Width           =   495
               End
               Begin VB.Label Label8 
                  BackColor       =   &H00FFC0C0&
                  Caption         =   "����/��Ƥ����ϸ��"
                  Height          =   375
                  Left            =   120
                  TabIndex        =   10
                  Top             =   360
                  Width           =   1575
               End
            End
            Begin VB.Label Label10 
               Caption         =   "��ٿƽ��ۣ�"
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
            TabCaption(0)   =   "����¼��"
            TabPicture(0)   =   "frmHEENT_ResultInput.frx":0038
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "fraInfo"
            Tab(0).Control(1)=   "fraQuery"
            Tab(0).ControlCount=   2
            TabCaption(1)   =   "����¼��"
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
            Tab(1).Control(7)=   "ccmd��ѯ��λ"
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
            Tab(1).Control(13)=   "ccrp����"
            Tab(1).Control(13).Enabled=   0   'False
            Tab(1).Control(14)=   "cchkBchResult(0)"
            Tab(1).Control(14).Enabled=   0   'False
            Tab(1).Control(15)=   "cchkBchResult(1)"
            Tab(1).Control(15).Enabled=   0   'False
            Tab(1).ControlCount=   16
            Begin VB.CheckBox cchkBchResult 
               Caption         =   "δ����"
               Height          =   180
               Index           =   1
               Left            =   2760
               TabIndex        =   92
               Top             =   4320
               Value           =   1  'Checked
               Width           =   1095
            End
            Begin VB.CheckBox cchkBchResult 
               Caption         =   "������"
               Height          =   180
               Index           =   0
               Left            =   1440
               TabIndex        =   91
               Top             =   4320
               Width           =   1095
            End
            Begin CCRProgressBar6.ccrpProgressBar ccrp���� 
               Height          =   375
               Left            =   120
               Top             =   4560
               Visible         =   0   'False
               Width           =   5535
               _ExtentX        =   9763
               _ExtentY        =   661
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
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
               FormatString    =   "���������"
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
               Caption         =   "�� ѯ"
               Height          =   375
               Left            =   4920
               TabIndex        =   75
               Top             =   3960
               Width           =   735
            End
            Begin VB.CheckBox cchkCompanyBatch 
               BackColor       =   &H00C0FFC0&
               Caption         =   "��λ����"
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
            Begin VB.CommandButton ccmd��ѯ��λ 
               Caption         =   "��λ��λ"
               Height          =   375
               Left            =   3960
               TabIndex        =   72
               Top             =   3960
               Width           =   855
            End
            Begin VB.Frame fraQueryBatch 
               Caption         =   "������ѯ�����Ա"
               Height          =   2895
               Left            =   120
               TabIndex        =   58
               Top             =   360
               Width           =   5535
               Begin VB.CheckBox cchk��������� 
                  BackColor       =   &H008080FF&
                  Caption         =   "�������Ա�����Ϊ���������¼��"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   77
                  Top             =   2520
                  Value           =   1  'Checked
                  Width           =   3615
               End
               Begin VB.TextBox ctxt������� 
                  Height          =   300
                  Left            =   1320
                  TabIndex        =   64
                  Top             =   720
                  Width           =   2415
               End
               Begin VB.TextBox ctxt���� 
                  Height          =   300
                  Left            =   1320
                  TabIndex        =   63
                  Top             =   1080
                  Width           =   2415
               End
               Begin VB.TextBox ctxt�Ա� 
                  Height          =   300
                  Left            =   1320
                  TabIndex        =   62
                  Top             =   1440
                  Width           =   2415
               End
               Begin VB.TextBox ctxt���� 
                  Height          =   300
                  Left            =   1320
                  TabIndex        =   61
                  Top             =   1800
                  Width           =   2415
               End
               Begin VB.TextBox ctxt��λ���� 
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
               Begin MSComCtl2.DTPicker DTP¼������ 
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
                  Caption         =   "��������"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   71
                  Top             =   720
                  Width           =   975
               End
               Begin VB.Label Label17 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "����"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   70
                  Top             =   1080
                  Width           =   975
               End
               Begin VB.Label Label16 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "�Ա�"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   69
                  Top             =   1440
                  Width           =   975
               End
               Begin VB.Label Label15 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "����"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   68
                  Top             =   1800
                  Width           =   975
               End
               Begin VB.Label Label14 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "��λ����"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   67
                  Top             =   2160
                  Width           =   975
               End
               Begin VB.Label Label11 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "����¼������"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   66
                  Top             =   360
                  Width           =   1095
               End
            End
            Begin VB.CheckBox cchkDateBatch 
               BackColor       =   &H00C0FFC0&
               Caption         =   "�������"
               Height          =   255
               Left            =   120
               TabIndex        =   56
               Top             =   3480
               Width           =   1215
            End
            Begin VB.CommandButton ccmdClear 
               Caption         =   "�� ��"
               Height          =   375
               Left            =   3960
               TabIndex        =   48
               Top             =   3360
               Width           =   855
            End
            Begin VB.CommandButton ccmdRemove 
               Caption         =   "�� ��"
               Height          =   375
               Left            =   4920
               TabIndex        =   47
               Top             =   3360
               Width           =   735
            End
            Begin VB.Frame fraQuery 
               Caption         =   "��ѯ�����Ա"
               Height          =   4455
               Left            =   -74880
               TabIndex        =   40
               Top             =   3000
               Width           =   5535
               Begin VB.CheckBox cchkSigResult 
                  Caption         =   "δ����"
                  Height          =   255
                  Index           =   1
                  Left            =   1680
                  TabIndex        =   90
                  Top             =   720
                  Value           =   1  'Checked
                  Width           =   1095
               End
               Begin VB.CheckBox cchkSigResult 
                  Caption         =   "������"
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   89
                  Top             =   720
                  Width           =   1095
               End
               Begin VB.CommandButton ccmdQuery 
                  Caption         =   "��   ѯ"
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
                  Caption         =   "����"
                  Height          =   255
                  Left            =   3120
                  TabIndex        =   42
                  Top             =   240
                  Width           =   735
               End
               Begin VB.CheckBox cchkDate 
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "�������"
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
                  ToolTipText     =   "˫���Զ����������Ϣ�����������"
                  Top             =   1200
                  Width           =   5295
                  _cx             =   2088772732
                  _cy             =   2088768922
                  Appearance      =   1
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
                  FormatString    =   "���������"
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
                  Caption         =   "������"
                  Height          =   180
                  Left            =   4560
                  TabIndex        =   87
                  Top             =   840
                  Width           =   540
               End
            End
            Begin VB.Frame fraInfo 
               Caption         =   "������Ϣ"
               Height          =   2535
               Left            =   -74880
               TabIndex        =   25
               Top             =   360
               Width           =   5535
               Begin VB.CommandButton ccmdLocate 
                  Caption         =   "��λ��λ"
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
                  Caption         =   "����¼������"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   39
                  Top             =   360
                  Width           =   1095
               End
               Begin VB.Label Label5 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "��λ����"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   38
                  Top             =   2160
                  Width           =   975
               End
               Begin VB.Label Label4 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "����"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   37
                  Top             =   1800
                  Width           =   975
               End
               Begin VB.Label Label3 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "�Ա�"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   36
                  Top             =   1440
                  Width           =   975
               End
               Begin VB.Label Label2 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "����"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   35
                  Top             =   1080
                  Width           =   975
               End
               Begin VB.Label Label1 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "��������"
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
               Caption         =   "������"
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
         Begin MSComctlLib.Toolbar ctlb������ 
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
            Begin VB.CommandButton ccmdUnfilledAllPass 
               Caption         =   "��ǰδ��д��ȫ������"
               Height          =   375
               Left            =   6000
               TabIndex        =   53
               Top             =   120
               Width           =   2055
            End
            Begin VB.CommandButton ccmdAutoFill 
               Caption         =   "ȫ������"
               Height          =   375
               Left            =   12720
               TabIndex        =   52
               Top             =   120
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.CommandButton ccmdClearResult 
               Caption         =   "��յ�ǰ���"
               Height          =   375
               Left            =   10200
               TabIndex        =   51
               Top             =   120
               Width           =   1455
            End
            Begin VB.CommandButton ccmdSave 
               Caption         =   "����ȫ�����"
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
            Caption         =   "ҽ����"
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
'2012-03-01 �ڵ��
'���� ��ٿƽ��¼�봰�壬����Ӧ��������

Option Explicit
Public mblnInUse As Boolean
Private WithEvents mobjGUI As cls����ͨ�ö���
Attribute mobjGUI.VB_VarHelpID = -1
Private mstr��쵥�� As String
Private mstrϵͳ��� As String
Private mlobjRec As Object

'��ѯ���
Private mstrDoctorName As String
Private mobjQueryResult As Object
Private mcolIndex As New Collection
Private indX, indY As Integer       '��¼�����vsflexgrid�����ꡣ
Private lcolResult As Collection    '��������ϣ�item:[�����Ŀ���ƣ������]��
Private lcolItem As Collection      '���������Ŀ���������[�����Ŀ���ƣ������]��

'2012-07-14 �ڵ�� ��
'���ӿ��һ�����Ϣ����
Private priDeptName As String
Private priDeptNo As String
Private priDeptResultName As String
'2012-07-14 �ڵ�� ��

'��¼�ڵ�һ�α��������֮������ٴ��޸Ľ������Ҫ������������޸ģ��Ƿ񱣴桱֮�����ʾ��
'-1����ʾδ��ȡ�������ݿ������������Ϣ��
'0����ʾ���˵Ľ��δ¼�����
'1����ʾ���ݿ������и��˵Ľ�������ڽ�����δ���޸Ĺ���
'2����ʾ���ݿ������и��˵Ľ�������������޸Ĺ���ֻ����Ϊ2��ʱ�򣬲Żᵯ����������޸ģ��Ƿ񱣴桱����
'3����ʾû��Ȩ�޽����޸Ĳ�����
Private ResultChanged As Integer

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private DrawLineWidth As Integer        '��ͼ������ϸ�ı��
Private DrawState As Integer            '��¼��ǰ��ͼ״̬��δ���޸Ĳ����� -1��δ��ͼ0����ͼ1����Ƥ��2
Private mstr���ͼƬ��Ŀ��� As String  '�����ٿƽ��ͼƬ��Ŀ��ţ����ݿ��м�¼��ֵΪ��01069��

Private lobj������������ As Object    'Ϊ���������ṩ������
Private EyeMapCheck(6050, 2) As Integer
Private pointCnt As Long

'���ܣ����ص�ǰ�����Ƿ��Ѿ����ر�־������ϵͳƽ̨��Ҫ��ġ�
Public Property Get pblnInUse() As Boolean
    On Error Resume Next
    pblnInUse = mblnInUse
End Property


'2012-07-14 �ڵ��
Private Sub cchkBchResult_Click(Index As Integer)
    Dim i, coptIndex As Integer
    For i = 0 To coptClasses.Count - 1
        If coptClasses(i).Value = True Then coptIndex = i
    Next
    sub��ѯ�б���ʾ coptIndex
End Sub

'2012-07-14 �ڵ��
Private Sub cchkSigResult_Click(Index As Integer)
    Dim i, coptIndex As Integer
    For i = 0 To coptClasses.Count - 1
        If coptClasses(i).Value = True Then coptIndex = i
    Next
    sub��ѯ�б���ʾ coptIndex
End Sub

Private Sub cchkˢ����_Click()
    If Not cchkˢ����.Visible Then Exit Sub
    
    If SSTPersonalInfo.Tab = 0 Then
        ctxtBarCode.Text = ""
        If cchkˢ����.Value = 0 Then sub��ȡϵͳ��Ź̶�����
        ctxtBarCode.Enabled = True
        ctxtBarCode.SetFocus
        ctxtBarCode.SelStart = Len(ctxtBarCode)
        ctxtBarCode.SelLength = 0
    Else
        ctxt�������.Text = ""
        ctxt�������.Enabled = True
        ctxt�������.SetFocus
    End If
End Sub

Private Sub ccmdAutoFill_Click()
    On Error GoTo errHandler
    ccmdUnfilledAllPass_Click
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�������¼��", "frmHEENT_ResultInput", "ccmdAutoFill_Click", 6666, lstrError, False
End Sub

Private Sub ccmdClear_Click()
    cgrdInfoBatch.Clear
    cgrdInfoBatch.rows = 1
    cgrdInfoBatch.FormatString = "���������"
    TotalPeopleBatch.Caption = 0
End Sub

Private Sub ccmdClearResult_Click()
    If SSTResultIn.Tab = 0 Then
        If coptClasses(2).Value = False Then
            ResultEye.Clear
            sub���������ͷ��ʽ_�ۿ�
        Else
            ResultEyeArmy.Clear
            sub���������ͷ��ʽ_�ۿ�_��˲���
        End If
    ElseIf SSTResultIn.Tab = 1 Then
        ResultEar.Clear
        sub���������ͷ��ʽ_���Ǻ��
    End If
End Sub

'ccmdLocate��ʱ���ص���
Private Sub ccmdLocate_Click()
    On Error GoTo errHandler
    Dim lobjRec As Object                       '��λ��λ���صĽ����¼��
    Set lobjRec = pobjҵ�����.func��λ��λ     '������λ��λ���档
    
    '��ȡ��λ�ĵ�λ����ʾ�ڡ���λ���ơ�¼����С�(��ʱֻ��ʾ����λ���ơ�)
    '-----��֪�������費��Ҫ������ģ�������趨��˲��ӡ�
    If Not lobjRec Is Nothing Then
        If lobjRec.recordcount > 0 Then
            ctxtCompanyName.Text = IIf(IsNull(lobjRec("��λ����")), "", lobjRec("��λ����"))
        End If
    End If
    Set lobjRec = Nothing
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�������¼��", "frmHEENT_ResultInput", "ccmdLocate_Click", 6666, lstrError, False
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
    lstrWhere = " and �������='" & coptClasses(coptIndex).Caption & "'"
        
    '��װ��ѯ����
    If cchkDate.Value = 1 Then
        lstrWhere = lstrWhere & " and �������>='" & Format(cdtpDate.Value, "yyyy-mm-dd 00:00:00") & "' and �������<='" & Format(cdtpDate.Value, "yyyy-mm-dd 23:59:59") & "'"
    End If
    
    If cchkName.Value = 1 Then
        If ctxtCheckName.Text = "" Then
            MsgBox ("��Ҫ��ѯ����������������Ϊ�ա�")
            Exit Sub
        End If
        lstrWhere = lstrWhere & " and ����='" & Trim(ctxtCheckName.Text) & "'"
    End If
    
    '2012-07-14 �ڵ�� ��
    '���Ĳ�ѯ����������8/48Сʱ�ж����ݡ������޸�ʱ���ʼ�ղ������ѯ����С�
    '��ѯ���ݱ�����ݷ����ϴ�仯�����޸ģ������⡣

    '���ÿ������������������Ա�޸�ʱ�����¸��¡���������Ϣ���С��������״̬����'2'��Ϊ'3'�ģ���ѯʱ���ԡ�
    sub���¿��޸Ľ����Ա�޸�״̬
    
    Set lobjTmp = CreateObject("ְҵ�������¼��.clscommon")
    Set mobjQueryResult = lobjTmp.func��ȡ���޸Ľ���_������_�����Ա��Ϣ(lstrWhere, priDeptName)
    
    sub��ѯ�б���ʾ coptIndex
    '2012-07-14 �ڵ�� ��
    
    Set lobjTmp = Nothing
    Set lobjRec = Nothing
    lstrWhere = ""
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�������¼��", "frmHEENT_ResultInput", "ccmdQuery_Click", 6666, lstrError, False
End Sub

'2012-07-13 �ڵ��
'�޸�֮ǰ������ӵ��Ƴ�����������ctrl�������Ƴ�
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

'�ܵı��溯������ʱ���趨�ǣ����������ť���ۿ�(����ͼƬ)�Ͷ��Ǻ�ƵĽ��һ�𱣴����ʾ��
Private Sub ccmdSave_Click()
    On Error GoTo errHandler
    
    Dim lstrCheck, lstrItem, lstrResult As String
    Dim i, j, isOk As Integer
    Dim lobjTmp As Object
    Dim lcolConclusion As String '��ٿƵ�������
    
    '¼����������ʱ���ܲ���
    ccmdUnfilledAllPass.Enabled = False
    ccmdAutoFill.Enabled = False
    ccmdSave.Enabled = False
    ctlb������.Buttons(2).Enabled = False
    ccmdClear.Enabled = False
    SSTResultIn.Enabled = False
    
    Set lcolResult = New Collection
    Set lcolItem = New Collection
    
    '�ٱ������еĽ������
    If SSTPersonalInfo.Tab = 0 Then                 '��ʱΪ����¼��
    
        '���浥����Ŀ��ҽ������
        lcolConclusion = ctxtConclun.Text
        pobjҵ�����.sub������д������ ctxtBarCode.Text, priDeptName, lcolConclusion, um�û����
        
        '�ȱ�����ͼƬ�������乤������˲��ӵ��ۿ���ͼƬ�豣��
        If coptClasses(0).Value = False And DrawState <> -1 Then ccmdSavePicture_Click
        'If SSTResultIn.Tab = 1 Then GoTo ENTFill    '��ת�����Ǻ�ƽ����Ӳ���
        If coptClasses(2).Value = False Then        '���������ۿƽ�����
            For i = 2 To 15
                If ResultEye.RowHidden(i) = False Then
                    If Not (ResultEye.TextMatrix(i, 1) = "����" Or ResultEye.TextMatrix(i, 1) = "����" Or ResultEye.TextMatrix(i, 1) = "�ۿ�����") Then
                        lstrCheck = sub��ӵ�����(ResultEye.TextMatrix(i, 2), (ResultEye.TextMatrix(i, 1) & "-��"), lstrCheck)
                        lstrCheck = sub��ӵ�����(ResultEye.TextMatrix(i, 7), (ResultEye.TextMatrix(i, 1) & "-��"), lstrCheck)
                    ElseIf ResultEye.TextMatrix(i, 1) = "����" Then
                        lstrCheck = sub��ӵ�����(ResultEye.TextMatrix(i, 3), "����-Զ����-��", lstrCheck)
                        lstrCheck = sub��ӵ�����(ResultEye.TextMatrix(i, 5), "����-������-��", lstrCheck)
                        lstrCheck = sub��ӵ�����(ResultEye.TextMatrix(i, 8), "����-Զ����-��", lstrCheck)
                        lstrCheck = sub��ӵ�����(ResultEye.TextMatrix(i, 10), "����-������-��", lstrCheck)
                    ElseIf ResultEye.TextMatrix(i, 1) = "����" Then
                        lstrCheck = sub��ӵ�����(ResultEye.TextMatrix(i, 3), "����-Զ����-��", lstrCheck)
                        lstrCheck = sub��ӵ�����(ResultEye.TextMatrix(i, 5), "����-������-��", lstrCheck)
                        lstrCheck = sub��ӵ�����(ResultEye.TextMatrix(i, 8), "����-Զ����-��", lstrCheck)
                        lstrCheck = sub��ӵ�����(ResultEye.TextMatrix(i, 10), "����-������-��", lstrCheck)
                    ElseIf ResultEye.TextMatrix(i, 1) = "�ۿ�����" Then
                        lstrCheck = sub��ӵ�����(ResultEye.TextMatrix(i, 2), "�ۿ�����", lstrCheck)
                    End If
                End If
            Next i
        Else
            '�����˲��ӵĵ������������ù������� sub��ӵ����� �� subSave
            
            '���������۱����Ŀ���
            For i = 2 To 5
                lstrCheck = sub��ӵ�����(ResultEyeArmy.TextMatrix(i, 2), (ResultEyeArmy.TextMatrix(i, 1) & "-��"), lstrCheck)
                lstrCheck = sub��ӵ�����(ResultEyeArmy.TextMatrix(i, 6), (ResultEyeArmy.TextMatrix(i, 1) & "-��"), lstrCheck)
            Next i
            
            '�����۱��£��ֱ���3������Ŀ�Ľ����д
            For i = 7 To 11
                For j = 2 To 8
                    If j <= 4 Then lstrCheck = sub��ӵ�����(ResultEyeArmy.TextMatrix(i, j), (ResultEyeArmy.TextMatrix(i, 1) & "-" & ResultEyeArmy.TextMatrix(6, j) & "-��"), lstrCheck)
                    If j >= 6 Then lstrCheck = sub��ӵ�����(ResultEyeArmy.TextMatrix(i, j), (ResultEyeArmy.TextMatrix(i, 1) & "-" & ResultEyeArmy.TextMatrix(6, j) & "-��"), lstrCheck)
                Next j
            Next i
            
            '"���"�Ľ����д
            lstrCheck = sub��ӵ�����(ResultEyeArmy.TextMatrix(12, 2), "�ۿ����", lstrCheck)
        End If
        
'ENTFill:
        'If SSTResultIn.Tab = 1 Then
            For i = 1 To 18
                If ResultEar.RowHidden(i) = False Then
                    If i <= 5 Or i >= 16 Then
                        lstrCheck = sub��ӵ�����(ResultEar.TextMatrix(i, ResultEar.cols - 1), ResultEar.TextMatrix(i, 0), lstrCheck)
                    Else
                        lstrCheck = sub��ӵ�����(ResultEar.TextMatrix(i, ResultEar.cols - 1), ResultEar.TextMatrix(i, 0) & "-" & ResultEar.TextMatrix(i, 1), lstrCheck)
                    End If
                End If
            Next
        'End If
        
    Else '��ʱΪ����¼��
         '---------
         
         
         '---------
    End If
    
    'lstrcheck�ַ������
    If (Not lstrCheck = "") And (Not ResultChanged = 2) Then
        isOk = MsgBox("������Ŀδ��д�����ȷ��������" & Chr(10) & "δ��д������¼�����ݿ⣡" & Chr(10) & Chr(10) & Trim(lstrCheck), vbOKCancel)
        If isOk = 2 Then
            Set lcolResult = Nothing
            Set lcolItem = Nothing
            ccmdUnfilledAllPass.Enabled = True
            ccmdAutoFill.Enabled = True
            ccmdSave.Enabled = True
            ctlb������.Buttons(2).Enabled = True
            ccmdClearResult.Enabled = True
            Exit Sub
        End If
    End If
        
    If ResultChanged = 2 Then
        isOk = MsgBox("�Ƿ񱣴�������Ա���޸Ľ����", vbOKCancel)
        If isOk = 1 Then
            subSave         '�����������ɹ���ʾ
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
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�������¼��", "frmHEENT_ResultInput", "ccmdSave_Click", 6666, lstrError, False
End Sub

Private Sub ccmdSelInfo_Click()
    On Error GoTo errHandler
    Dim lobjTmp, lobjRec As Object
    Dim i As Integer, j As Integer
    Dim lstrWhere As String
    Dim coptIndex As Integer
    
    'ÿ��������ѯǰ������������ı�ʶȥ��
    cchk���������.Value = 0
    
    For i = 0 To coptClasses.Count - 1
        If coptClasses(i).Value = True Then coptIndex = i
    Next
    lstrWhere = " and �������='" & coptClasses(coptIndex).Caption & "'"
        
    '��װ��ѯ����
    If cchkDateBatch.Value = 1 Then
        lstrWhere = lstrWhere & " and �������>='" & Format(cdtpDateBatch.Value, "yyyy-mm-dd 00:00:00") & "' and �������<='" & Format(cdtpDateBatch.Value, "yyyy-mm-dd 23:59:59") & "'"
    End If
    
    If cchkCompanyBatch.Value = 1 Then
        lstrWhere = lstrWhere & " and ��λ����='" & Trim(ctxtQueyCompanyBatch.Text) & "'"
    End If
    
    '2012-07-14 �ڵ�� ��
    '���Ĳ�ѯ����������8/48Сʱ�ж����ݡ������޸�ʱ���ʼ�ղ������ѯ����С�
    '��ѯ���ݱ�����ݷ����ϴ�仯�����޸ģ������⡣

    '���ÿ������������������Ա�޸�ʱ�����¸��¡���������Ϣ���С��������״̬����'2'��Ϊ'3'�ģ���ѯʱ���ԡ�
    sub���¿��޸Ľ����Ա�޸�״̬
    
    Set lobjTmp = CreateObject("ְҵ�������¼��.clscommon")
    Set mobjQueryResult = lobjTmp.func��ȡ���޸Ľ���_������_�����Ա��Ϣ(lstrWhere, priDeptName)
    
    sub��ѯ�б���ʾ coptIndex
    '2012-07-14 �ڵ�� ��
    
    Set lobjTmp = Nothing
    Set lobjRec = Nothing
    lstrWhere = ""
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�������¼��", "frmHEENT_ResultInput", "ccmdQuery_Click", 6666, lstrError, False
End Sub

Private Sub ccmdUnfilledAllPass_Click()
    On Error GoTo errHandler
    
    Dim hasFilled As Boolean
    Dim i, j As Integer
    
    If SSTResultIn.Tab = 1 Then GoTo ENTFill
    If coptClasses(2).Value = False Then
        'ְҵ�����ͷ�����Ա��δ��д���Զ�����������������
        For i = 2 To 15
            If ResultEye.RowHidden(i) = False Then  '����������
                hasFilled = True
                For j = 2 To 5
                    If i <> 5 And i <> 6 And ResultEye.TextMatrix(i, j) = "" Then hasFilled = False: Exit For
                Next j
                If hasFilled = False Then For j = 2 To 5: ResultEye.TextMatrix(i, j) = "����": Next
                
                hasFilled = True
                For j = 7 To 10
                    If i <> 5 And i <> 6 And ResultEye.TextMatrix(i, j) = "" Then hasFilled = False: Exit For
                Next j
                If hasFilled = False Then For j = 7 To 10: ResultEye.TextMatrix(i, j) = "����": Next j
                
                If ResultEye.TextMatrix(5, 3) = "" Then ResultEye.TextMatrix(5, 3) = "����"
                If ResultEye.TextMatrix(5, 5) = "" Then ResultEye.TextMatrix(5, 5) = "����"
                If ResultEye.TextMatrix(5, 8) = "" Then ResultEye.TextMatrix(5, 8) = "����"
                If ResultEye.TextMatrix(5, 10) = "" Then ResultEye.TextMatrix(5, 10) = "����"
                If ResultEye.TextMatrix(6, 3) = "" Then ResultEye.TextMatrix(6, 3) = "����"
                If ResultEye.TextMatrix(6, 5) = "" Then ResultEye.TextMatrix(6, 5) = "����"
                If ResultEye.TextMatrix(6, 8) = "" Then ResultEye.TextMatrix(6, 8) = "����"
                If ResultEye.TextMatrix(6, 10) = "" Then ResultEye.TextMatrix(6, 10) = "����"
                
                If i = 15 And ResultEye.TextMatrix(i, 2) = "����" And ResultEye.TextMatrix(i, 6) = "" Then
                    ResultEye.TextMatrix(15, 6) = "����"
                End If
            End If
        Next i
        sub���������ͷ��ʽ_�ۿ�
    Else
        '��˲��ӵ�δ��д���Զ�����������������
        For i = 2 To 5
            If ResultEyeArmy.TextMatrix(i, 2) = "" Then
                For j = 2 To 4: ResultEyeArmy.TextMatrix(i, j) = "����": Next
            End If
            If ResultEyeArmy.TextMatrix(i, 6) = "" Then
                For j = 6 To 8: ResultEyeArmy.TextMatrix(i, j) = "����": Next
            End If
        Next i
        
        For i = 7 To 11
            For j = 2 To 8
                If j <> 5 And ResultEyeArmy.TextMatrix(i, j) = "" Then ResultEyeArmy.TextMatrix(i, j) = "����"
            Next j
        Next i
    
        If ResultEyeArmy.TextMatrix(12, 2) = "" Then
            For j = 2 To 8: ResultEyeArmy.TextMatrix(12, j) = "����": Next
        End If
        
        sub���������ͷ��ʽ_�ۿ�_��˲���
    End If
    Exit Sub
    
ENTFill:
    For i = 1 To 18
        If ResultEar.RowHidden(i) = False And ResultEar.TextMatrix(i, ResultEar.cols - 1) = "" Then ResultEar.TextMatrix(i, ResultEar.cols - 1) = "����"
    Next
    sub���������ͷ��ʽ_���Ǻ��
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�������¼��", "frmHEENT_ResultInput", "ccmdUnfilledAllPass_Click", 6666, lstrError, False
End Sub

Private Sub ccmd��ѯ��λ_Click()
    Dim lobjRec As Object                       '��λ��λ���صĽ����¼��
    
    On Error GoTo errHandler
    Set lobjRec = pobjҵ�����.func��λ��λ     '������λ��λ���档
    
    '��ȡ��λ�ĵ�λ����ʾ�ڡ���λ���ơ�¼����С�(��ʱֻ��ʾ����λ���ơ�)
    If Not lobjRec Is Nothing Then
        If lobjRec.recordcount > 0 Then
            ctxtQueyCompanyBatch.Text = IIf(IsNull(lobjRec("��λ����")), "", lobjRec("��λ����"))
        End If
    End If
    'flag����.Value = 1
    Exit Sub
errHandler:
    'If Err.Number = 0 Then Exit Sub
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����沿��", "FrmImportExcel", "ccmd��λ��λ_Click", 6666, lstrError, False
End Sub

Private Sub cgrdInfo_DblClick()
    If coptClasses(2).Value = False Then
        ResultEye.Clear
        sub���������ͷ��ʽ_�ۿ�
    Else
        ResultEyeArmy.Clear
        sub���������ͷ��ʽ_�ۿ�_��˲���
    End If
    ResultEar.Clear
    sub���������ͷ��ʽ_���Ǻ��
    indX = cgrdinfo.MouseRow
    indY = cgrdinfo.MouseCol
    If indX < 0 Or indY < 0 Then
        Exit Sub
    ElseIf indX >= 0 And indX < cgrdinfo.rows And indY >= 0 And indY < cgrdinfo.cols Then
        ctxtBarCode.Text = cgrdinfo.TextMatrix(indX, 0)
        ctxtBarCode_LostFocus
        
        '2012-07-03 �ڵ�� ��
        'ÿ�ζ��������Ϣʱ���ж��Ƿ񳬹��޸�ʱ�䡣
        '�Դ˿��Ʊ��水ť�Ƿ���á�
        If pobjҵ�����.sub�Ƿ����޸�ʱ�䷶Χ��(ctxtBarCode.Text, priDeptName, 8) = False Then
            ctlb������.Buttons(2).Enabled = False
        End If
        '2012-07-03 �ڵ�� ��
    End If
End Sub

'''Private Sub cgrdInfoBatch_Click()
'''    cgrdInfoBatch.SelectionMode = flexSelectionByRow
'''End Sub

Private Sub cgrdInfoBatch_DblClick()
    If cchk���������.Value = 0 Then
        If coptClasses(2).Value = False Then
            ResultEye.Clear
            sub���������ͷ��ʽ_�ۿ�
        Else
            ResultEyeArmy.Clear
            sub���������ͷ��ʽ_�ۿ�_��˲���
        End If
        ResultEar.Clear
        sub���������ͷ��ʽ_���Ǻ��
    End If
    indX = cgrdInfoBatch.MouseRow
    indY = cgrdInfoBatch.MouseCol
    If indX <= 0 Or indY < 0 Then
        Exit Sub
    ElseIf indX > 0 And indX < cgrdInfoBatch.rows And indY >= 0 And indY < cgrdInfoBatch.cols Then
        ctxt�������.Text = cgrdInfoBatch.TextMatrix(indX, 0)
        ctxt�������_lostfocus
    End If
End Sub

'2012-05-11 ��¶
'�������е�������ģ�� �ɽ���ѡ��
Private Sub Cmd����ģ��_Click()
    frmConclusion.lobj���� = priDeptName
    frmConclusion.lobj���ұ�� = priDeptNo
    frmConclusion.lobjҽ����� = um�û����
    frmConclusion.lobjʱ�� = Now
    frmConclusion.Show
End Sub
'2012-05-11 ��¶

Private Sub coptClasses_Click(Index As Integer)
    If coptClasses(0).Value = True Or coptClasses(1).Value = True Then
        ResultEye.Visible = True
        ResultEyeArmy.Visible = False
        sub���������ͷ��ʽ_�ۿ�
    End If
    If coptClasses(2).Value = True Then
        ResultEye.Visible = False
        ResultEyeArmy.Visible = True
        sub���������ͷ��ʽ_�ۿ�_��˲���
    End If
    sub���������ͷ��ʽ_���Ǻ��
    If coptClasses(0).Value = True Then fraQuery.Caption = "��ѯ�����Ա(ְҵ����)"
    If coptClasses(1).Value = True Then fraQuery.Caption = "��ѯ�����Ա(���乤��)"
    If coptClasses(2).Value = True Then fraQuery.Caption = "��ѯ�����Ա(��˲���)"
    
    Dim coptIndex As Integer
    coptIndex = Index
    sub��ѯ�б���ʾ coptIndex
End Sub
Private Sub ctxtAge_Change()
    If ctxtAge.Text = "" Then Exit Sub
    If IsNumeric(CLng(ctxtAge.Text)) = False Then
        MsgBox ("�������ΪС��150�����֣�")
        Exit Sub
    ElseIf CLng(ctxtAge.Text) >= 150 Then
        MsgBox ("����ҪС��150��")
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
    Dim str���ҽ��� As String
    Dim lcolְҵ������ As Object
    lstrNo = Trim(ctxtBarCode.Text)
    
    '���������Ƿ����
    Dim lobjTmp As Object
    Set lobjTmp = CreateObject("ְҵ�������¼��.clscommon")
    Set mlobjRec = lobjTmp.func��ȡ�����Ա������Ϣ(lstrNo)
    If mlobjRec.recordcount = 0 Then
        Set mlobjRec = Nothing
        
        '��յ�ǰ������Ϣ
        ctxtBarCode.Enabled = True
        ctxtName.Text = ""
        ctxtSex.Text = ""
        ctxtAge.Text = ""
        ctxtCompanyName.Text = ""
        Exit Sub
    End If
    
    '�ж��Ƿ�ÿ����д���Ա�����Ȩ��
    Set lobjTmp = CreateObject("ְҵ�������¼��.clscommon")
    If lobjTmp.func��ȡ�����Ա��������Ϣ(lstrNo, priDeptName) Then
        Set lobjTmp = Nothing
        '�������еĸ�����Ϣ�����е������
        LoadPersonalInfo (lstrNo)
        
        Set lcolְҵ������ = CreateObject("ְҵ������.clsManageMedicalExam")
        str���ҽ��� = lcolְҵ������.func���ؿ��ҽ���(ctxtBarCode.Text, priDeptName)
        ctxtConclun.Text = str���ҽ���
        
        'һ��ȷ����ǰ�����Ա��ţ��Ͳ��ܸ��ġ����ǣ���ս������ݡ�
        ctxtBarCode.Enabled = False
        ctxtName.Enabled = False
        ctxtSex.Enabled = False
        ctxtAge.Enabled = False
        ctxtCompanyName.Enabled = False '��ʵ��λ�ҵ���֮������С���λ��λ����ť�����ǿ��Ըĵġ�
        If ResultChanged <> 3 Then
            ccmdUnfilledAllPass.Enabled = True
            ccmdAutoFill.Enabled = True
            ccmdSave.Enabled = True
            ctlb������.Buttons(2).Enabled = True
            ccmdClearResult.Enabled = True
            '2012-06-27 �ڵ�� ��
            'ÿ�ζ��������Ϣʱ���ж��Ƿ񳬹��޸�ʱ�䡣
            '�Դ˿��Ʊ��水ť�Ƿ���á�
            If pobjҵ�����.sub�Ƿ����޸�ʱ�䷶Χ��(ctxtBarCode.Text, priDeptName, 8) = False Then
                ctlb������.Buttons(2).Enabled = False
                ccmdSave.Enabled = False
            End If
            '2012-06-27 �ڵ�� ��
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
        MsgBox ("û�и������Ӧ�������Ա��Ϣ��")
        If cgrdinfo.rows > 0 Then cgrdinfo.RemoveItem
        subClear   '''2012-07-04 �ڵ�� ��ʱע�ͣ�setfocus���������Ŵ���ʱ������ѭ������ѯ����ʧЧ
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

Private Sub ctxt�������_lostfocus()
    Dim lstrNo As String
    Dim i As Integer
    Dim str���ҽ��� As String
    Dim lcolְҵ������ As Object
    lstrNo = Trim(ctxt�������.Text)
    
    '���������Ƿ����
    Dim lobjTmp As Object
    Set lobjTmp = CreateObject("ְҵ�������¼��.clscommon")
    Set mlobjRec = lobjTmp.func��ȡ�����Ա������Ϣ(lstrNo)
    If mlobjRec.recordcount = 0 Then
        Set mlobjRec = Nothing

        '��յ�ǰ������Ϣ
        ctxt�������.Enabled = True
        ctxt����.Text = ""
        ctxt�Ա�.Text = ""
        ctxt����.Text = ""
        ctxt��λ����.Text = ""
        Exit Sub
    End If
    
    Set lobjTmp = CreateObject("ְҵ�������¼��.clscommon")
    If lobjTmp.func��ȡ�����Ա��������Ϣ(lstrNo, priDeptName) Then
        Set lobjTmp = Nothing
       
        LoadPersonalInfoBatch (lstrNo)
        If cchk���������.Value = 0 Then
            Set lcolְҵ������ = CreateObject("ְҵ������.clsManageMedicalExam")
            str���ҽ��� = lcolְҵ������.func���ؿ��ҽ���(ctxt�������.Text, priDeptName)
            ctxtConclun.Text = str���ҽ���
        End If
        'һ��ȷ����ǰ�����Ա��ţ��Ͳ��ܸ��ġ����ǣ���ս������ݡ�
        ctxt�������.Enabled = False
        ctxt����.Enabled = False
        ctxt�Ա�.Enabled = False
        ctxt����.Enabled = False
        ctxt��λ����.Enabled = False '��ʵ��λ�ҵ���֮������С���λ��λ����ť�����ǿ��Ըĵġ�
        If ResultChanged <> 3 Then
            ccmdUnfilledAllPass.Enabled = True
            ccmdAutoFill.Enabled = True
            ccmdSave.Enabled = True '''''''''''''''''''''''''''�����ˣ�Ϊɶ֮ǰд����false
            ctlb������.Buttons(3).Enabled = True
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
        ctlb������.Buttons(2).Enabled = False
        ctlb������.Buttons(3).Enabled = True
        ccmdSave.Enabled = False
    Else
        Set lobjTmp = Nothing
        MsgBox ("�������Աû�иÿ��ҵ������Ŀ��")
        cgrdInfoBatch.RemoveItem
        subClear
    End If
End Sub

Private Sub ResultEye_AfterEdit(ByVal row As Long, ByVal col As Long)
    sub���������ͷ��ʽ_�ۿ�
End Sub

Private Sub ResultEye_DblClick()
    If ResultChanged = 3 Then Exit Sub
    indX = ResultEye.MouseRow
    indY = ResultEye.MouseCol
    ResultEye.Editable = flexEDKbdMouse
    If indX < 0 Or indY < 0 Then
        Exit Sub
    ElseIf indX >= 0 And indX < ResultEye.rows And indY >= 0 And indY < ResultEye.cols Then
        sub¼���޸����� indX, indY
        If ResultChanged = 1 Then ResultChanged = 2
    End If
End Sub


'ֻ�ǲ��Ա��湦�ܡ��ⲿ��д���൱������������
'���а������ mouseCol �� mouseRow �ı����ǲ���ֻ����һ�Σ�֮��ϵͳ�Զ�ɾ������
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
        For i = 2 To 10: ResultEye.TextMatrix(indX, i) = "����": Next
        If ResultChanged = 1 Then ResultChanged = 2
    ElseIf indY >= 2 And indY <= 5 And ResultEye.TextMatrix(indX, indY) = "" Then
        For i = 2 To 5: ResultEye.TextMatrix(indX, i) = "����": Next
        If ResultChanged = 1 Then ResultChanged = 2
    ElseIf indY >= 7 And indY <= 10 And ResultEye.TextMatrix(indX, indY) = "" Then
        For i = 7 To 10: ResultEye.TextMatrix(indX, i) = "����": Next
        If ResultChanged = 1 Then ResultChanged = 2
    End If
    sub���������ͷ��ʽ_�ۿ�
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    ctxtBarCode.SetFocus
    ctxtBarCode.SelStart = Len(ctxtBarCode)
    ctxtBarCode.SelLength = 0
    
End Sub

Private Sub Form_Load()
    Dim lcol��������ť As New Collection '�������ϵİ�ť��ʼ�����ϡ�
    
    On Error GoTo errHandler
    
    '��������Ѿ���ʼ�������ٽ��г�ʼ����
    If mblnInUse Then Exit Sub
    
    '���ô�������ʹ�õı�־��
    mblnInUse = True
    
    '��ʼ������ͨ�ö���ÿ�������Ӧһ������ͨ�ö���ʵ�������ɻ��ã��м��мǣ���
    Set mobjGUI = New cls����ͨ�ö���
    
    '���ù�����������Ҫ�ĸ��ְ�ť��
    With lcol��������ť
        .Add "��ս���(&N)110"
        .Add "����"
        .Add "��������(&S)"
        .Add "|"
        .Add "�˳�"
    End With
    
    'Ϊ��Ҫͨ������ͨ�ö�����Ƶĸ��ֿؼ�����ʼֵ��
    With mobjGUI
        Set .Form = Me
        Set .c������ = ctlb������
    End With
    
    '���ý���ͨ�ö����ṩ�ķ������Խ���ؼ����г�ʼ����
    mobjGUI.subInitialize lcol��������ť, ""

    '������ʱ��ҽʦ����Ϊ��ǰ�û�����
    mstrDoctorName = um�û���
    LabelDoctor.Caption = LabelDoctor.Caption & " " & mstrDoctorName
    
    '����Ȩ�����á������ڸý����ϸ�����ť�������ؼ���ʹ�á�
    '���¹�����ʱ�У��鿴���޸��뱣�棬�����֡�������ɾ��������Ҳû��ɾ����ť��
    Dim lobjTmp As Object
    Set lobjTmp = CreateObject("ְҵ������.clsPermissionConfigure")
    If lobjTmp.func���Ҳ���Ȩ��(um�û����, "ְҵ�����_��ٿƽ��¼��_�޸�") = False Then
        ResultChanged = 3           '����vsFlexGrid.Enabled
        ccmdUnfilledAllPass.Visible = False
        ccmdAutoFill.Visible = False
        ccmdSave.Visible = False
        ccmdClearResult.Visible = False
        ctlb������.Buttons(2).Visible = False
    End If
    
    '2012-05-22 ���� ������
    '����Ȩ������
    If lobjTmp.func���Ҳ���Ȩ��(um�û����, "ְҵ�����_��ٿƽ��¼��_�����޸�") = False Then
        ctlb������.Buttons(3).Visible = False
    End If
    '2012-05-22 ������
    Set lobjTmp = Nothing
    
    '�������ݽ��٣�������ִ���ˡ����ң�Ĭ�ϼ���ְҵ��������
    sub���������ͷ��ʽ_�ۿ�
    sub���������ͷ��ʽ_���Ǻ��
    
    '���水ť�趨
    cdtpConclusionDate = Now       '¼�����������Ϊ��ǰ����
    cdtpDate.Value = Now           '��ѯ���������Ϊ��ǰ����
    ctxtBarCode.Enabled = True
    ccmdUnfilledAllPass.Enabled = False
    ccmdAutoFill.Enabled = False
    ccmdSave.Enabled = False
    ctlb������.Buttons(2).Enabled = False
    ctlb������.Buttons(3).Enabled = False
    ccmdClearResult.Enabled = False
    SSTResultIn.Enabled = False
    If ResultChanged <> 3 Then ResultChanged = -1
    coptClasses(0).Value = True
    DTP¼������.Value = Now
    cdtpDateBatch.Value = Now
    '������ѯ���ܲ�ѯ�����ȡ
    
    
    'lobj������������
        
    '��ͼ�����趨
    DrawState = -1       'form_loadʱ����ͼ״̬Ϊ��δ���޸Ĳ����桱
    If coptClasses(0).Value = True And SSTResultIn.Tab = 0 Then Picture3.Picture = Nothing '�����乤������˲�����ͼƬ����
    subԭͼ����
    
    'ѯ��frame��ʼ�趨
    fraQuery.Caption = "��ѯ�����Ա(ְҵ����)"

    '2012-07-03 �ڵ�� ��
    '����ϵͳ��Ź̶����֡�ʡ������Ҫ���иı�ϵͳ��Ź���
    '��ȡϵͳ��Ź̶����֡�
    sub��ȡϵͳ��Ź̶�����
    '2012-07-03 �ڵ�� ��
    
    '2012-07-14 �ڵ�� ��
    '��ʼ����ѯ���棬������ѯ�б��ʽ����ʼ�����һ�����Ϣ��
    priDeptName = "��ٿ�"
    priDeptNo = "01"
    priDeptResultName = "��ٿ�"
    ccmdQuery_Click
    SSTPersonalInfo.Tab = 1: ccmdSelInfo_Click: SSTPersonalInfo.Tab = 0
    '2012-07-14 �ڵ�� ��
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�������¼��", "frmHEENT_ResultInput", "Form_Load", 6666, lstrError, False
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
        ctlb������.Buttons(3).Enabled = False
    ElseIf SSTPersonalInfo.Tab = 1 Then
        ccmdSave.Enabled = False
        ctlb������.Buttons(2).Enabled = False
    End If
End Sub

Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    On Error GoTo errHandler
    Dim i As Integer
    
    Cancel = True
    Select Case Operate
    Case "��ս���"
        subClear
    Case "����"
        '2012-07-03 �ڵ�� ��
        '�ж��Ƿ����޸�ʱ�䷶Χ��
        If pobjҵ�����.sub�Ƿ����޸�ʱ�䷶Χ��(Trim(ctxtBarCode.Text), priDeptName, 8) = False Then
            MsgBox ("���ϴ��޸��Ѿ�����8Сʱ���������Ա��ϵ����޸�Ȩ�޺��ټ�����")
            Exit Sub
        End If
        '2012-07-03 �ڵ�� ��
        
        '2012-07-15 �ڵ�� ��
        'û��¼��������ʱ����ʾ�Ҳ����档
        If Len(Trim(ctxtConclun.Text)) > 0 Then
            MsgBox "�㻹û��Ϊ�����½���"
            GoTo errHandler
        End If
        '2012-07-15 �ڵ�� ��
        
        '2012-07-03 �ڵ�� ��
        '����һ���ֶ�"�޸���ʼʱ��"���޸ġ�ͬʱ�޸ĸÿ��ҵ������¼��״̬��
        pobjҵ�����.sub�޸���ʼʱ�� Trim(ctxtBarCode.Text), priDeptName
        pobjҵ�����.sub�޸Ľ��¼��״̬ Trim(ctxtBarCode.Text), priDeptNo, "2"  '   01��ʾ��ٿ�
        pobjҵ�����.sub���¼���޸����״̬ Trim(ctxtBarCode.Text), "4"
        '2012-07-03 �ڵ�� ��
        
        ccmdSave_Click
        
        '2012-07-15 �ڵ�� ��
        '������֮�����½��в�ѯ��
        ccmdQuery_Click
        i = SSTPersonalInfo.Tab
        SSTPersonalInfo.Tab = 1: ccmdSelInfo_Click: SSTPersonalInfo.Tab = i
        '2012-07-15 �ڵ�� ��
    Case "��������"
        '2012-07-15 �ڵ�� ��
        'û��¼��������ʱ����ʾ�Ҳ����档
        If Len(Trim(ctxtConclun.Text)) > 0 Then
            MsgBox "�㻹û��Ϊ�����½���"
            GoTo errHandler
        End If
        '2012-07-15 �ڵ�� ��
        
        sub��������
        
        '2012-07-15 �ڵ�� ��
        '������֮�����½��в�ѯ��
        ccmdQuery_Click
        i = SSTPersonalInfo.Tab
        SSTPersonalInfo.Tab = 1: ccmdSelInfo_Click: SSTPersonalInfo.Tab = i
        '2012-07-15 �ڵ�� ��
    Case "�˳�"
        Unload frmHEENT_ResultInput
        Set frmHEENT_ResultInput = Nothing
    End Select
    
    Exit Sub
errHandler:
    sfsub������ "ְҵ�������¼��", "frmHEENT_ResultInput", "mobjGUI_BeforeOperate", Err.Number, Err.Description, False
    Cancel = True
    Exit Sub
    Resume
End Sub

Sub sub���������ͷ��ʽ_�ۿ�()
    Dim i As Integer
    
    '��������С��λ��(�ֶ�����)
    ResultEye.Height = 3500
    ResultEye.Left = 120
    ResultEye.Top = 240
    ResultEye.Width = 8500
    
    '�����п������λ��
    'ResultEye.AutoSize 2, ResultEye.Cols - 1, 1, 650 '����޸������ݣ��������ڲ���ԭ���ؿ��ݲ�����䡣
    ResultEye.ColWidth(0) = 960
    ResultEye.ColWidth(1) = 640
    For i = 0 To 10: ResultEye.ColAlignment(i) = flexAlignCenterCenter: Next
    For i = 2 To 10: ResultEye.ColWidth(i) = 800: Next
    
    '����ְҵ�����ͷ��乤����ĳЩ����Ҫ����
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
    
    '0~15�У�0~1�з�Χ�ڣ�ֵ��ͬ��Ҫ�ϲ���Ԫ��
    ResultEye.MergeCompare = flexMCIncludeNulls
    ResultEye.MergeCells = flexMergeFree
    For i = 0 To 15: ResultEye.MergeRow(i) = True: Next
    For i = 0 To 1: ResultEye.MergeCol(i) = True: Next
    
    '0��1�У�ǰ���и�ʽ�趨
    ResultEye.TextMatrix(0, 0) = "��Ŀ": ResultEye.TextMatrix(0, 1) = "��Ŀ"
    ResultEye.TextMatrix(1, 0) = "�۱�": ResultEye.TextMatrix(1, 1) = "�۱�"
    ResultEye.TextMatrix(2, 0) = "ɫ��": ResultEye.TextMatrix(2, 1) = "ɫ��"
    ResultEye.TextMatrix(3, 0) = "����Ӧ": ResultEye.TextMatrix(3, 1) = "����Ӧ"
    ResultEye.TextMatrix(4, 0) = "��Ұ": ResultEye.TextMatrix(4, 1) = "��Ұ"
    ResultEye.TextMatrix(5, 0) = "����": ResultEye.TextMatrix(6, 0) = "����"
    ResultEye.TextMatrix(5, 1) = "����": ResultEye.TextMatrix(6, 1) = "����"
    ResultEye.TextMatrix(7, 0) = "��ǰ��": ResultEye.TextMatrix(7, 1) = "��ǰ��"
    ResultEye.TextMatrix(8, 0) = "������϶��" & Chr(10) & "�������": ResultEye.TextMatrix(9, 0) = "������϶��" & Chr(10) & "�������": ResultEye.TextMatrix(10, 0) = "������϶��" & Chr(10) & "�������": ResultEye.TextMatrix(11, 0) = "������϶��" & Chr(10) & "�������": ResultEye.TextMatrix(12, 0) = "������϶��" & Chr(10) & "�������"
    ResultEye.TextMatrix(8, 1) = "��Ĥ": ResultEye.TextMatrix(9, 1) = "��Ĥ": ResultEye.TextMatrix(10, 1) = "ǰ��": ResultEye.TextMatrix(11, 1) = "��Ĥ": ResultEye.TextMatrix(12, 1) = "��״��"
    ResultEye.TextMatrix(13, 0) = "������": ResultEye.TextMatrix(13, 1) = "������"
    ResultEye.TextMatrix(14, 0) = "�۵�": ResultEye.TextMatrix(14, 1) = "�۵�"
    ResultEye.TextMatrix(15, 0) = "�ۿ�����": ResultEye.TextMatrix(15, 1) = "�ۿ�����"
    
    '2~10�и�ʽ�趨���м���Ϊ��6�У�ֵȫΪ***��Ϊ�����е�Ԫ�񲻺ϲ���Ĭ�����ء�
    ResultEye.ColHidden(6) = True
    For i = 1 To 14: ResultEye.TextMatrix(i, 6) = "***": Next
    For i = 2 To 10: ResultEye.TextMatrix(0, i) = "�����": Next
    For i = 2 To 5: ResultEye.TextMatrix(1, i) = "��": Next
    For i = 7 To 10: ResultEye.TextMatrix(1, i) = "��": Next
    ResultEye.TextMatrix(5, 2) = "Զ����": ResultEye.TextMatrix(5, 7) = "Զ����"
    ResultEye.TextMatrix(6, 2) = "Զ����": ResultEye.TextMatrix(6, 7) = "Զ����"
    ResultEye.TextMatrix(5, 4) = "������": ResultEye.TextMatrix(5, 9) = "������"
    ResultEye.TextMatrix(6, 4) = "������": ResultEye.TextMatrix(6, 9) = "������"
    
    'ĳЩ��Ԫ�񲻿ɱ༭
    '��֪������ν������ʱ��ÿ�α༭��ˢ�±�ͷ��ʽ
    
    '��ͷ��ɫ���Բ�̫һ����Ҳ���Ը���һ����ɫ��
End Sub

Sub sub¼���޸�����(ByVal paraRow As Integer, ByVal paraCol As Integer)
    Dim i As Integer
    If SSTResultIn.Tab = 0 Then
        '�ۿƽ����д��
        If coptClasses(2).Value = False Then
            ResultEye.Select paraRow, paraCol
            ResultEye.EditCell
        Else
            ResultEyeArmy.Select paraRow, paraCol
            ResultEyeArmy.EditCell
        End If
    Else
        '���Ǻ�ƽ����д��
        ResultEar.Select paraRow, paraCol
        ResultEar.EditCell
    End If
    
End Sub

Sub LoadPersonalInfoBatch(ByVal paraSysNo As String)
    On Error GoTo errHandler
    
    Dim lobjTmp, lobjRec As Object
    Set lobjTmp = CreateObject("ְҵ�������¼��.clscommon")
    Set mlobjRec = lobjTmp.func��ȡ�����Ա������Ϣ(paraSysNo)
    If mlobjRec.recordcount > 0 Then
        ctxt���� = mlobjRec("����")
        ctxt�Ա� = mlobjRec("�Ա�")
        ctxt���� = mlobjRec("����")
        ctxt��λ���� = mlobjRec("��λ����")
        
        '�����������
        If mlobjRec("�������") = "ְҵ����" Then coptClasses(0).Value = True
        If mlobjRec("�������") = "���乤��" Then coptClasses(1).Value = True
        If mlobjRec("�������") = "��˲���" Then coptClasses(2).Value = True
        
        
         '�������еĸ�����Ϣ�����е������
        If cchk���������.Value = 0 Then
            '��ʾ��Ƭ
            Set lobjRec = CreateObject("ְҵ������.clspersonexamed")
            lobjRec.ϵͳ��� = ctxt�������.Text
            Picture4.Enabled = True
            Picture4.Visible = True
            Picture4.Picture = lobjRec.��Ƭ
        
        
            mstr���ͼƬ��Ŀ��� = lobjTmp.func��ȡ�����Ŀ���("��״�廷�漰����ͼ")
            DrawState = -1
            
            Set lobjRec = lobjTmp.func�Ƿ��Ѿ�����(ctxt�������.Text, priDeptName)
            If lobjRec.recordcount = 0 Then
                If ResultChanged <> 3 Then ResultChanged = 0
            ElseIf lobjRec.recordcount > 0 Then
                If ResultChanged <> 3 Then ResultChanged = 1
                If coptClasses(2).Value = False Then
                    sub��д���е������_�ۿ� lobjRec
                Else
                    sub��д���е������_�ۿ�_��˲��� lobjRec
                End If
                sub��д���е������_���Ǻ�� lobjRec
            End If
            
            '��ʾ�۾������ͼ
            If coptClasses(0).Value = False Then Picture3.Picture = lobjTmp.func��ȡ���ͼƬ(ctxt�������.Text, mstr���ͼƬ��Ŀ���, "��״�廷�漰����ͼ.bmp")  '01069���۾������ͼ����Ŀ��š�
        End If
    Else
        MsgBox ("û�и������Ӧ�������Ա��Ϣ��")
        Exit Sub
    End If
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�������¼��", "frmHEENT_ResultInput", "LoadPersonalInfo", 6666, lstrError, False
End Sub

Sub LoadPersonalInfo(ByVal paraSysNo As String)
    On Error GoTo errHandler
    
    Dim lobjTmp, lobjRec As Object
    Set lobjTmp = CreateObject("ְҵ�������¼��.clscommon")
    Set mlobjRec = lobjTmp.func��ȡ�����Ա������Ϣ(paraSysNo)
    If mlobjRec.recordcount > 0 Then
        ctxtName = mlobjRec("����")
        ctxtSex = mlobjRec("�Ա�")
        ctxtAge = mlobjRec("����")
        ctxtCompanyName = mlobjRec("��λ����")
            
        '��ʾ��Ƭ
        Set lobjRec = CreateObject("ְҵ������.clspersonexamed")
        lobjRec.ϵͳ��� = ctxtBarCode.Text
        Picture2.Enabled = True
        Picture2.Visible = True
        Picture2.Picture = lobjRec.��Ƭ
            
        '�����������
        If mlobjRec("�������") = "ְҵ����" Then coptClasses(0).Value = True
        If mlobjRec("�������") = "���乤��" Then coptClasses(1).Value = True
        If mlobjRec("�������") = "��˲���" Then coptClasses(2).Value = True
            
        mstr���ͼƬ��Ŀ��� = lobjTmp.func��ȡ�����Ŀ���("��״�廷�漰����ͼ")
        DrawState = -1
            
        Set lobjRec = lobjTmp.func�Ƿ��Ѿ�����(ctxtBarCode.Text, priDeptName)
        If lobjRec.recordcount = 0 Then
            If ResultChanged <> 3 Then ResultChanged = 0
        ElseIf lobjRec.recordcount > 0 Then
            If ResultChanged <> 3 Then ResultChanged = 1
            If coptClasses(2).Value = False Then
                sub��д���е������_�ۿ� lobjRec
            Else
                sub��д���е������_�ۿ�_��˲��� lobjRec
            End If
            sub��д���е������_���Ǻ�� lobjRec
        End If
    
        '��ʾ�۾������ͼ
        If coptClasses(0).Value = False Then Picture3.Picture = lobjTmp.func��ȡ���ͼƬ(ctxtBarCode.Text, mstr���ͼƬ��Ŀ���, "��״�廷�漰����ͼ.bmp")  '01069���۾������ͼ����Ŀ��š�
    Else
        MsgBox ("û�и������Ӧ�������Ա��Ϣ��")
        Exit Sub
    End If
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�������¼��", "frmHEENT_ResultInput", "LoadPersonalInfo", 6666, lstrError, False
End Sub

'��������������Ա��������Ϣ(��ѯ�������ݲ����)���ָ���������Ϊform_loadʱ����
Sub subClear()
    TotalPeople.Caption = 0
    TotalPeopleBatch.Caption = 0
    
    '��յ�ǰ������Ϣ
    cdtpConclusionDate = Now
    ctxtBarCode.Text = ""
    ctxtBarCode.Enabled = True
    ctxtName.Text = ""
    ctxtSex.Text = ""
    ctxtAge.Text = ""
    ctxtCompanyName.Text = ""
    cgrdinfo.rows = 1
    
    '������Ϣ���
    ctxt�������.Text = ""
    ctxt�������.Enabled = True
    ctxt����.Text = ""
    ctxt�Ա�.Text = ""
    ctxt����.Text = ""
    ctxt��λ����.Text = ""
    cgrdInfoBatch.rows = 1
    
    '������Ϣ��־���
    cchk���������.Value = 0
    
    Picture1.Picture = Nothing
    Picture2.Picture = Nothing
    Picture3.Picture = Nothing
    Picture4.Picture = Nothing
    ctxtConclun.Text = ""

'    '��ղ�ѯ�������һ��Ҫ�е�,Ҳûдȫ��
'    cchkDate.Value = 0
'    cdtpDate.Value = Now
'    cgrdInfo.Clear

    '��ս����
    ResultEye.Clear
    sub���������ͷ��ʽ_�ۿ�
    ResultEyeArmy.Clear
    sub���������ͷ��ʽ_�ۿ�_��˲���
    ResultEar.Clear
    sub���������ͷ��ʽ_���Ǻ��

    '�ָ�Ϊform_loadʱ��״̬��
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
        ctxt�������.Enabled = True
        ctxt����.Enabled = True
        ctxt�Ա�.Enabled = True
        ctxt����.Enabled = True
        ctxt��λ����.Enabled = True
    End If
    
    ccmdSave.Enabled = False
    ctlb������.Buttons(2).Enabled = False
    ctlb������.Buttons(3).Enabled = False
    ccmdClearResult.Enabled = False
    SSTResultIn.Enabled = False
    If ResultChanged <> 3 Then ResultChanged = -1
    coptClasses(0).Value = True
    coptClasses(0).Enabled = True: coptClasses(1).Enabled = True: coptClasses(2).Enabled = True
    '��ͼ�����趨
    If coptClasses(0).Value = True Then Picture3.Picture = Nothing '�����乤������˲�����ͼƬ����
    
    '2012-06-21 �ڵ�� ��
    '��ʼ����ǰ¼��״̬(��ǰ�ж�����Ȩ���޸ģ����ޣ�ֱ�Ӹ�ֵΪ3)
    ResultChanged = IIf(ResultChanged <> 3, -1, 3)
    cchkˢ����_Click
    '2012-06-21 �ڵ�� ��
End Sub

Private Function sub��ӵ�����(ByVal paraResult As String, ByVal paraItem As String, ByVal paraCheck As String) As String
    If paraResult = "" Then
        paraCheck = paraCheck & IIf(paraCheck = "", "", Chr(10)) & paraItem
    Else
        lcolItem.Add paraItem
        lcolResult.Add paraResult
    End If
    sub��ӵ����� = paraCheck
End Function

Sub subSave()
    On Error GoTo errHandler
    
    Dim isOk As Boolean
    Dim lstrTmp As String
    Dim lobjTmp As Object
    Set lobjTmp = CreateObject("ְҵ�������¼��.clscommon")
    isOk = lobjTmp.func���浥�������(ctxtBarCode.Text, mstrDoctorName, cdtpConclusionDate.Value, lcolItem, lcolResult, "ְҵ�����_�����Ϣ_��ٿ�")
    subClear
    If ResultChanged <> 3 Then ResultChanged = 1
    If isOk = True Then MsgBox ("����ɹ���")
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�������¼��", "frmHEENT_ResultInput", "subSave", 6666, lstrError, False
End Sub

'û�а�����˲��ӵĽ����д����Ϊ��������ܴ�
Sub sub��д���е������_�ۿ�(ByVal paraRec As Object)
    Dim strArray
    Dim i, j, k, L, intTmp As Integer
    Dim hasSaved As Boolean
      
    For i = 1 To paraRec.recordcount
CONTINUE:
        If IsNull(paraRec("�����")) = True And paraRec.EOF = False Then
            paraRec.MoveNext
            GoTo CONTINUE
        ElseIf paraRec.EOF = True Then
            Exit Sub
        End If
        
        hasSaved = False
        strArray = Split(paraRec("��Ŀ����"), "-", -1, vbBinaryCompare)
        L = UBound(strArray)
        If L = 0 Then   'ֻ����ְͨҵ��������ġ�������
            If paraRec("��Ŀ����") = "�ۿ�����" Then
                For j = 2 To 10: ResultEye.TextMatrix(15, j) = paraRec("�����"): Next
                hasSaved = True
            End If
        End If
        
        If L = 2 And hasSaved = False Then   'ֻ����ְͨҵ��������ġ����ۡ�����������
            If strArray(0) = "����" Then
                If (strArray(1) = "Զ����" And strArray(2) = "��") Then ResultEye.TextMatrix(5, 3) = paraRec("�����"): hasSaved = True
                If (strArray(1) = "Զ����" And strArray(2) = "��") Then ResultEye.TextMatrix(5, 8) = paraRec("�����"): hasSaved = True
                If (strArray(1) = "������" And strArray(2) = "��") Then ResultEye.TextMatrix(5, 5) = paraRec("�����"): hasSaved = True
                If (strArray(1) = "������" And strArray(2) = "��") Then ResultEye.TextMatrix(5, 10) = paraRec("�����"): hasSaved = True
            End If
            If strArray(0) = "����" Then
                If (strArray(1) = "Զ����" And strArray(2) = "��") Then ResultEye.TextMatrix(6, 3) = paraRec("�����"): hasSaved = True
                If (strArray(1) = "Զ����" And strArray(2) = "��") Then ResultEye.TextMatrix(6, 8) = paraRec("�����"): hasSaved = True
                If (strArray(1) = "������" And strArray(2) = "��") Then ResultEye.TextMatrix(6, 5) = paraRec("�����"): hasSaved = True
                If (strArray(1) = "������" And strArray(2) = "��") Then ResultEye.TextMatrix(6, 10) = paraRec("�����"): hasSaved = True
            End If
        End If
            
        If L = 1 And hasSaved = False Then
            For j = 2 To 14
                If ResultEye.TextMatrix(j, 1) = strArray(0) Then
                    If strArray(1) = "��" Then
                        For k = 2 To 5: ResultEye.TextMatrix(j, k) = paraRec("�����"): Next k
                    Else
                        For k = 7 To 10: ResultEye.TextMatrix(j, k) = paraRec("�����"): Next k
                    End If
                    hasSaved = True
                End If
            Next j
        End If
        paraRec.MoveNext
    Next i
End Sub



'-------����Ϊ��˲������õĺ�������Щ������������������--------

Private Sub ResultEyeArmy_AfterEdit(ByVal row As Long, ByVal col As Long)
    sub���������ͷ��ʽ_�ۿ�_��˲���
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
        For i = 2 To 8: ResultEyeArmy.TextMatrix(indX, i) = "����": Next
        If ResultChanged = 1 Then ResultChanged = 2
    ElseIf indX >= 2 And indX <= 5 And indY >= 2 And indY <= 4 And ResultEyeArmy.TextMatrix(indX, indY) = "" Then
        For i = 2 To 4: ResultEyeArmy.TextMatrix(indX, i) = "����": Next
        If ResultChanged = 1 Then ResultChanged = 2
    ElseIf indX >= 2 And indX <= 5 And indY >= 6 And indY <= 8 And ResultEyeArmy.TextMatrix(indX, indY) = "" Then
        For i = 6 To 8: ResultEyeArmy.TextMatrix(indX, i) = "����": Next
        If ResultChanged = 1 Then ResultChanged = 2
    ElseIf indX >= 7 And indX <= 11 Then
        ResultEyeArmy.TextMatrix(indX, indY) = "����"
        If ResultChanged = 1 Then ResultChanged = 2
    End If
    sub���������ͷ��ʽ_�ۿ�_��˲���
End Sub

Private Sub ResultEyeArmy_DblClick()
    If ResultChanged = 3 Then Exit Sub
    indX = ResultEyeArmy.MouseRow
    indY = ResultEyeArmy.MouseCol
    ResultEyeArmy.Editable = flexEDKbdMouse
    If indX < 0 Or indY < 0 Then
        Exit Sub
    ElseIf indX >= 0 And indX < ResultEyeArmy.rows And indY >= 0 And indY < ResultEyeArmy.cols Then
        sub¼���޸����� indX, indY
        If ResultChanged = 1 Then ResultChanged = 2
    End If
End Sub

Sub sub���������ͷ��ʽ_�ۿ�_��˲���()
    Dim i As Integer
    
    '�������λ�úʹ�С����ְҵ���������乤��һ��
    ResultEyeArmy.Height = ResultEye.Height
    ResultEyeArmy.Left = ResultEye.Left
    ResultEyeArmy.Top = ResultEye.Top
    ResultEyeArmy.Width = ResultEye.Width
    
    '�����п������λ��
    'ResultEyeArmy.AutoSize 2, ResultEyeArmy.Cols - 1, 1, 650 '����޸������ݣ��������ڲ���ԭ���ؿ��ݲ�����䡣
    ResultEyeArmy.ColWidth(0) = 960
    ResultEyeArmy.ColWidth(1) = 640
    For i = 0 To 8: ResultEyeArmy.ColAlignment(i) = flexAlignCenterCenter: Next
    
    '0~12�У�0~1�з�Χ�ڣ�ֵ��ͬ��Ҫ�ϲ���Ԫ��
    ResultEyeArmy.MergeCompare = flexMCIncludeNulls
    ResultEyeArmy.MergeCells = flexMergeFree
    ResultEyeArmy.MergeRow(12) = True
    For i = 0 To 5: ResultEyeArmy.MergeRow(i) = True: Next
    For i = 0 To 1: ResultEyeArmy.MergeCol(i) = True: Next
    
    '0��1�У�ǰ���и�ʽ�趨
    ResultEyeArmy.TextMatrix(0, 0) = "��Ŀ": ResultEyeArmy.TextMatrix(0, 1) = "��Ŀ"
    ResultEyeArmy.TextMatrix(1, 0) = "�۱�": ResultEyeArmy.TextMatrix(1, 1) = "�۱�"
    ResultEyeArmy.TextMatrix(2, 0) = "����": ResultEyeArmy.TextMatrix(3, 0) = "����"
    ResultEyeArmy.TextMatrix(2, 1) = "����": ResultEyeArmy.TextMatrix(3, 1) = "����"
    ResultEyeArmy.TextMatrix(4, 0) = "��Ĥ": ResultEyeArmy.TextMatrix(4, 1) = "��Ĥ"
    ResultEyeArmy.TextMatrix(5, 0) = "��Ĥ": ResultEyeArmy.TextMatrix(5, 1) = "��Ĥ"
    ResultEyeArmy.TextMatrix(6, 0) = "������϶��" & Chr(10) & "�������": ResultEyeArmy.TextMatrix(7, 0) = "������϶��" & Chr(10) & "�������": ResultEyeArmy.TextMatrix(8, 0) = "������϶��" & Chr(10) & "�������": ResultEyeArmy.TextMatrix(9, 0) = "������϶��" & Chr(10) & "�������": ResultEyeArmy.TextMatrix(10, 0) = "������϶��" & Chr(10) & "�������": ResultEyeArmy.TextMatrix(11, 0) = "������϶��" & Chr(10) & "�������"
    ResultEyeArmy.TextMatrix(7, 1) = "�۳�״": ResultEyeArmy.TextMatrix(8, 1) = "��״": ResultEyeArmy.TextMatrix(9, 1) = "Ƭ״": ResultEyeArmy.TextMatrix(10, 1) = "����": ResultEyeArmy.TextMatrix(11, 1) = "����"
    ResultEyeArmy.TextMatrix(12, 0) = "�ۿ����": ResultEyeArmy.TextMatrix(12, 1) = "�ۿ����"
    
    '2~8�и�ʽ�趨���м���Ϊ��5�У�ֵȫΪ***��Ϊ�����е�Ԫ�񲻺ϲ���Ĭ�����ء�
    ResultEyeArmy.ColHidden(5) = True
    For i = 1 To 11: ResultEyeArmy.TextMatrix(i, 5) = "***": Next
    For i = 2 To 8: ResultEyeArmy.TextMatrix(0, i) = "�����": Next
    For i = 2 To 4: ResultEyeArmy.TextMatrix(1, i) = "��": Next
    For i = 6 To 8: ResultEyeArmy.TextMatrix(1, i) = "��": Next
    ResultEyeArmy.TextMatrix(6, 2) = "������": ResultEyeArmy.TextMatrix(6, 3) = "ǰ����": ResultEyeArmy.TextMatrix(6, 4) = "���"
    ResultEyeArmy.TextMatrix(6, 6) = "������": ResultEyeArmy.TextMatrix(6, 7) = "ǰ����": ResultEyeArmy.TextMatrix(6, 8) = "���"
End Sub

Sub sub��д���е������_�ۿ�_��˲���(ByVal paraRec As Object)
    Dim strArray
    Dim i, j, k, Upd, intTmp As Integer
    Dim hasSaved As Boolean
    
    For i = 1 To paraRec.recordcount
CONTINUE:
        If IsNull(paraRec("�����")) = True And paraRec.EOF = False Then
            paraRec.MoveNext
            GoTo CONTINUE
        ElseIf paraRec.EOF = True Then
            Exit Sub
        End If
        
        hasSaved = False
        strArray = Split(paraRec("��Ŀ����"), "-", -1, vbBinaryCompare)
        Upd = UBound(strArray)
        If Upd = 0 Then   'ֻ����ְͨҵ��������ġ�������
            If paraRec("��Ŀ����") = "�ۿ����" Then
                For j = 2 To 8: ResultEyeArmy.TextMatrix(12, j) = paraRec("�����"): Next
                hasSaved = True
            End If
        End If
        
        If Upd = 2 And hasSaved = False Then   'ֻ����ְͨҵ��������ġ����ۡ�����������
            For j = 7 To 11
                For k = 2 To 4
                    If ResultEyeArmy.TextMatrix(j, 1) = strArray(0) And ResultEyeArmy.TextMatrix(6, k) = strArray(1) Then
                        If strArray(2) = "��" Then
                            ResultEyeArmy.TextMatrix(j, k) = paraRec("�����")
                        Else
                            ResultEyeArmy.TextMatrix(j, k + 4) = paraRec("�����")
                        End If
                        hasSaved = True
                    End If
                Next k
            Next j
        End If
            
        If Upd = 1 And hasSaved = False Then
            For j = 2 To 5
                If ResultEyeArmy.TextMatrix(j, 1) = strArray(0) Then
                    If strArray(1) = "��" Then
                        For k = 2 To 4: ResultEyeArmy.TextMatrix(j, k) = paraRec("�����"): Next k
                    Else
                        For k = 6 To 8: ResultEyeArmy.TextMatrix(j, k) = paraRec("�����"): Next k
                    End If
                    hasSaved = True
                End If
            Next j
        End If
        paraRec.MoveNext
    Next i
End Sub


'-----------���Ǻ�Ƶ����ú���
Private Sub ResultEar_AfterEdit(ByVal row As Long, ByVal col As Long)
    sub���������ͷ��ʽ_���Ǻ��
End Sub

Private Sub ResultEar_Click()
    If ResultChanged = 3 Then Exit Sub
    indX = ResultEar.MouseRow
    indY = ResultEar.MouseCol
    ResultEar.Editable = flexEDKbdMouse
    If indX < 0 Or indY < 0 Then Exit Sub
    ResultEar.TextMatrix(indX, ResultEar.cols - 1) = "����"
    sub���������ͷ��ʽ_���Ǻ��
End Sub

Private Sub ResultEar_DblClick()
    If ResultChanged = 3 Then Exit Sub
    indX = ResultEar.MouseRow
    indY = ResultEar.MouseCol
    ResultEar.Editable = flexEDKbdMouse
    If indX < 0 Or indY < 0 Then
        Exit Sub
    ElseIf indX >= 0 And indX < ResultEar.rows And indY = ResultEar.cols - 1 Then
        sub¼���޸����� indX, indY
        If ResultChanged = 1 Then ResultChanged = 2
    End If
End Sub

Sub sub��д���е������_���Ǻ��(ByVal paraRec As Object)
    Dim strArray
    Dim i, j, L As Integer
    
    paraRec.movefirst
    For j = 1 To paraRec.recordcount
CONTINUE:
        If IsNull(paraRec("�����")) = True And paraRec.EOF = False Then
            paraRec.MoveNext
            GoTo CONTINUE
        ElseIf paraRec.EOF = True Then
            Exit Sub
        End If
        
        strArray = Split(paraRec("��Ŀ����"), "-", -1, vbBinaryCompare)
        L = UBound(strArray)
        If L = 0 Then
            For i = 1 To ResultEar.rows - 1
                If ResultEar.TextMatrix(i, 0) = strArray(0) Then ResultEar.TextMatrix(i, ResultEar.cols - 1) = paraRec("�����")
            Next i
        Else
            For i = 2 To ResultEar.rows - 1
                If ResultEar.TextMatrix(i, 0) = strArray(0) And ResultEar.TextMatrix(i, 1) = strArray(1) Then
                    ResultEar.TextMatrix(i, ResultEar.cols - 1) = paraRec("�����")
                End If
            Next i
        End If
        paraRec.MoveNext
    Next j
End Sub

Sub sub���������ͷ��ʽ_���Ǻ��()
    Dim i As Integer
    
    '�������λ�úʹ�С��
    ResultEar.Height = ResultEye.Height     '��֪����ԭ����ʾʱ����ͬʱ����eye��ear�Ľ��������Ըɴ��С��Ϊһ����
    ResultEar.Left = ResultEye.Left
    ResultEar.Top = ResultEye.Top
    ResultEar.Width = ResultEye.Width
    
    '�����п������λ��
    'ResultEar.AutoSize 2, ResultEar.Cols - 1, 0, 0
    ResultEar.ColWidth(0) = 800
    ResultEar.ColWidth(1) = 500
    ResultEar.ColWidth(2) = ResultEar.Width - ResultEar.ColWidth(0) - ResultEar.ColWidth(1)
    ResultEar.AllowUserResizing = flexResizeColumns
    For i = 0 To 2
        ResultEar.ColAlignment(i) = flexAlignCenterCenter
    Next
    
    '������Ҫ�ϲ��ĵ�Ԫ��
    ResultEar.MergeCompare = flexMCIncludeNulls
    ResultEar.MergeCells = flexMergeFree
    ResultEar.MergeCol(0) = True: ResultEar.MergeCol(1) = True
    For i = 0 To ResultEar.rows - 1: ResultEar.MergeRow(i) = True: Next
    
    '�������ñ�ͷ���ݣ�
    'ְҵ������ͷ����
    ResultEar.TextMatrix(4, 0) = "���": ResultEar.TextMatrix(4, 1) = "���"
    ResultEar.TextMatrix(5, 0) = "�ж�": ResultEar.TextMatrix(5, 1) = "�ж�"
    ResultEar.TextMatrix(6, 0) = "����": ResultEar.TextMatrix(7, 0) = "����"    '����˲��ӹ���
    ResultEar.TextMatrix(6, 1) = "��": ResultEar.TextMatrix(7, 1) = "��"        '����˲��ӹ���
    ResultEar.TextMatrix(3, 0) = "��": ResultEar.TextMatrix(3, 1) = "��"
    ResultEar.TextMatrix(14, 0) = "��ǻ": ResultEar.TextMatrix(15, 0) = "��ǻ"
    ResultEar.TextMatrix(14, 1) = "ճĤ": ResultEar.TextMatrix(15, 1) = "����"
    ResultEar.TextMatrix(16, 0) = "�ʺ�": ResultEar.TextMatrix(16, 1) = "�ʺ�"  '����˲��ӹ���
    
    '���乤����ͷ����
    ResultEar.TextMatrix(1, 0) = "����": ResultEar.TextMatrix(1, 1) = "����"
    ResultEar.TextMatrix(2, 0) = "���": ResultEar.TextMatrix(2, 1) = "���"
    
    '��˲��ӱ�ͷ����
    ResultEar.TextMatrix(8, 0) = "�����": ResultEar.TextMatrix(9, 0) = "�����"
    ResultEar.TextMatrix(8, 1) = "��": ResultEar.TextMatrix(9, 1) = "��"
    ResultEar.TextMatrix(10, 0) = "��ͻ": ResultEar.TextMatrix(11, 0) = "��ͻ"
    ResultEar.TextMatrix(10, 1) = "��": ResultEar.TextMatrix(11, 1) = "��"
    ResultEar.TextMatrix(12, 0) = "��": ResultEar.TextMatrix(13, 0) = "��"
    ResultEar.TextMatrix(12, 1) = "ճĤ": ResultEar.TextMatrix(13, 1) = "��Ѫ"
    ResultEar.TextMatrix(17, 0) = "��ǻ": ResultEar.TextMatrix(17, 1) = "��ǻ"
    
    '������ͷ
    ResultEar.TextMatrix(0, 0) = "��Ŀ": ResultEar.TextMatrix(0, 1) = "��Ŀ"
    ResultEar.TextMatrix(0, 2) = "�����"
    ResultEar.TextMatrix(18, 0) = "���Ǻ������": ResultEar.TextMatrix(18, 1) = "���Ǻ������"
    
    '�����Ŀ���������Ǹ������������Ҫ����
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


'-----------------------����Ϊ��ͼ���Ʋ���-------------------------
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
    Dim i, resultTmp As Integer     'paraExp���С��10,��������
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

'���Ӷ��Ը߰�~~~~~~~6040����
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
    Set lobjTmp = CreateObject("ְҵ�������¼��.clscommon")
    lobjTmp.func������ͼƬ Picture3.Image, ctxtBarCode.Text, mstr���ͼƬ��Ŀ���, cdtpConclusionDate.Value    '01069 ���۾������ͼ����Ŀ��š��ڡ������Ŀ���á��ı����м�¼��
    If DrawState = 0 Then
        Call sub��ӵ�����("����", "��״�廷�漰����ͼ", "")
    ElseIf DrawState <> -1 Then
        Call sub��ӵ�����("������", "��״�廷�漰����ͼ", "")
    End If
    MsgBox ("���ͼƬ����ɹ���")
    DrawState = 0
    Exit Sub
End Sub

Private Sub ccmdLoadOriginalPicture_Click()
    Dim lobjTmp, lobjRec As Object
    Dim isOk As Integer
    Set lobjTmp = CreateObject("ְҵ�������¼��.clscommon")
    Set lobjRec = lobjTmp.func���ҽ��ͼƬ(ctxtBarCode.Text, mstr���ͼƬ��Ŀ���)
    If lobjRec.recordcount > 0 Then
        isOk = MsgBox("����ԭͼ��ɾ�����еĽ����ȷ��������", vbOKCancel)
        If isOk = 2 Then Exit Sub
        lobjTmp.funcɾ�����ͼƬ ctxtBarCode.Text, mstr���ͼƬ��Ŀ���
        DrawState = 0
    End If
    Picture3.Picture = lobjTmp.func��ȡ���ͼƬ(ctxtBarCode.Text, mstr���ͼƬ��Ŀ���, "��״�廷�漰����ͼ.bmp")
End Sub

Sub subԭͼ����()
    Dim i, j, rows, cols As Integer
    Picture3.ScaleMode = 3
    
    Set Picture3.Picture = LoadPicture(App.Path & "\��״�廷�漰����ͼ.bmp")
    cols = Picture3.ScaleWidth - 1
    rows = Picture3.ScaleHeight - 1

    pointCnt = 0
    For i = 1 To cols
        For j = 1 To rows
            'trueʱ����Ϊ��ɫ��falseʱ����Ϊ��ɫ
            If Hex(GetPixel(Picture3.hdc, i, j)) = Hex(&H0) Then
                pointCnt = pointCnt + 1
                EyeMapCheck(pointCnt, 1) = i
                EyeMapCheck(pointCnt, 2) = j
            End If
        Next
    Next
    Set Picture3.Picture = Nothing

End Sub

'���ܣ������ύ�����Ա�������
'ʱ�䣺2012-04-26
'���ߣ�����
Private Sub sub��������()
    Dim lstrCheck, lstrItem, lstrResult As String
    Dim i, j, isOk As Integer
    Dim lobjTmp As Object
    
On Error GoTo errHandler
    '¼����������ʱ���ܲ���
    ccmdUnfilledAllPass.Enabled = False
    ccmdAutoFill.Enabled = False
    ccmdSave.Enabled = False
    ctlb������.Buttons(3).Enabled = False
    ccmdClear.Enabled = False
    SSTResultIn.Enabled = False
    
    Set lcolResult = New Collection
    Set lcolItem = New Collection

    If coptClasses(2).Value = False Then        '���������ۿƽ�����
        For i = 2 To 15
            If ResultEye.RowHidden(i) = False Then
                If Not (ResultEye.TextMatrix(i, 1) = "����" Or ResultEye.TextMatrix(i, 1) = "����" Or ResultEye.TextMatrix(i, 1) = "�ۿ�����") Then
                    lstrCheck = sub��ӵ�����(ResultEye.TextMatrix(i, 2), (ResultEye.TextMatrix(i, 1) & "-��"), lstrCheck)
                    lstrCheck = sub��ӵ�����(ResultEye.TextMatrix(i, 7), (ResultEye.TextMatrix(i, 1) & "-��"), lstrCheck)
                ElseIf ResultEye.TextMatrix(i, 1) = "����" Then
                    lstrCheck = sub��ӵ�����(ResultEye.TextMatrix(i, 3), "����-Զ����-��", lstrCheck)
                    lstrCheck = sub��ӵ�����(ResultEye.TextMatrix(i, 5), "����-������-��", lstrCheck)
                    lstrCheck = sub��ӵ�����(ResultEye.TextMatrix(i, 8), "����-Զ����-��", lstrCheck)
                    lstrCheck = sub��ӵ�����(ResultEye.TextMatrix(i, 10), "����-������-��", lstrCheck)
                ElseIf ResultEye.TextMatrix(i, 1) = "����" Then
                    lstrCheck = sub��ӵ�����(ResultEye.TextMatrix(i, 3), "����-Զ����-��", lstrCheck)
                    lstrCheck = sub��ӵ�����(ResultEye.TextMatrix(i, 5), "����-������-��", lstrCheck)
                    lstrCheck = sub��ӵ�����(ResultEye.TextMatrix(i, 8), "����-Զ����-��", lstrCheck)
                    lstrCheck = sub��ӵ�����(ResultEye.TextMatrix(i, 10), "����-������-��", lstrCheck)
                ElseIf ResultEye.TextMatrix(i, 1) = "�ۿ�����" Then
                    lstrCheck = sub��ӵ�����(ResultEye.TextMatrix(i, 2), "�ۿ�����", lstrCheck)
                End If
            End If
        Next i
    Else
        '�����˲��ӵĵ������������ù������� sub��ӵ����� �� subSave
            
        '���������۱����Ŀ���
        For i = 2 To 5
            lstrCheck = sub��ӵ�����(ResultEyeArmy.TextMatrix(i, 2), (ResultEyeArmy.TextMatrix(i, 1) & "-��"), lstrCheck)
            lstrCheck = sub��ӵ�����(ResultEyeArmy.TextMatrix(i, 6), (ResultEyeArmy.TextMatrix(i, 1) & "-��"), lstrCheck)
        Next i
            
        '�����۱��£��ֱ���3������Ŀ�Ľ����д
        For i = 7 To 11
            For j = 2 To 8
                If j <= 4 Then lstrCheck = sub��ӵ�����(ResultEyeArmy.TextMatrix(i, j), (ResultEyeArmy.TextMatrix(i, 1) & "-" & ResultEyeArmy.TextMatrix(6, j) & "-��"), lstrCheck)
                If j >= 6 Then lstrCheck = sub��ӵ�����(ResultEyeArmy.TextMatrix(i, j), (ResultEyeArmy.TextMatrix(i, 1) & "-" & ResultEyeArmy.TextMatrix(6, j) & "-��"), lstrCheck)
            Next j
        Next i
            
        '"���"�Ľ����д
        lstrCheck = sub��ӵ�����(ResultEyeArmy.TextMatrix(12, 2), "�ۿ����", lstrCheck)
    End If
        
    For i = 1 To 18
        If ResultEar.RowHidden(i) = False Then
            If i <= 5 Or i >= 16 Then
                lstrCheck = sub��ӵ�����(ResultEar.TextMatrix(i, ResultEar.cols - 1), ResultEar.TextMatrix(i, 0), lstrCheck)
            Else
                lstrCheck = sub��ӵ�����(ResultEar.TextMatrix(i, ResultEar.cols - 1), ResultEar.TextMatrix(i, 0) & "-" & ResultEar.TextMatrix(i, 1), lstrCheck)
            End If
        End If
    Next
    
     'lstrcheck�ַ������
    If (Not lstrCheck = "") And (Not ResultChanged = 2) Then
        isOk = MsgBox("������Ŀδ��д�����ȷ��������" & Chr(10) & "δ��д������¼�����ݿ⣡" & Chr(10) & Chr(10) & Trim(lstrCheck), vbOKCancel)
        If isOk = 2 Then
            Set lcolResult = Nothing
            Set lcolItem = Nothing
            ccmdUnfilledAllPass.Enabled = True
            ccmdAutoFill.Enabled = True
            ccmdSave.Enabled = True
            ctlb������.Buttons(2).Enabled = True
            ccmdClearResult.Enabled = True
            Exit Sub
        End If
    End If
        
    If ResultChanged = 2 Then
        isOk = MsgBox("�Ƿ񱣴�������Ա���޸Ľ����", vbOKCancel)
        If isOk = 1 Then
            subSaveBatch         '�����������ɹ���ʾ
        Else
            LoadPersonalInfoBatch (ctxt�������)
        End If
    Else
        subSaveBatch
    End If
    
    Set lcolResult = Nothing
    Set lcolItem = Nothing
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�������¼��", "frmHEENT_ResultInput", "ccmdSave_Click", 6666, lstrError, False
    
End Sub
'�������浽���ݿ�
Private Sub subSaveBatch()
    On Error GoTo errHandler
    
    If cgrdInfoBatch.rows < 1 Then
        MsgBox ("��ȷ��¼����Ա��Ŀ�Ƿ���ȷ��")
        Exit Sub
    End If
    Dim ccrpValue As Integer
    Dim ccrpI As Integer
    Dim isOk As Boolean
    Dim lstrTmp As String
    Dim lobjTmp As Object
    Dim barCode As Collection
    Dim lcolConclusion As String '��ٿƵ�������
    Dim i As Integer
    Set barCode = New Collection
    For i = 1 To cgrdInfoBatch.rows - 1
        barCode.Add cgrdInfoBatch.TextMatrix(i, 0)
    Next i
    
    '��ʾ���������
    ccrpI = barCode.Count
    ccrp����.Max = ccrpI * 2
    ccrp����.Visible = True
    ccrp����.Caption = "0%"
    ccrp����.Value = 0
    
    
    Set lobjTmp = CreateObject("ְҵ�������¼��.clsCommon")
    For i = 1 To barCode.Count
        isOk = lobjTmp.func���浥�������(barCode(i), mstrDoctorName, DTP¼������.Value, lcolItem, lcolResult, "ְҵ�����_�����Ϣ_��ٿ�")
        ccrp����.Caption = Int(i / ccrp����.Max * 100) & "%"
        ccrp����.Value = ccrp����.Value + 1
        If i = barCode.Count Then ccrpValue = Int(i / ccrp����.Max * 100)
    Next i
    
    If ResultChanged <> 3 Then ResultChanged = 1
    If isOk = True Then
        For i = 1 To barCode.Count
            '���浥����Ŀ��ҽ������
            lcolConclusion = ctxtConclun.Text
            pobjҵ�����.sub������д������ barCode(i), priDeptName, lcolConclusion, um�û����
            '01069 ���۾������ͼ����Ŀ��š��ڡ������Ŀ���á��ı����м�¼��
            lobjTmp.func������ͼƬ Picture3.Image, barCode(i), mstr���ͼƬ��Ŀ���, cdtpConclusionDate.Value
            If DrawState = 0 Then
                Call sub��ӵ�����("����", "��״�廷�漰����ͼ", "")
            ElseIf DrawState <> -1 Then
                Call sub��ӵ�����("������", "��״�廷�漰����ͼ", "")
            End If
            ccrp����.Caption = Int(i / ccrp����.Max * 100) + ccrpValue & "%"
            ccrp����.Value = ccrp����.Value + 1
            
            '2012-07-03 �ڵ�� ��
            '����һ���ֶ�"�޸���ʼʱ��"���޸ġ�ͬʱ�޸ĸÿ��ҵ������¼��״̬��
            pobjҵ�����.sub�޸���ʼʱ�� barCode(i), priDeptName
            pobjҵ�����.sub�޸Ľ��¼��״̬ barCode(i), priDeptNo, "2"
            pobjҵ�����.sub���¼���޸����״̬ barCode(i), "4"
            '2012-07-03 �ڵ�� ��
        Next i
        MsgBox ("��������ɹ���")
        subClear
    Else
        subClear
    End If
        ccrp����.Visible = False
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�������¼��", "frmHEENT_ResultInput", "subSave", 6666, lstrError, False
End Sub

'2012-07-03 �ڵ��
Sub sub��ȡϵͳ��Ź̶�����()
    '��ȡ����������
    Dim lobjRec As Object
    Set lobjRec = dafuncGetData("select getDate()")
    ctxtBarCode.Text = um����վ��� & um���������� & Format(lobjRec(0), "yyyy")
    Set lobjRec = Nothing
End Sub

'2012-07-14 �ڵ��
Sub sub���¿��޸Ľ����Ա�޸�״̬()
    Dim lobjRec As Object
    Dim strSQL As String
    Dim canModify As Boolean
    
    strSQL = "select ϵͳ���,�������״̬ from ְҵ�����_���������ݿ� where substring(�������״̬," & priDeptNo & ",1)='1' or substring(�������״̬," & priDeptNo & ",1)='2'"
    Set lobjRec = dafuncGetData(strSQL)
    If lobjRec.recordcount = 0 Then Exit Sub
    lobjRec.movefirst
    While lobjRec.EOF <> True
        canModify = pobjҵ�����.sub�Ƿ����޸�ʱ�䷶Χ��(lobjRec("ϵͳ���"), priDeptName, 8)
        If canModify = False Then Call pobjҵ�����.sub�޸Ľ��¼��״̬(lobjRec("ϵͳ���"), priDeptNo, "3")
        lobjRec.MoveNext
    Wend
End Sub

'2012-07-14 �ڵ��
Sub sub��ѯ�б���ʾ(ByVal coptIndex As Integer)
    mobjQueryResult.Filter = ""
    
    If mobjQueryResult.recordcount > 0 Then
    
        If SSTPersonalInfo.Tab = 0 Then
            If cchkSigResult(0).Value = 1 And cchkSigResult(1).Value = 0 Then
                mobjQueryResult.Filter = "��дʱ��<>null"
            ElseIf cchkSigResult(0).Value = 0 And cchkSigResult(1).Value = 1 Then
                mobjQueryResult.Filter = "��дʱ��=null"
            ElseIf cchkSigResult(0).Value = 0 And cchkSigResult(1).Value = 0 Then
                mobjQueryResult.Filter = "ϵͳ���='xxx'"
            Else
                mobjQueryResult.Filter = ""
            End If
        ElseIf SSTPersonalInfo.Tab = 1 Then
            If cchkBchResult(0).Value = 1 And cchkBchResult(1).Value = 0 Then
                mobjQueryResult.Filter = "��дʱ��<>null"
            ElseIf cchkBchResult(0).Value = 0 And cchkBchResult(1).Value = 1 Then
                mobjQueryResult.Filter = "��дʱ��=null"
            ElseIf cchkBchResult(0).Value = 0 And cchkBchResult(1).Value = 0 Then
                mobjQueryResult.Filter = "ϵͳ���='xxx'"
            Else
                mobjQueryResult.Filter = ""
            End If
        End If
        
        If mobjQueryResult.Filter <> "" And mobjQueryResult.Filter <> 0 And mobjQueryResult.Filter <> "ϵͳ���='xxx'" Then
            mobjQueryResult.Filter = mobjQueryResult.Filter & " and �������='" & coptClasses(coptIndex).Caption & "'"
        Else
            mobjQueryResult.Filter = "�������='" & coptClasses(coptIndex).Caption & "'"
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
