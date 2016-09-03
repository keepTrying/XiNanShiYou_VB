VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmresultreport 
   Caption         =   "体检体格报告"
   ClientHeight    =   9465
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15645
   LinkTopic       =   "体检体格报告"
   ScaleHeight     =   9465
   ScaleWidth      =   15645
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   4560
      TabIndex        =   17
      Top             =   8520
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      Height          =   1815
      Left            =   7560
      TabIndex        =   10
      Text            =   "Text4"
      Top             =   7440
      Width           =   7815
   End
   Begin VB.TextBox Text3 
      Height          =   1815
      Left            =   7560
      TabIndex        =   9
      Text            =   "Text3"
      Top             =   5280
      Width           =   7815
   End
   Begin VB.TextBox Text2 
      Height          =   1935
      Left            =   7560
      TabIndex        =   8
      Text            =   "Text2"
      Top             =   2760
      Width           =   7815
   End
   Begin VB.TextBox Text1 
      Height          =   1935
      Left            =   7560
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   360
      Width           =   7815
   End
   Begin VB.TextBox txt_time 
      Height          =   270
      Left            =   1200
      TabIndex        =   6
      ToolTipText     =   "可点击右边按钮选择公司。"
      Top             =   960
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "查询"
      Height          =   495
      Left            =   4560
      TabIndex        =   5
      Top             =   960
      Width           =   2055
   End
   Begin VB.CommandButton cmd_selectcompany 
      Caption         =   "选择公司"
      Height          =   495
      Left            =   4560
      TabIndex        =   2
      Top             =   360
      Width           =   2055
   End
   Begin VB.TextBox txt_companyname 
      Height          =   270
      Left            =   1200
      TabIndex        =   1
      ToolTipText     =   "可点击右边按钮选择公司。"
      Top             =   480
      Width           =   3015
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2220
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   3916
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   60162049
      CurrentDate     =   42115
   End
   Begin VSFlex8Ctl.VSFlexGrid cgrdDept 
      Height          =   4335
      Left            =   120
      TabIndex        =   16
      ToolTipText     =   "双击查看各项结论"
      Top             =   4080
      Width           =   7215
      _cx             =   12726
      _cy             =   7646
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
      AllowUserResizing=   0
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
      FormatString    =   "科室|文字结论|医师姓名"
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
   Begin VB.Label Label7 
      Caption         =   "参加体检人员："
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "建议："
      Height          =   255
      Left            =   7560
      TabIndex        =   14
      Top             =   7200
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "体检结论："
      Height          =   255
      Left            =   7560
      TabIndex        =   13
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "体检结果："
      Height          =   255
      Left            =   7560
      TabIndex        =   12
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "体检简述："
      Height          =   255
      Left            =   7560
      TabIndex        =   11
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "公司名称："
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "体检日期："
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   975
   End
End
Attribute VB_Name = "frmresultreport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
