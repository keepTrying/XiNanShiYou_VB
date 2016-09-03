VERSION 5.00
Object = "{D9347025-9612-11D1-9D75-00C04FCC8CDC}#1.0#0"; "MSBCODE9.OCX"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frm打印体检清单 
   BackColor       =   &H00FFFFFF&
   Caption         =   "职业病体检清单"
   ClientHeight    =   10275
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   ScaleHeight     =   10275
   ScaleWidth      =   11400
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton ccmdPrint 
      Caption         =   "打印"
      Height          =   495
      Left            =   5400
      TabIndex        =   32
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton ccmdExit 
      Caption         =   "退出"
      Height          =   495
      Left            =   7200
      TabIndex        =   31
      Top             =   1440
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   1935
      Left            =   9000
      ScaleHeight     =   1875
      ScaleWidth      =   1515
      TabIndex        =   18
      Top             =   840
      Width           =   1575
   End
   Begin VSFlex8Ctl.VSFlexGrid cgrdItems 
      Height          =   495
      Index           =   1
      Left            =   960
      TabIndex        =   19
      Top             =   3960
      Width           =   9735
      _cx             =   2088780563
      _cy             =   2088764265
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
      BackColorBkg    =   16777215
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   2
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   0
      Cols            =   10
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   0
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
   Begin VSFlex8Ctl.VSFlexGrid cgrdItems 
      Height          =   495
      Index           =   2
      Left            =   960
      TabIndex        =   20
      Top             =   4800
      Width           =   9735
      _cx             =   2088780563
      _cy             =   2088764265
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
      BackColorBkg    =   16777215
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
      GridLines       =   2
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   0
      Cols            =   10
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   0
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
   Begin VSFlex8Ctl.VSFlexGrid cgrdItems 
      Height          =   495
      Index           =   3
      Left            =   960
      TabIndex        =   21
      Top             =   5400
      Width           =   9735
      _cx             =   2088780563
      _cy             =   2088764265
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
      BackColorBkg    =   16777215
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
      GridLines       =   2
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   0
      Cols            =   10
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   0
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
   Begin VSFlex8Ctl.VSFlexGrid cgrdItems 
      Height          =   495
      Index           =   4
      Left            =   960
      TabIndex        =   22
      Top             =   5880
      Width           =   9735
      _cx             =   2088780563
      _cy             =   2088764265
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
      BackColorBkg    =   16777215
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
      GridLines       =   2
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   0
      Cols            =   10
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   0
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
   Begin VSFlex8Ctl.VSFlexGrid cgrdItems 
      Height          =   495
      Index           =   5
      Left            =   960
      TabIndex        =   23
      Top             =   6480
      Width           =   9735
      _cx             =   2088780563
      _cy             =   2088764265
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
      BackColorBkg    =   16777215
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
      GridLines       =   2
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   0
      Cols            =   10
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   0
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
   Begin VSFlex8Ctl.VSFlexGrid cgrdItems 
      Height          =   495
      Index           =   6
      Left            =   960
      TabIndex        =   24
      Top             =   7080
      Width           =   9735
      _cx             =   2088780563
      _cy             =   2088764265
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
      BackColorBkg    =   16777215
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
      GridLines       =   2
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   0
      Cols            =   10
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   0
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
   Begin VSFlex8Ctl.VSFlexGrid cgrdItems 
      Height          =   495
      Index           =   7
      Left            =   960
      TabIndex        =   25
      Top             =   7680
      Width           =   9735
      _cx             =   2088780563
      _cy             =   2088764265
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
      BackColorBkg    =   16777215
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
      GridLines       =   2
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   0
      Cols            =   10
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   0
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
   Begin VSFlex8Ctl.VSFlexGrid cgrdItems 
      Height          =   495
      Index           =   8
      Left            =   960
      TabIndex        =   26
      Top             =   8280
      Width           =   9735
      _cx             =   2088780563
      _cy             =   2088764265
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
      BackColorBkg    =   16777215
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
      GridLines       =   2
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   0
      Cols            =   10
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   0
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
   Begin VSFlex8Ctl.VSFlexGrid cgrdItems 
      Height          =   495
      Index           =   9
      Left            =   960
      TabIndex        =   27
      Top             =   8880
      Width           =   9735
      _cx             =   2088780563
      _cy             =   2088764265
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
      BackColorBkg    =   16777215
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
      GridLines       =   2
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   0
      Cols            =   10
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   0
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
   Begin VSFlex8Ctl.VSFlexGrid cgrdItems 
      Height          =   495
      Index           =   10
      Left            =   960
      TabIndex        =   28
      Top             =   9480
      Width           =   9735
      _cx             =   2088780563
      _cy             =   2088764265
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
      BackColorBkg    =   16777215
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
      GridLines       =   2
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   0
      Cols            =   10
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   0
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
   Begin VSFlex8Ctl.VSFlexGrid cgrdItems 
      Height          =   495
      Index           =   11
      Left            =   960
      TabIndex        =   29
      Top             =   10080
      Width           =   9735
      _cx             =   2088780563
      _cy             =   2088764265
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
      BackColorBkg    =   16777215
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
      GridLines       =   2
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   0
      Cols            =   10
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   0
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
   Begin VSFlex8Ctl.VSFlexGrid cgrdItems 
      Height          =   495
      Index           =   12
      Left            =   960
      TabIndex        =   30
      Top             =   10680
      Width           =   9735
      _cx             =   2088780563
      _cy             =   2088764265
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
      BackColorBkg    =   16777215
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
      GridLines       =   2
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   0
      Cols            =   10
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   0
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
   Begin VSFlex8Ctl.VSFlexGrid cgrdItems 
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   33
      Top             =   2880
      Width           =   9735
      _cx             =   2088780563
      _cy             =   2088763842
      Appearance      =   0
      BorderStyle     =   0
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
      BackColorSel    =   16777215
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   2
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   0
      Cols            =   10
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   0
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
   Begin VSFlex8Ctl.VSFlexGrid cgrdItems 
      Height          =   495
      Index           =   13
      Left            =   960
      TabIndex        =   34
      Top             =   11280
      Width           =   9735
      _cx             =   2088780563
      _cy             =   2088764265
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
      BackColorSel    =   16777215
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   2
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   0
      Cols            =   10
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   0
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
   Begin VB.Line Line7 
      X1              =   5640
      X2              =   6240
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line6 
      X1              =   3960
      X2              =   4920
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line5 
      X1              =   1800
      X2              =   3120
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line4 
      X1              =   7080
      X2              =   8280
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line3 
      X1              =   5640
      X2              =   6240
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line2 
      X1              =   3960
      X2              =   4920
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line1 
      X1              =   1440
      X2              =   3120
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label clblWorkAge 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5880
      TabIndex        =   17
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFFFFF&
      Caption         =   "工龄："
      Height          =   255
      Left            =   5160
      TabIndex        =   16
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label clblProfession 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4080
      TabIndex        =   15
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFFFF&
      Caption         =   "工种："
      Height          =   255
      Left            =   3480
      TabIndex        =   14
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label clblDegree 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   13
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFFF&
      Caption         =   "文化程度："
      Height          =   255
      Left            =   960
      TabIndex        =   12
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label clblNationality 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7200
      TabIndex        =   11
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "民族："
      Height          =   255
      Left            =   6600
      TabIndex        =   10
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label clblAge 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5760
      TabIndex        =   9
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "年龄："
      Height          =   255
      Left            =   5160
      TabIndex        =   8
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label clblSex 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4080
      TabIndex        =   7
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "性别："
      Height          =   255
      Left            =   3480
      TabIndex        =   6
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label clblName 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "姓名："
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   2160
      Width           =   615
   End
   Begin BARCODELibCtl.BarCodeCtrl BarCode 
      Height          =   735
      Left            =   2040
      TabIndex        =   3
      Top             =   1200
      Width           =   3015
      Style           =   7
      SubStyle        =   0
      Validation      =   0
      LineWeight      =   3
      Direction       =   0
      ShowData        =   1
      Value           =   "123456 Code-128"
      ForeColor       =   0
      BackColor       =   16777215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "体检编号："
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "职业病体检清单"
      Height          =   495
      Left            =   4320
      TabIndex        =   1
      Top             =   600
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "四川省疾病预防控制中心"
      Height          =   495
      Left            =   3360
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "frm打印体检清单"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''
'2012-06-18 于登淼
'添加整个窗体，用来实现单个体检人员体检清单赋值与初始化。
'预览与打印功能在其它窗口实现。
'''''''''''''''

Option Explicit

Public sysNo As String
Private paraDept(0 To 13) As String '记录每个科室体检项目
Private numInOneLine As Integer     '每行填写多少个项目(不包括医师签名)

Private Sub ccmdExit_Click()
    Unload Me
End Sub

Private Sub ccmdPrint_Click()
    Me.ccmdPrint.Visible = False
    Me.ccmdExit.Visible = False
    Form_Load
    
    Dim devPrinter As Printer
     For Each devPrinter In Printers
        If devPrinter.DeviceName = "清单打印机" Then
           '设定为系统缺省打印机。
          Set Printer = devPrinter
           ' 终止查找打印机。
          Exit For
        End If
     Next

    Me.PrintForm
    Me.ccmdPrint.Visible = True
    Me.ccmdExit.Visible = True
End Sub

Private Sub Form_Load()
 Dim devPrinter As Printer
 Dim bqprt As String
 Dim qdprt As String
 bqprt = ""
 qdprt = ""
     For Each devPrinter In Printers
        If devPrinter.DeviceName = "标签打印机" Then
          bqprt = devPrinter
        End If
        If devPrinter.DeviceName = "清单打印机" Then
          qdprt = devPrinter
        End If
     Next
     
   If bqprt = "标签打印机" And qdprt = "清单打印机" Then
   Else
        Dim tips
         tips = MsgBox("没有设置标签打印机！", vbOKOnly + vbCritical, "提示")
      Exit Sub
         
   End If
    subFontInit
    subPersonalInfoInit
    subDeptInit
    subItemInit
    subFormatInit
End Sub

'2012-06-18 于登淼
'控制标题字体大小。label1与label2右侧对齐时，字体刚好也对齐
Sub subFontInit()
    Label1.FontSize = 24
    Label2.FontSize = 24
End Sub

Sub subPersonalInfoInit()
    Dim lobjRec As Object
    Dim strSQL As String

    If sysNo = "" Then sysNo = "0011201200003943" '''''''''''''测试行''''''''''''
    BarCode.Value = sysNo
    
    strSQL = "select * from 职业病体检_体检人员基本信息表 where 系统编号='" & sysNo & "'"
    Set lobjRec = dafuncGetData(strSQL)
    
    '载入文字类基本信息
    clblName.Caption = lobjRec("姓名")
    clblSex.Caption = lobjRec("性别")
    clblAge.Caption = lobjRec("年龄")
    clblNationality.Caption = lobjRec("民族")
    clblDegree.Caption = lobjRec("文化程度")
    clblProfession.Caption = lobjRec("现工种")
    clblWorkAge.Caption = lobjRec("工龄")
    
    '载入体检照片（现场拍照，不是身份证照片）
    Set lobjRec = CreateObject("职业病对象.clspersonexamed")
    lobjRec.系统编号 = sysNo
    Picture1.Picture = lobjRec.像片
    
    Set lobjRec = Nothing
End Sub

Sub subFormatInit()
    Dim i, j, k, curDept As Integer
    Dim preTop, preHeight As Integer
    Dim curTop, curHeight As Integer
    
    '0科室，受检者个人信息项目。表格大小与位置设定
    With cgrdItems(0)
        For j = 0 To .Cols - 1  '显示的时候，让水平、垂直滚动条不可用。
            .ColWidth(j) = (.Width + 500) / (numInOneLine + 2)
        Next j
        preTop = .Top           '表格位置初始化，都以cgrditems(0)的为准
        preHeight = .Height - 30
    End With
        
    '其它科室体检项目部分。表格大小与位置设定
    For i = 1 To UBound(paraDept)
        cgrdItems(i).Visible = True
        curTop = preTop + preHeight
        If paraDept(i) = "" Then
            curDept = 0
        Else
            curDept = CInt(paraDept(i))
        End If
        If curDept > 0 Then
            curHeight = cgrdItems(i).RowHeight(0) * Int(curDept / (numInOneLine)) + 20  '手动调整格式
            If curDept Mod (numInOneLine) <> 0 Then curHeight = curHeight + cgrdItems(i).RowHeight(0)
            preTop = curTop
            preHeight = curHeight
            With cgrdItems(i)
                '.FixedCols = 1
                .Top = curTop
                .Height = curHeight
                For j = 0 To .Cols - 1  '显示的时候，让水平、垂直滚动条不可用。
                    .ColWidth(j) = (.Width + 500) / (numInOneLine + 2)
                Next j
            End With
        Else
            cgrdItems(i).Visible = False
        End If
    Next
    
    '所有科室表头和医师签名处设定
    For i = 0 To UBound(paraDept)
        With cgrdItems(i)
            .MergeCompare = flexMCIncludeNulls
            .MergeCells = flexMergeFree
            If i = 0 Then
                '.FixedRows = 1
                For j = 0 To numInOneLine: .ColAlignment(j) = flexAlignCenterCenter: Next
            End If
            For j = 0 To .Rows - 1: cgrdItems(i).TextMatrix(j, numInOneLine + 1) = " ": Next
            For j = 0 To numInOneLine + 1: .MergeCol(j) = True: Next
            For j = 0 To .Rows - 1: .MergeRow(j) = True: Next
            .ColAlignment(0) = flexAlignCenterCenter
        End With
    Next i
                
End Sub

Sub subDeptInit()
    Dim lobjRec As Object
    Dim strSQL As String
    Dim i As Integer
    Dim pCnt As Integer
    Dim preDept, curDept As Integer
    
    strSQL = "select 体检项目 from 职业病体检_体检结果视图 where 系统编号='" & sysNo & "'"
    Set lobjRec = dafuncGetData(strSQL)
    lobjRec.movefirst
    For i = 0 To lobjRec.RecordCount - 1
        If i = 0 Then
            preDept = CInt(Left(lobjRec("体检项目"), 2))
            pCnt = 0
        End If
        curDept = CInt(Left(lobjRec("体检项目"), 2))
        If preDept <> curDept Then
            paraDept(preDept) = pCnt
            pCnt = 0
            preDept = curDept
        End If
        pCnt = pCnt + 1
        lobjRec.movenext
        If lobjRec.EOF = True Then paraDept(curDept) = pCnt
    Next
End Sub

Sub subItemInit()
    Dim lobjRec As Object
    Dim strSQL As String
    Dim i, j, k, idx As Integer
    Dim curDept As String
    Dim preDept As String
    Dim crow, ccol As Integer
    
    '受检者个人信息项目部分
    numInOneLine = 4
    cgrdItems(0).Rows = 1
    cgrdItems(0).TextMatrix(0, 0) = "体检科室"
    For k = 1 To numInOneLine: cgrdItems(0).TextMatrix(0, k) = "体检项目": Next
    cgrdItems(0).TextMatrix(0, numInOneLine + 1) = "科室医师签字"
    'cgrdItems(0).TextMatrix(1, 1) = "受检者个人信息录入"
    'cgrdItems(0).TextMatrix(1, 0) = "受检者个人信息"    '科室名称
    
    '其它科室项目部分
    strSQL = "select a.名称,b.体检项目 from 职业病体检_体检项目设置表 a, 职业病体检_体检结果视图 b where b.系统编号='" & sysNo & "' and a.编码=b.体检项目 order by b.体检项目"
    Set lobjRec = dafuncGetData(strSQL)
    
    idx = 0
    lobjRec.movefirst
    For k = 0 To lobjRec.RecordCount - 1
        If k = 0 Then
            preDept = CInt(Left(lobjRec("体检项目"), 2))
        End If
        curDept = CInt(Left(lobjRec("体检项目"), 2))
        If curDept <> preDept Then
            idx = 0
            preDept = curDept
        End If
        crow = Int(idx / numInOneLine)
        ccol = idx Mod numInOneLine
        If idx = 0 Then cgrdItems(curDept).Cols = numInOneLine + 2
        If ccol = 0 Then cgrdItems(curDept).Rows = cgrdItems(curDept).Rows + 1
        cgrdItems(curDept).TextMatrix(crow, ccol + 1) = lobjRec("名称")
        idx = idx + 1
        lobjRec.movenext
    Next
    
    '其它科室名称
    Dim lobjDict As Object
    Dim lobjDept As Object
    Set lobjDict = CreateObject("字典管理.clsDictionary")
    Set lobjDept = lobjDict.fetch("职业病体检科室字典")
    lobjDept.movefirst
    For i = 1 To lobjDept.RecordCount
        If Right(lobjDept("名称"), 1) = "科" Then
            If Len(lobjDept("名称")) > 6 Then
                With cgrdItems(i)
                    For j = 0 To .Rows - 1
                        cgrdItems(i).TextMatrix(j, 0) = Left(lobjDept("名称"), Len(lobjDept("名称")) - 3) & Chr(10) & Right(lobjDept("名称"), 3)
                    Next
                End With
            Else
                With cgrdItems(i)
                    For j = 0 To .Rows - 1
                        cgrdItems(i).TextMatrix(j, 0) = lobjDept("名称")
                    Next
                End With
            End If
        End If
        lobjDept.movenext
    Next
    Set lobjDict = Nothing
    Set lobjDept = Nothing
End Sub

Public Function subPrint()
    ccmdPrint_Click
End Function
