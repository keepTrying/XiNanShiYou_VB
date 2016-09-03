VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmFinalConclusion 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "最终结论录入窗口"
   ClientHeight    =   11880
   ClientLeft      =   1635
   ClientTop       =   2070
   ClientWidth     =   16425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11880
   ScaleWidth      =   16425
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox libPicture 
      Height          =   615
      Left            =   480
      ScaleHeight     =   555
      ScaleWidth      =   1755
      TabIndex        =   44
      Top             =   10320
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Height          =   11895
      Left            =   0
      ScaleHeight     =   11835
      ScaleWidth      =   16395
      TabIndex        =   0
      Top             =   0
      Width           =   16455
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   12255
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   16575
         Begin VB.Frame fraDeptItem 
            Caption         =   "科室与项目结果、结论"
            Height          =   5415
            Left            =   360
            TabIndex        =   18
            Top             =   6480
            Width           =   16215
            Begin VB.TextBox Text结果描述 
               BackColor       =   &H00FFFFFF&
               Height          =   735
               Left            =   7800
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   55
               Top             =   4080
               Width           =   3135
            End
            Begin VB.CheckBox cchkUnfilled 
               BackColor       =   &H00FFC0C0&
               Caption         =   "忽略未填结果项"
               Height          =   255
               Left            =   13800
               TabIndex        =   21
               Top             =   5040
               Width           =   1695
            End
            Begin VB.TextBox ctxtDetpConclusion 
               BackColor       =   &H00C0C0FF&
               ForeColor       =   &H00FF0000&
               Height          =   1575
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   20
               Top             =   3720
               Width           =   7695
            End
            Begin VB.CheckBox cchkAbnormal 
               BackColor       =   &H00FFC0C0&
               Caption         =   "结果仅显示不正常项"
               Height          =   255
               Left            =   11520
               TabIndex        =   19
               Top             =   5040
               Width           =   1935
            End
            Begin VSFlex8Ctl.VSFlexGrid cgrdItem 
               Height          =   4695
               Left            =   11400
               TabIndex        =   22
               Top             =   240
               Width           =   4695
               _cx             =   2088771673
               _cy             =   2088771673
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
               Cols            =   5
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   "项目编号|体检项目|体检结果|体检医师|单项结论"
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
            Begin VSFlex8Ctl.VSFlexGrid cgrdDept 
               Height          =   3375
               Left            =   120
               TabIndex        =   23
               ToolTipText     =   "双击查看各项结论"
               Top             =   240
               Width           =   11175
               _cx             =   2088783103
               _cy             =   2088769345
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
            Begin VB.Label Label9 
               Caption         =   "科室结果描述："
               Height          =   255
               Left            =   7920
               TabIndex        =   56
               Top             =   3840
               Visible         =   0   'False
               Width           =   2415
            End
         End
         Begin VB.Frame fraFinal 
            Caption         =   "填写最终结论"
            Height          =   5895
            Left            =   12480
            TabIndex        =   24
            Top             =   600
            Width           =   3855
            Begin VB.CommandButton cmd选择复查项目 
               Caption         =   "选择复查项目"
               Height          =   375
               Left            =   1800
               TabIndex        =   63
               Top             =   4440
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.CommandButton comd保存意见 
               Caption         =   "保存意见"
               Height          =   375
               Left            =   2640
               TabIndex        =   58
               Top             =   2040
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.CommandButton Command1 
               Caption         =   "添加选中项到体检结果输入框"
               Height          =   420
               Left            =   120
               TabIndex        =   53
               Top             =   5400
               Width           =   3255
            End
            Begin VB.TextBox ctxtReviewItem 
               Height          =   495
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   49
               Top             =   4800
               Width           =   3615
            End
            Begin VB.CommandButton Cmd结论模版 
               Caption         =   "处理意见模版"
               Height          =   345
               Index           =   1
               Left            =   1200
               TabIndex        =   46
               Top             =   2040
               Width           =   1335
            End
            Begin VB.CommandButton Cmd结论模版 
               Caption         =   "结论模板"
               Height          =   315
               Index           =   0
               Left            =   1920
               TabIndex        =   45
               Top             =   720
               Width           =   1455
            End
            Begin VB.OptionButton cchk标准 
               Caption         =   "需复查"
               ForeColor       =   &H000000FF&
               Height          =   255
               Index           =   1
               Left            =   2400
               TabIndex        =   42
               Top             =   240
               Width           =   975
            End
            Begin VB.OptionButton cchk标准 
               Caption         =   "不复查"
               Height          =   255
               Index           =   0
               Left            =   1200
               TabIndex        =   41
               Top             =   240
               Width           =   855
            End
            Begin VB.ComboBox cbox结论模板 
               Height          =   300
               IMEMode         =   2  'OFF
               Left            =   1200
               Style           =   2  'Dropdown List
               TabIndex        =   34
               Top             =   720
               Visible         =   0   'False
               Width           =   2535
            End
            Begin VB.TextBox ctxtDiagnose 
               Height          =   975
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   26
               Top             =   2400
               Width           =   3615
            End
            Begin VB.TextBox ctxtConclusion 
               Height          =   855
               Left            =   120
               MaxLength       =   500
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   25
               Top             =   1080
               Width           =   3615
            End
            Begin MSComCtl2.DTPicker cdtpConclusion 
               Height          =   255
               Left            =   1440
               TabIndex        =   27
               Top             =   3360
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   450
               _Version        =   393216
               Format          =   59965440
               CurrentDate     =   41009
            End
            Begin VB.TextBox ctxtReview 
               Height          =   615
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   47
               Top             =   3840
               Width           =   3615
            End
            Begin VB.Label Label7 
               BackColor       =   &H00FFC0FF&
               Caption         =   "复查项目："
               Height          =   255
               Left            =   120
               TabIndex        =   50
               Top             =   4560
               Width           =   1095
            End
            Begin VB.Label Label6 
               BackColor       =   &H00FFC0FF&
               Caption         =   "复查原因："
               Height          =   255
               Left            =   120
               TabIndex        =   48
               Top             =   3600
               Width           =   1095
            End
            Begin VB.Label Label2 
               BackColor       =   &H00FFC0FF&
               Caption         =   "模板筛选："
               Height          =   255
               Left            =   120
               TabIndex        =   35
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label4 
               BackColor       =   &H00FFC0FF&
               Caption         =   "处理意见："
               Height          =   255
               Left            =   120
               TabIndex        =   31
               Top             =   2040
               Width           =   1095
            End
            Begin VB.Label Label3 
               BackColor       =   &H00FFC0FF&
               Caption         =   "体检结果："
               Height          =   255
               Left            =   120
               TabIndex        =   30
               Top             =   720
               Width           =   975
            End
            Begin VB.Label Label1 
               BackColor       =   &H00FFC0FF&
               Caption         =   "结论日期："
               Height          =   255
               Left            =   120
               TabIndex        =   28
               Top             =   3360
               Width           =   1095
            End
            Begin VB.Label llabDoctor 
               BackColor       =   &H00FFC0FF&
               Caption         =   "主治医师："
               Height          =   255
               Left            =   120
               TabIndex        =   29
               Top             =   3480
               Visible         =   0   'False
               Width           =   1935
            End
         End
         Begin VB.Frame fraPerson 
            Caption         =   "查询人员信息"
            Height          =   6615
            Left            =   3960
            TabIndex        =   15
            Top             =   840
            Width           =   8415
            Begin VSFlex8Ctl.VSFlexGrid cgrdInfo 
               Height          =   4695
               Left            =   120
               TabIndex        =   16
               Top             =   720
               Width           =   8175
               _cx             =   14420
               _cy             =   8281
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
               Cols            =   0
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
            Begin VB.Label Label11 
               Caption         =   "总人数："
               Height          =   255
               Left            =   120
               TabIndex        =   60
               Top             =   480
               Width           =   735
            End
            Begin VB.Label Label10 
               Caption         =   "Label10"
               Height          =   255
               Left            =   960
               TabIndex        =   59
               Top             =   480
               Width           =   1695
            End
            Begin VB.Label Label5 
               BackColor       =   &H00C0FFFF&
               Caption         =   "按保存，选中行都保存为当前的结论和处理意见；按BackSpace移除选中行"
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   120
               TabIndex        =   17
               ToolTipText     =   "双击查看科室结论"
               Top             =   240
               Width           =   8175
            End
         End
         Begin VB.Frame fraQuery 
            Caption         =   "筛选体检人员"
            Height          =   5535
            Left            =   240
            TabIndex        =   2
            Top             =   840
            Width           =   3735
            Begin VB.CommandButton Com评残 
               Caption         =   "评残"
               Height          =   375
               Left            =   120
               TabIndex        =   62
               Top             =   3480
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.CommandButton Com退回 
               Caption         =   "退回复核"
               Height          =   375
               Left            =   2640
               TabIndex        =   61
               Top             =   3480
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.OptionButton coptType 
               BackColor       =   &H00FFFFC0&
               Caption         =   "未下建议"
               Height          =   255
               Index           =   6
               Left            =   2160
               TabIndex        =   57
               Top             =   4080
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.CommandButton Command2 
               Caption         =   "查看信息"
               Enabled         =   0   'False
               Height          =   375
               Left            =   2160
               TabIndex        =   54
               Top             =   4920
               Width           =   1095
            End
            Begin MSComCtl2.DTPicker cdtpDateTo 
               Height          =   255
               Left            =   1560
               TabIndex        =   14
               Top             =   1440
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   450
               _Version        =   393216
               Format          =   59965440
               CurrentDate     =   40969
            End
            Begin VB.CommandButton ccmdQuery 
               Caption         =   "查 询"
               Height          =   375
               Left            =   480
               TabIndex        =   3
               Top             =   4920
               Width           =   1095
            End
            Begin VB.CheckBox cchkTemplate 
               BackColor       =   &H00C0FFC0&
               Caption         =   "体检表模板"
               Height          =   300
               Left            =   240
               TabIndex        =   43
               Top             =   960
               Width           =   1335
            End
            Begin VB.OptionButton coptType 
               BackColor       =   &H00FFFFC0&
               Caption         =   "待复查"
               Height          =   255
               Index           =   5
               Left            =   2160
               TabIndex        =   40
               Top             =   5040
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.OptionButton coptType 
               BackColor       =   &H00FFFFC0&
               Caption         =   "已发报告"
               Height          =   255
               Index           =   4
               Left            =   480
               TabIndex        =   39
               Top             =   5040
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.OptionButton coptType 
               BackColor       =   &H00FFFFC0&
               Caption         =   "已复核"
               Height          =   255
               Index           =   3
               Left            =   480
               TabIndex        =   38
               Top             =   4560
               Width           =   1095
            End
            Begin VB.OptionButton coptType 
               BackColor       =   &H00FFFFC0&
               Caption         =   "待复核"
               Height          =   255
               Index           =   2
               Left            =   2160
               TabIndex        =   37
               Top             =   4560
               Width           =   1095
            End
            Begin VB.OptionButton coptType 
               BackColor       =   &H00FFFFC0&
               Caption         =   "未下结论"
               Height          =   255
               Index           =   1
               Left            =   2160
               TabIndex        =   36
               Top             =   4080
               Width           =   1095
            End
            Begin VB.ComboBox Ccmb体检人类别 
               Enabled         =   0   'False
               Height          =   300
               Left            =   2040
               TabIndex        =   13
               Text            =   "体检类别"
               Top             =   360
               Width           =   1575
            End
            Begin VB.ComboBox ccmb体检人类型 
               Enabled         =   0   'False
               Height          =   300
               Left            =   120
               TabIndex        =   12
               Text            =   "体检人员类型"
               Top             =   360
               Width           =   1695
            End
            Begin VB.OptionButton coptType 
               BackColor       =   &H00FFFFC0&
               Caption         =   "体检中"
               Height          =   255
               Index           =   0
               Left            =   480
               TabIndex        =   11
               Top             =   4080
               Width           =   1095
            End
            Begin VB.ComboBox ccmbTemplate 
               Height          =   300
               Left            =   1560
               TabIndex        =   10
               Top             =   960
               Width           =   2055
            End
            Begin VB.CheckBox cchkDate 
               BackColor       =   &H00C0FFC0&
               Caption         =   "体检日期 从"
               Height          =   255
               Left            =   240
               TabIndex        =   9
               Top             =   1440
               Value           =   1  'Checked
               Width           =   1335
            End
            Begin VB.CheckBox cchkCompanyName 
               BackColor       =   &H00C0FFC0&
               Caption         =   "单位名称"
               Height          =   255
               Left            =   240
               TabIndex        =   8
               Top             =   3000
               Width           =   1335
            End
            Begin VB.TextBox ctxtCompanyName 
               Height          =   270
               Left            =   1560
               TabIndex        =   7
               Top             =   3000
               Width           =   2055
            End
            Begin VB.CheckBox cchkBarCode 
               BackColor       =   &H00C0FFC0&
               Caption         =   "体检条码号"
               Height          =   255
               Left            =   240
               TabIndex        =   6
               Top             =   2520
               Width           =   1335
            End
            Begin VB.TextBox ctxtBarCode 
               Height          =   270
               Left            =   1560
               TabIndex        =   5
               Top             =   2520
               Width           =   2055
            End
            Begin VB.CommandButton ccmdLocate 
               Caption         =   "单位定位"
               Height          =   375
               Left            =   1200
               TabIndex        =   4
               Top             =   3480
               Width           =   1095
            End
            Begin MSComCtl2.DTPicker cdtpDateFrom 
               Height          =   255
               Left            =   1560
               TabIndex        =   51
               Top             =   1920
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   450
               _Version        =   393216
               Format          =   59965440
               CurrentDate     =   40969
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0FFC0&
               Caption         =   "到"
               Height          =   180
               Left            =   1320
               TabIndex        =   52
               Top             =   1920
               Width           =   180
            End
         End
         Begin MSComctlLib.Toolbar ctlb工具栏 
            Height          =   540
            Left            =   120
            TabIndex        =   32
            Top             =   0
            Width           =   11895
            _ExtentX        =   20981
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
            Begin VB.CheckBox Check1 
               Caption         =   "Check1"
               Height          =   255
               Index           =   1000
               Left            =   4200
               TabIndex        =   33
               Top             =   1200
               Width           =   1215
            End
            Begin MSComDlg.CommonDialog ccmdFile 
               Left            =   960
               Top             =   0
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
               Flags           =   6148
            End
         End
      End
   End
End
Attribute VB_Name = "frmFinalConclusion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'2012-04-10 于登淼
'增加 最终结论录入窗体，及相应部件功能
'最低操作权限：1、查询符合条件人员的各个科室结论和各科各项体检结果。
'              2、但修改不能保存，也不能打印报告。

Option Explicit
Public mblnInUse As Boolean
Private WithEvents mobjGUI As cls界面通用对象
Attribute mobjGUI.VB_VarHelpID = -1
Private mstr体检单号 As String
'Private mstr系统编号 As String
Private mlobjRec As Object
Private mstr权限标志 As Boolean

'查询结果
Private mstrDoctorName As String
Private mobjQueryResult As Object
Private mcolIndex As New Collection
Private indX, indY As Integer       '记录鼠标点击vsflexgrid的坐标。
Private resql As String     '记录每次查询的sql
 
'该界面共用对象
Private pobj体检表模板 As Object
Private pobj体检 As Object
Private pobj体检结果业务 As Object
Private pobj科室 As Object
Private pstrPerson As String        '当前单个体检人员系统编号,cgrdInfo双击后更新
Private pobjItem As Object

Private mstrSearchString As String
Private mstr体检表名称 As String

'2012-06-21 于登淼
'标记当前选中的体检人员体检状态。
'主要用于判断 未下结论、已下结论、已审核、已发放报告、待复查。
Private mstrState As String

'2012-08-22 于登淼 ↓
'添加科室变量
Private pobjDept As Object
'2012-08-22 于登淼 ↑

'2012-04-10 于登淼
'功能：返回当前窗体是否已经加载标志。这是系统平台所要求的。
Public Property Get pblnInUse() As Boolean
    On Error Resume Next
    pblnInUse = mblnInUse
End Property
'功能：选择结论模板
'作者：翁乔
'时间：2012-06-01
Private Sub cbox结论模板_Click()
    '2012-06-25 于登淼 ↓
    '模板结论需要回车后加在末尾
    ctxtConclusion.Text = ctxtConclusion.Text & cbox结论模板.Text
    ctxtDiagnose.Text = ctxtDiagnose.Text & cbox结论模板.Text    '2015-10-16
    '2012-06-25 于登淼 ↑
End Sub

'2012-04-12 于登淼
'仅显示不合格项，为了方便主治医师下最终结论时，查看结果。
Private Sub cchkAbnormal_Click()
    Dim i As Integer
    If cchkAbnormal.Value = 1 Then      '仅显示不正常项
        For i = 1 To cgrdItem.rows - 1: cgrdItem.RowHidden(i) = False: Next
        For i = 1 To cgrdItem.rows - 1  '默认第4列为单项结论
            If cgrdItem.TextMatrix(i, 4) = "合格" Then cgrdItem.RowHidden(i) = True
        Next
    Else
        '修改人：张令 2012.12.12        ↓↓
        '当“结果仅显示不正常项”未勾选时，重新显示数据。
'        If cchkUnfilled.Value = 0 Then Exit Sub
'        cgrdItem.Clear
'        Set cgrdItem.DataSource = pobjItem
'        For i = 1 To cgrdItem.rows - 1
'            If cgrdItem.TextMatrix(i, 4) = "不合格" Then cgrdItem.RowHidden(i) = False
'        Next
        With cgrdItem
            cgrdItem.Clear
            Set .DataSource = pobjItem
    '        .Col = 0
            .Sort = flexSortStringAscending
            .AutoSize 0, .cols - 1, 0, 0
            .SelectionMode = flexSelectionListBox
            .AllowSelection = False
        End With
    End If
    '修改人：张令 2012.12.12       ↓↓
    '当“忽略未填结果项”为勾选，“结果仅显示不正常项”为未勾选时执行cchkUnfilled_Click
    If cchkUnfilled.Value = 1 And cchkAbnormal.Value = 0 Then
        cchkUnfilled_Click
    End If
    '修改人：张令 2012.12.12       ↑↑
End Sub


Private Sub cchkTemplate_Click()
    If cchkTemplate.Value = 1 Then
        ccmb体检人类型.Enabled = True
        Ccmb体检人类别.Enabled = True
    Else
        ccmb体检人类型.Enabled = False
        Ccmb体检人类别.Enabled = False
    End If
End Sub

'2012-04-12 于登淼
'去掉没有填写结果的项，为了方便主治医师下最终结论时，查看结果。
Private Sub cchkUnfilled_Click()
    Dim i As Integer
    If cchkUnfilled.Value = 1 Then      '忽略未填写项
        For i = 1 To cgrdItem.rows - 1  '默认第1列为体检结果
            If cgrdItem.TextMatrix(i, 2) = "" Or (IsNull(cgrdItem.TextMatrix(i, 2)) = True) Then cgrdItem.RowHidden(i) = True
        Next
    Else
        '修改人：张令 2012.12.12        ↓↓
        '当“忽略未填结果项”未勾选时，重新显示数据。
'        If cchkAbnormal.Value = 1 Then Exit Sub
'        cgrdItem.Clear
'        Set cgrdItem.DataSource = pobjItem
'        For i = 1 To cgrdItem.rows - 1
'            If cgrdItem.TextMatrix(i, 2) <> "" Then cgrdItem.RowHidden(i) = True
'        Next
'        cchkAbnormal_Click
        With cgrdItem
            cgrdItem.Clear
            Set .DataSource = pobjItem
    '        .Col = 0
            .Sort = flexSortStringAscending
            .AutoSize 0, .cols - 1, 0, 0
            .SelectionMode = flexSelectionListBox
            .AllowSelection = False
        End With
        '修改人：张令 2012.12.12       ↑↑
    End If
    '修改人：张令 2012.12.12       ↓↓
    '当“忽略未填结果项”为未勾选，“结果仅显示不正常项”为勾选时执行cchkAbnormal_Click
    If cchkAbnormal.Value = 1 And cchkUnfilled.Value = 0 Then
        cchkAbnormal_Click
    End If
    '修改人：张令 2012.12.12       ↑↑
End Sub

'功能：实现结论模板的筛选
'作者：翁乔
'时间：2012-05-31
Private Sub cchk标准_Click(Index As Integer)
    
    Dim lobj结论 As Object
    Dim pub结论 As Object
    Dim i As Integer
    Dim sql As String 'func读取所有最终结论模板
    Set pub结论 = CreateObject("职业病对象.clsConclusionSet")
    
    '2012-07-03 于登淼 ↓
    '将复选框改为单选框，判断值与判断条件稍微改动
    'If cchk标准(0).Value = 1 And cchk标准(1).Value = 0 Then
    If cchk标准(0).Value = True Then
        sql = "select * from 系统管理_字典_体检结论模板表 where 结论标准='合格' and 科室编号=(select 编号 from 系统管理_字典_字典内容表 where ID='84' and 名称='最终结论录入')"
        Set lobj结论 = pub结论.func读取所有最终结论模板(sql)
        cbox结论模板.Clear
        'lobj结论.MoveFirst
        For i = 1 To lobj结论.RecordCount
            cbox结论模板.AddItem lobj结论("结论模板")
            lobj结论.MoveNext
        Next i
        ctxtReview.Text = ""
        ctxtReviewItem.Text = ""
    End If
    'If cchk标准(0).Value = 0 And cchk标准(1).Value = 1 Then
    If cchk标准(1).Value = True Then
        sql = "select * from 系统管理_字典_体检结论模板表 where 结论标准='不合格' and 科室编号=(select 编号 from 系统管理_字典_字典内容表 where ID='84' and 名称='最终结论录入')"
        Set lobj结论 = pub结论.func读取所有最终结论模板(sql)
        cbox结论模板.Clear
        'lobj结论.MoveFirst
        For i = 1 To lobj结论.RecordCount
            cbox结论模板.AddItem lobj结论("结论模板")
            lobj结论.MoveNext
        Next i
    End If
'''    If cchk标准(0).Value = 1 And cchk标准(1).Value = 1 Then
'''        sql = "select * from 系统管理_字典_体检结论模板表 and 科室编号=(select 编号 from 系统管理_字典_字典内容表 where id='84' and 名称='最终结论录入')"
'''        Set lobj结论 = pub结论.func读取所有最终结论模板(sql)
'''        cbox结论模板.Clear
'''        'lobj结论.MoveFirst
'''        For i = 1 To lobj结论.RecordCount
'''            cbox结论模板.AddItem lobj结论("结论模板")
'''            lobj结论.MoveNext
'''        Next i
'''    End If
'''    If cchk标准(0).Value = 0 And cchk标准(1).Value = 0 Then
'''        cbox结论模板.Clear
'''    End If
    If cchk标准(0).Value = True Then
        ctxtReview.Enabled = False
        ctxtReviewItem.Enabled = False
        cmd选择复查项目.Enabled = False
    Else
        ctxtReview.Enabled = True
        ctxtReviewItem.Enabled = True
        cmd选择复查项目.Enabled = True
    End If
    '2012-07-03 于登淼 ↑
End Sub

'2012-04-11 于登淼
'体检类别列表添加所有项，并找出相应体检表模板
Private Sub Ccmb体检人类别_Click()
    Dim lobj体检表模板集 As Object
    Dim lobj体检类别 As Object
    Dim lcolInfo As New Collection
    Dim lcol体检表编号 As Collection
    Dim i As Integer
    On Error GoTo errHandler
    
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
    sfsub错误处理 "职业病界面", "frmFinalConclusion", "ccmb体检人类别_click", Err.Number, Err.Description, True
End Sub

'2012-04-11 于登淼
'体检人员类别列表添加所有项，同时添加体检类别列表所有项
Private Sub ccmb体检人类型_Click()
    Dim lobj体检类型 As Object
    On Error GoTo errHandler
    Set lobj体检类型 = CreateObject("职业病对象.clsmedicalexam")
    lobj体检类型.体检类型 = ccmb体检人类型.ItemData(ccmb体检人类型.ListIndex)
    Call Ccmb体检人类别_Click
    Exit Sub
errHandler:
    sfsub错误处理 "职业病界面", "frmFinalConclusion", "Private Sub ccmb体检人类型_Click", Err.Number, Err.Description, True
End Sub

'2012-04-10 于登淼
'单位定位，为了方便查询某个单位人员的体检结果信息
Private Sub ccmdLocate_Click()
    On Error GoTo errHandler
    Dim lobjRec As Object                       '单位定位返回的结果记录。
    Set lobjRec = pobj业务对象.func单位定位     '启动单位定位界面。
    
    '获取定位的单位，显示在“单位名称”录入框中。(暂时只显示“单位名称”)
    '-----不知道这里需不需要在其他模块里面设定涉核部队。
    If Not lobjRec Is Nothing Then
        If lobjRec.RecordCount > 0 Then
            ctxtCompanyName.Text = IIf(IsNull(lobjRec("单位名称")), "", lobjRec("单位名称"))
        End If
    End If
    Set lobjRec = Nothing
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病界面", "frmFinalConclusion", "ccmdLocate_Click", 6666, lstrError, False
End Sub

''''2012-04-11 于登淼
''''将查询结果直接加到cgrdInfo列表的后面。
'''Private Sub ccmdAdd_Click()
'''    On Error GoTo errHandler
'''
'''
'''    Dim lobjTmp, lobjRec As Object
'''    Dim i As Integer, j As Integer
'''    Dim lstrWhere As String
'''
'''    lstrWhere = " and 体检表编号='" & ccmbTemplate.Text & "'"      '
'''    lstrWhere = " and 体检类型='" & ccmb体检人类型 & "'"
'''
'''    '组装查询条件
'''    If cchkDate.Value = 1 Then          '体检日期
'''        lstrWhere = lstrWhere & " and 体检日期='" & Format(cdtpDateTo.Value, "yyyy-mm-dd hh:mm:ss") & "'"
'''    End If
'''
'''    If cchkBarCode.Value = 1 Then                         '体检条码号
'''        lstrWhere = lstrWhere & " and 系统编号='" & Trim(ctxtBarCode) & "'"
'''    End If
'''
'''    If cchkCompanyName.Value = 1 Then                         '单位名称
'''        lstrWhere = lstrWhere & " and 单位名称='" & Trim(ctxtCompanyName) & "'"
'''    End If
'''
'''    If coptConclusion(0).Value = True Then                          '已下结论、未下结论(总结论)
'''        lstrWhere = lstrWhere & " and ((体检结论 is null) or 体检结论='')"
'''    Else
'''        lstrWhere = lstrWhere & " and (体检结论 is not null)"
'''    End If
'''
'''    If reSql = "0" Or reSql <> lstrWhere Then
'''        reSql = lstrWhere
'''        Set lobjTmp = CreateObject("职业病体检结果录入.clscommon")
'''        Set lobjRec = lobjTmp.func获取可修改结论的_特定科室的_体检人员基本信息(lstrWhere, "")
'''         '在现有基础上添加这些行，控制冗余信息显示
'''        If lobjRec.RecordCount > 0 Then
'''            lobjRec.MoveFirst
'''            With cgrdInfo
'''                For i = 1 To lobjRec.RecordCount
'''                    .AddItem ("")
'''                    For j = 0 To .Cols - 1
'''                        If .TextMatrix(0, j) = "体检条码号" Then
'''                            .TextMatrix(.Rows - 1, j) = lobjRec("系统编号")
'''                        Else
'''                            .TextMatrix(.Rows - 1, j) = lobjRec(.TextMatrix(0, j))
'''                        End If
'''                    Next
'''                    lobjRec.MoveNext
'''                Next
'''                .AutoSize 0, .Cols - 1, 0, 0
'''            End With
'''        End If
'''    ElseIf reSql = lstrWhere Then
'''        Exit Sub
'''    End If
'''
'''    Set lobjTmp = Nothing
'''    Set lobjRec = Nothing
'''    lstrWhere = ""
'''    Exit Sub
'''errHandler:
'''    Dim lstrError As String
'''    lstrError = func错误处理(Err.Number, Err.Description)
'''    sfsub错误处理 "职业病界面", "frmFinalConclusion", "ccmdQuery_Click", 6666, lstrError, False
'''End Sub

'2012-07-03 于登淼
'将查询结果放入cgrdinfo列表中，覆盖之前的结果。
Private Sub ccmdQuery_Click()
    Dim lstr体检表名称, lstr单位名称, lstr系统编号 As String
    Dim lstr开始日期, lstr结束日期 As Date
    Dim lstr体检表类型 As String           '2015-11-9 牟俊 增加体检表类型
    Dim i As Integer
    
    If cchkDate.Value = 1 Then
        '修改人：张令 2012.12.05
        'bug号：0000071
        '说明：开始日期与结束日期一样，与查询的日期不符。改为一天的0点到23点。  ↓↓
'        lstr开始日期 = cdtpDateTo.Value: lstr结束日期 = cdtpDateTo.Value
        lstr开始日期 = CStr(Format(cdtpDateTo.Value, "yyyy/mm/dd"))
        lstr结束日期 = CStr(Format(cdtpDateFrom.Value, "yyyy/mm/dd"))
        '2012.12.05    ↑↑
    Else
        lstr开始日期 = "1900-01-01 00:00:00": lstr结束日期 = "3000-01-01 00:00:00"
    End If
    
    If cchkTemplate.Value = 1 Then lstr体检表名称 = ccmbTemplate.Text
    If cchkBarCode.Value = 1 Then lstr系统编号 = ctxtBarCode.Text
    If cchkCompanyName.Value = 1 Then lstr单位名称 = ctxtCompanyName.Text
    
    Set mobjQueryResult = pobj业务对象.func体检管理界面查询(lstr开始日期, lstr结束日期, lstr体检表名称, lstr单位名称, "", "", "", lstr系统编号, "")
'    Set mobjQueryResult = pobj业务对象.func体检管理界面查询(lstr开始日期, lstr结束日期, lstr体检表名称, lstr单位名称, "", "", "", lstr系统编号, "", lstr体检表类型)
'    For i = 0 To 5
'        If coptType(i).Value = True Then mobjQueryResult.Filter = "体检状态='" & coptType(i).Caption & "'": Exit For
    '修改人：罗李奎 2012-12-12 ↓
'        If coptType(i).Index = 0 Then
        If coptType(0).Value = True Then
            mobjQueryResult.Filter = "体检状态='" & coptType(i).Caption & "'  or 体检状态='未录入受检者个人信息'"
'        Else
'            If coptType(i).Value = True Then mobjQueryResult.Filter = "体检状态='" & coptType(i).Caption & "'": Exit For
        ElseIf coptType(1).Value = True Then
            mobjQueryResult.Filter = "体检状态='未下结论'"
        ElseIf coptType(2).Value = True Then
            mobjQueryResult.Filter = "体检状态='待复核'"
        ElseIf coptType(3).Value = True Then
            mobjQueryResult.Filter = "体检状态='已复核' or 体检状态='已发报告' or 体检状态='待复查'"
        ElseIf coptType(6).Value = True Then     '2016-4-20 by 牟俊
            mobjQueryResult.Filter = "体检状态='未下建议'"
        End If
        '修改人：罗李奎 2012-12-12 ↑
'    Next
    Set cgrdInfo.DataSource = mobjQueryResult
    Set mcolIndex = New Collection
    For i = 0 To cgrdInfo.cols - 1
        mcolIndex.Add i, cgrdInfo.TextMatrix(0, i)
    Next
    cgrdInfo.ColHidden(mcolIndex("试管编号")) = True
    cgrdInfo.ColHidden(mcolIndex("体检结论")) = True
    
    'vsflexgrid列宽度按内容自动调整；点击表头按表头下内容排序
    cgrdInfo.AutoSize 0, cgrdInfo.cols - 1, 0, 0
    cgrdInfo.ExplorerBar = flexExSort
    cgrdInfo.DataMode = flexDMFree
    cgrdInfo.Col = 0
    cgrdInfo.Sort = flexSortGenericDescending
    
    ctxtConclusion.Text = ""  '清空结论和建议,隐藏保存意见按钮和结果描述框,退回复核按钮等  2016-4-20 by 牟俊
    ctxtDiagnose.Text = ""
    comd保存意见.Visible = False
    Text结果描述.Visible = False
    Com退回.Visible = False
    Label10.Caption = cgrdInfo.rows - 1
End Sub

'2012-04-12 于登淼
'单击查看单个科室完整结论
Private Sub cgrdDept_Click()
    Dim strTmp As String
    If cgrdDept.Row = 0 Then Exit Sub
    strTmp = cgrdDept.TextMatrix(cgrdDept.Row, 0)
    strTmp = Right(strTmp, Len(strTmp) - 3)
    sub列出单科室所有体检结果 (strTmp)     '固定第0列为科室名称
    '修改人：张令 2012.12.12    ↓↓
    '当按钮勾选时执行单击事件。
    If cchkUnfilled.Value = 1 Then
        cchkUnfilled_Click
    ElseIf cchkAbnormal.Value = 1 Then
        cchkAbnormal_Click
    End If
    '修改人：张令 2012.12.12    ↑↑
    
    '增加 各科结果描述   2015-11-5
    Dim Jlstr As String
    Dim JstrTmp As String
    Dim Jsql As Object
    JstrTmp = cgrdDept.TextMatrix(cgrdDept.Row, 0)
    JstrTmp = Left(JstrTmp, 2)
    Jlstr = "select 文字结论 from 职业病体检_科室结论表 where 科室='" & JstrTmp & " ' and 系统编号='" & mstr系统编号 & "' "
    Set Jsql = dafuncGetData(Jlstr)
    Text结果描述.Text = Jsql("文字结论")
    Text结果描述.Visible = True
'    Text结果描述.Top = 3000
    Text结果描述.Top = ctxtDetpConclusion.Top
    Text结果描述.Left = ctxtDetpConclusion.Left + ctxtDetpConclusion.Width
    Text结果描述.Width = cgrdDept.Width - ctxtDetpConclusion.Width
    Label10.Caption = cgrdInfo.rows
End Sub

'双击选中，宋科长要求，从此处复制意见
Private Sub cgrdDept_DblClick()
    cgrdDept.EditCell
    Text结果描述.Visible = False
End Sub

'修改人：张令 2012.12.12           ↓↓
'功能与单击事件一样，重复了。

'2012-04-12 于登淼
'双击查看每个科室的单项结果
'Private Sub cgrdDept_DblClick()
'    Dim strTmp As String
'    If cgrdDept.Row = 0 Then Exit Sub
'    strTmp = cgrdDept.TextMatrix(cgrdDept.Row, 0)
'    strTmp = Right(strTmp, Len(strTmp) - 3)
'    sub列出单科室所有体检结果 (strTmp)     '固定第0列为科室名称
'    cchkAbnormal_Click
'End Sub


'2012-04-12 于登淼
'双击查看科室结论
Private Sub cgrdInfo_DblClick()
        cgrdDept.Clear
        cgrdItem.Clear
        pstrPerson = cgrdInfo.TextMatrix(cgrdInfo.Row, 0)
        sub列出各科室结论
        '功能：界面初始化的时候隐藏
        '作者：翁乔
        '时间：2012-06-01
'        cbox结论模板.Visible = True
        cchk标准(1).Visible = True
        cchk标准(0).Visible = True
        '时间：2012-06-01
        
'''        '2012-06-21 于登淼 ↓
'''        '查询当前体检人员的体检状态，判断是否允许复核，更改当前工具栏操作状态
'''        Dim lobjRec As Object
'''        Set lobjRec = pobj业务对象.func体检管理界面查询("", "", "", "", "", "", "", pstrPerson)
'''        mstrState = lobjRec("体检状态")
'''        ctlb工具栏.Buttons(5).Enabled = (mstrState = "已下结论")
'''        ctlb工具栏.Buttons(7).Enabled = True   '预览报告
'''        Set lobjRec = Nothing
'''        '2012-06-21 于登淼 ↑
        
        '2012-06-25 于登淼 ↓
        '体检不合格项目自动填入体检总结果中
         If coptType(1).Value = True Then
            sub自动填入不合格项目 pstrPerson
        End If
       '2012-06-25 于登淼 ↑
       
        '双击时，结论模版默认有“建议：”几字       2015-10-19
        If ctxtDiagnose = "" Then
       ctxtDiagnose.Text = "建议："
        End If
        
       '选择人后“个人信息”按钮才能用   2015-10-22
         Command2.Enabled = True
         Command2.Visible = True
         
    '双击时8023处理意见不要（要让双击和单击效果一样）      2015-10-26   牟俊
    Dim resql As Object
    mstr系统编号 = cgrdInfo.TextMatrix(cgrdInfo.Row, mcolIndex("系统编号"))
    Set resql = dafuncGetData("select 体检表类型 From 职业病体检_体检人员基本信息表 where 系统编号='" & mstr系统编号 & "'")
'    If resql("体检表类型") = "8023部队" Or resql("体检表类型") = "涉核部队" Then
    If resql("体检表类型") = "8023部队" Or resql("体检表类型") = "涉核部队" Or resql("体检表类型") = "涉核部队YK" Then
        If mobjQueryResult.Filter = "体检状态='未下结论'" Then
        ctxtDiagnose.Enabled = False
        Cmd结论模版(1).Enabled = False
        Else
        ctxtDiagnose.Enabled = True
        Cmd结论模版(1).Enabled = True
        End If
    Else
        ctxtDiagnose.Enabled = True
        Cmd结论模版(1).Enabled = True

    End If
    '（双击和单击效果一样）当8023是“待复核”状态时，“保存结论”按钮有用（8023结论是在复核时下的）  2015-10-26 牟俊
    
    If mobjQueryResult.Filter = "体检状态='待复核'" Then
        If resql("体检表类型") = "8023部队" Or resql("体检表类型") = "涉核部队" Or resql("体检表类型") = "涉核部YK" Then
'        If ctlb工具栏.Buttons(5).Visible = True And (resql("体检表类型") = "8023部队" Or resql("体检表类型") = "涉核部队" Or resql("体检表类型") = "涉核部队YK") Then
'        If resql("体检表类型") = "8023部队" Then
        ctxtDiagnose.Enabled = True
        Cmd结论模版(1).Enabled = True
        ctlb工具栏.Buttons(3).Enabled = False
        Else
        ctxtDiagnose.Enabled = True
'        ctxtDiagnose.Enabled = False
        ctlb工具栏.Buttons(3).Enabled = False
        End If
    End If
     Text结果描述.Visible = False  '结果描述框隐藏
      '只有用户是宋科并且状态要是未下建议才显示保存建议按钮  2016-4-20 by 牟俊
     If (um用户编号 = "8827" Or um用户编号 = "0001") And mobjQueryResult.Filter = "体检状态='未下建议'" Then
     comd保存意见.Visible = True
     ctlb工具栏.Buttons(13).Enabled = True
     End If
    '给宋科一个退回复核的按钮，因为有时错误操作导致到已复核，需要重新退回待复核  2016-5-19 by 牟俊
    If (um用户编号 = "8827" Or um用户编号 = "0001") And coptType(3).Value = True Then
    Com退回.Visible = True
    End If
End Sub

'2012-06-21 于登淼
'单击查看科室结论，与双击的操作完全相同
'（双击容易误操作，但单击复杂度有些高）
Private Sub cgrdInfo_Click()
    indX = cgrdInfo.MouseRow
    indY = cgrdInfo.MouseCol
    If indX < 0 Or indY < 0 Then
        Exit Sub
    ElseIf indX >= 0 And indX < cgrdInfo.rows And indY >= 0 And indY < cgrdInfo.cols Then
        cgrdDept.Clear
        cgrdItem.Clear
        Text结果描述.Text = ""  '每次点击人员记录时清空原来的结果描述  2015-11-5
        pstrPerson = cgrdInfo.TextMatrix(indX, 0)
        sub列出各科室结论
        '功能：界面初始化的时候隐藏
        '作者：翁乔
        '时间：2012-06-01
''        cbox结论模板.Visible = True
        cchk标准(1).Visible = True
        cchk标准(0).Visible = True
        '时间：2012-06-01
        
''自动导入不合格项目  2016-5-13 by 牟俊
'        If coptType(1).Value = True Then
'            sub自动填入不合格项目 pstrPerson
'        End If

        
        '2012-08-22 于登淼 ↓
        '更改设定，ctxtDetpConclusion.text用来提示未体检完的科室和未填结果项。
        'ctxtDetpConclusion.Text = cgrdDept.TextMatrix(cgrdDept.Row, 1)      '固定第1列为科室结论
        ctxtDetpConclusion.Text = sub未体检完科室与体检项目(pstrPerson)
        '2012-08-22 于登淼 ↑
    
    End If
    '单击时，结论模版默认有“建议：”几字       2015-10-19
    If ctxtDiagnose = "" Then
       ctxtDiagnose.Text = "建议："
    End If
    
    '选择人后“个人信息”按钮才能用   2015-10-22
    Command2.Enabled = True
    Command2.Visible = True
    
    
    '单击时8023处理意见不要（要让单击和双击效果一样）      2015-10-26
    Dim resql As Object
    mstr系统编号 = cgrdInfo.TextMatrix(cgrdInfo.Row, mcolIndex("系统编号"))
    Set resql = dafuncGetData("select 体检表类型 From 职业病体检_体检人员基本信息表 where 系统编号='" & mstr系统编号 & "'")
'    If resql("体检表类型") = "8023部队"  Then
    If resql("体检表类型") = "8023部队" Or resql("体检表类型") = "涉核部队" Or resql("体检表类型") = "涉核部队YK" Then
        If mobjQueryResult.Filter = "体检状态='未下结论'" Then
        ctxtDiagnose.Enabled = False
        Cmd结论模版(1).Enabled = False
        Else
        ctxtDiagnose.Enabled = True
        Cmd结论模版(1).Enabled = True
        End If
    Else
        ctxtDiagnose.Enabled = True
        Cmd结论模版(1).Enabled = True

    End If
    '（双击和单击效果一样）当8023是“待复核”状态时，“保存结论”按钮有用（8023结论是在复核时下的）  2015-10-26
    
    If mobjQueryResult.Filter = "体检状态='待复核'" Then
        If resql("体检表类型") = "8023部队" Or resql("体检表类型") = "涉核部队" Or resql("体检表类型") = "涉核部YK" Then
'        If ctlb工具栏.Buttons(5).Visible = True And (resql("体检表类型") = "8023部队" Or resql("体检表类型") = "涉核部队" Or resql("体检表类型") = "涉核部YK") Then
'        If resql("体检表类型") = "8023部队" Then
        ctxtDiagnose.Enabled = True
        Cmd结论模版(1).Enabled = True
        ctlb工具栏.Buttons(3).Enabled = False
        Else
        ctxtDiagnose.Enabled = True
'        ctxtDiagnose.Enabled = False
        ctlb工具栏.Buttons(3).Enabled = False
        End If
     End If
     Text结果描述.Visible = False  '结果描述框隐藏
     '只有用户是宋科(编号为8827)并且状态要是未下建议才显示保存建议按钮  2016-4-20 by 牟俊
     If (um用户编号 = "8827" Or um用户编号 = "0001") And mobjQueryResult.Filter = "体检状态='未下建议'" Then
     comd保存意见.Visible = True
     ctlb工具栏.Buttons(13).Enabled = True    '预览
     End If
    '给宋科一个退回复核的按钮，因为有时错误操作导致到已复核，需要重新退回待复核  2016-5-19 by 牟俊
    If (um用户编号 = "8827" Or um用户编号 = "0001") And coptType(3).Value = True Then
    Com退回.Visible = True
    End If
End Sub

'2012-04-12 于登淼
'按下BackSpace键，ascii码为8，就将选中列从列表中删除。
'删除操作，从最后一个选中行往前删。这个是必须的！
'因为每次删一行，所有行的编号会重新分配
Private Sub cgrdInfo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        Dim i, selRow As Integer
        For i = cgrdInfo.SelectedRows - 1 To 0 Step -1
            selRow = cgrdInfo.SelectedRow(i)
            cgrdInfo.RemoveItem (selRow)
        Next
        '修改人：张令 2012.12.05
        'bug号：0000070
        '说明：移除人员后清空其他表格和文本框数据   ↓↓
        cgrdDept.rows = 1
        ctxtConclusion.Text = ""
        '2012.12.05  ↑↑
    End If
End Sub

Private Sub cgrdItem_DblClick()
If Left(cgrdItem.TextMatrix(cgrdItem.RowSel, 0), 2) = "08" Then
formdct.Label2.Caption = cgrdInfo.TextMatrix(cgrdInfo.RowSel, 0)
formdct.Show 1
End If


End Sub

'套用已有的结论模板 可进行选择
Private Sub Cmd结论模版_Click(Index As Integer)
    If Index = 0 Then
         frmConclusion.lobj科室 = "结论模版"
    Else
         frmConclusion.lobj科室 = "意见模版"
     End If
    frmConclusion.lobj调用科室 = Me.name
'    frmConclusion.lobj科室 = priDeptName
'    frmConclusion.lobj科室编号 = priDeptNo
    frmConclusion.lobj医生编号 = um用户编号
    frmConclusion.lobj时间 = Now
    frmConclusion.Show
End Sub


Private Sub cmd选择复查项目_Click()
    frmSelectItem.Show 1
End Sub

'保存建议和结论   2016-4-20 by 牟俊
Private Sub comd保存意见_Click()
        Dim teststyle As String
        Dim i, selRow As Integer
        selRow = cgrdInfo.SelectedRow(i)
        teststyle = cgrdInfo.TextMatrix(selRow, 8)
        Dim sjjl As String
        Dim SyNo As String
        sjjl = ctxtConclusion.Text & "_00_" & Trim(ctxtDiagnose.Text)
        SyNo = cgrdInfo.TextMatrix(selRow, 0)
        Dim lob As Object
        Set lob = dafuncGetData("select 姓名 from 系统管理_员工基本信息表 where 编号='" & um用户编号 & "'")
        dafuncGetData "update 职业病体检_科室结论表 set 文字结论='" & sjjl & "',医生编号='" & lob("姓名") & "',结论日期=getdate() where 系统编号='" & SyNo & "' and 科室=(select 编号 from 系统管理_字典_字典内容表 where 名称='最终结论录入' and 描述='职业病体检_科室');"
        dafuncGetData ("update 职业病体检_体检基本信息表 set 体检结论='" & ctxtConclusion.Text & "', 诊断和处理意见='" & Trim(ctxtDiagnose.Text) & "',体检状态='5'  where 系统编号='" & SyNo & "'")
        ccmdQuery_Click
        cgrdDept.rows = 1
        comd保存意见.Visible = False
End Sub

Private Sub Command1_Click()
If cgrdItem.TextMatrix(cgrdItem.RowSel, 0) <> "" Then
       '去掉字符串左边的数字， 2015-10-19 牟俊
         Dim stmp As String
         stmp = cgrdDept.TextMatrix(cgrdDept.RowSel, 0)
         Dim slen As Integer
         slen = Len(stmp)
         Dim clen As Integer
         clen = slen - 2
         Dim ctemp As String
         ctemp = Right(stmp, clen)
    If ctxtConclusion.Text = "" Then
'        ctxtConclusion.Text = ctemp + "-" + cgrdItem.TextMatrix(cgrdItem.RowSel, 1) + "-" + cgrdItem.TextMatrix(cgrdItem.RowSel, 2) + "；"
        ctxtConclusion.Text = ctemp + "-" + cgrdItem.TextMatrix(cgrdItem.RowSel, 2) + "；" '只要结果，不要前面的项目 2016-4-18 by 牟俊
    Else
'        ctxtConclusion.Text = ctxtConclusion.Text + ctemp + "-" + cgrdItem.TextMatrix(cgrdItem.RowSel, 1) + "-" + cgrdItem.TextMatrix(cgrdItem.RowSel, 2) + "；"
        ctxtConclusion.Text = ctxtConclusion.Text + cgrdItem.TextMatrix(cgrdItem.RowSel, 2) + "；" '只要结果，不要前面的项目  2016-4-18 by 牟俊
    End If
Else
End If
End Sub


'增加一个能查看完整个人信息的按钮 2015-10-22↓
Private Sub Command2_Click()
If cgrdInfo.Row < 1 Then
            MsgBox "没有需要修改的记录！", vbOKOnly + vbExclamation, "系统提示"
            Exit Sub
        End If

'        访问记号 = 1
        mstr系统编号 = cgrdInfo.TextMatrix(cgrdInfo.Row, mcolIndex("系统编号"))
        'frmCareerHstRegt.ctxtsysno = cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("系统编号"))
        frmInfromation.Show 1, Me
End Sub


'增加评残按钮 2016-5-30 by 牟俊
Private Sub Com评残_Click()
frmAssessDeformity.Show 1
End Sub

Private Sub Com退回_Click()
If MsgBox("你确认要退回该条记录到待复核吗？", vbYesNo + vbQuestion + vbDefaultButton2, "系统提示") = vbYes Then
            pobj业务对象.func写入单人当前体检状态 pstrPerson, 5 '"已下结论"
            ccmdQuery_Click
            cgrdDept.rows = 1
            MsgBox "已成功退回待复核，请查看。"
End If
End Sub

'2012-07-03 于登淼
'cgrdinfo内容随体检状态改变而改变，同时，界面权限也发生变化。
Private Sub coptType_Click(Index As Integer)
    cgrdDept.rows = 1
    ctxtConclusion.Text = ""
    ctxtReview.Text = ""
    ctxtReviewItem.Text = ""
    '2013-1-17 刘云乐
    '直接触发查询事件
'    If Not mobjQueryResult Is Nothing Then
'        If coptType(Index).Caption <> "体检中" Then
'            mobjQueryResult.Filter = "体检状态='" & coptType(Index).Caption & "'"
'
'        Else
'            mobjQueryResult.Filter = "体检状态='体检中' or 体检状态='未录入受检者个人信息'"
'        End If
'    End If
'    Set cgrdInfo.DataSource = mobjQueryResult
    ccmdQuery_Click
    '2013-1-17 刘云乐
    If cgrdInfo.rows > 1 Then
        cgrdInfo.AutoSize 0, cgrdInfo.cols - 1, 0, 0
        cgrdInfo.ExplorerBar = flexExSort
        cgrdInfo.DataMode = flexDMFree
        cgrdInfo.Col = 0
        cgrdInfo.Sort = flexSortGenericDescending
    End If
        
    ctlb工具栏.Buttons(6).Visible = False
    ctlb工具栏.Buttons(7).Visible = False
    ctlb工具栏.Buttons(8).Visible = False
    ctlb工具栏.Buttons(9).Visible = False
    ctlb工具栏.Buttons(10).Visible = False
    ctlb工具栏.Buttons(3).Enabled = (coptType(Index).Caption = "未下结论") Or (coptType(Index).Caption = "已下结论") Or (coptType(Index).Caption = "待复查")
    ctlb工具栏.Buttons(5).Enabled = (coptType(Index).Caption = "已下结论") Or (coptType(Index).Caption = "待复查") Or (coptType(Index).Caption = "待复核")
    ctlb工具栏.Buttons(7).Enabled = (coptType(Index).Caption = "已下结论") Or (coptType(Index).Caption = "已复核") Or (coptType(Index).Caption = "已发报告") Or (coptType(Index).Caption = "待复查")
    ctlb工具栏.Buttons(9).Enabled = (coptType(Index).Caption = "已复核") Or (coptType(Index).Caption = "已发报告") Or (coptType(Index).Caption = "待复查")
    ctlb工具栏.Buttons(11).Enabled = (coptType(Index).Caption = "已复核") Or (coptType(Index).Caption = "已发报告") Or (coptType(Index).Caption = "待复查")
     
    '待复查状态可以打印体检报告；翁乔；2012-10-24
    '所有状态都能预览报告。 修改人：张令 2013.03.01
'    ctlb工具栏.Buttons(13).Enabled = (coptType(Index).Caption = "已复核") Or (coptType(Index).Caption = "已发报告") Or (coptType(Index).Caption = "待复查")

    '让未下结论，待复核，已复核 都可以预览  牟俊  2015-11-9
    ctlb工具栏.Buttons(13).Enabled = (coptType(Index).Caption = "已复核") Or (coptType(Index).Caption = "已发报告") Or (coptType(Index).Caption = "待复查") Or (coptType(Index).Caption = "未下结论") Or (coptType(Index).Caption = "待复核") Or (coptType(Index).Caption = "未下建议")
    '修改人：张令 2012.12.05
    'bug号：0000062
    '说明：当保存结论为不可用时，mstr权限标志赋false   ↓↓
    If ctlb工具栏.Buttons(3).Enabled = True Then
        mstr权限标志 = True
    Else
        mstr权限标志 = False
    End If
    '2012.12.05    ↑↑
    If coptType(2).Value = True Then
        ctlb工具栏.Buttons(15).Visible = True
        ctlb工具栏.Buttons(16).Visible = True
    Else
        ctlb工具栏.Buttons(15).Visible = False
        ctlb工具栏.Buttons(16).Visible = False
    End If
    '未下建议(6)，待复核(2)，已复核(3)三个中 取消结论 按钮可以用
    If coptType(2).Value = True Or coptType(6).Value = True Then
'    If coptType(2).Value = True Or coptType(6).Value = True Or coptType(3).Value = True Then
        ctlb工具栏.Buttons(15).Visible = True
    End If
'未下结论也可以预览  2016-7-25 by 牟俊
    If coptType(0).Value = True Then
        ctlb工具栏.Buttons(13).Enabled = True
    End If
    cchk标准(0).Value = True
End Sub





'2012-04-10 于登淼
Private Sub Form_Load()
    Dim lcol工具栏按钮 As New Collection '工具栏上的按钮初始化集合。
    
    On Error GoTo errHandler
    
    '如果窗体已经初始化过则不再进行初始化。
    If mblnInUse Then Exit Sub
    mstr权限标志 = True     '默认有权限
    '显示进度。
    frmProcess.proPercent.max = 8
    frmProcess.Label1.Caption = "正在初始化界面，请等待..."
    frmProcess.proPercent.Value = 1
    frmProcess.Show
    DoEvents
    Me.Enabled = False
    MousePointer = 11
    '设置窗体正在使用的标志。
    mblnInUse = True
    
    '初始化界面通用对象（每个界面对应一个界面通用对象实例，不可混用，切记切记）。
    Set mobjGUI = New cls界面通用对象
    
    '设置工具栏上所需要的各种按钮。
    With lcol工具栏按钮
        .Add "清空界面(&C)110"    '1
        .Add "移除人员(&M)104"    '2
        .Add "保存结论(&S)101"    '3
        .Add "|"    '4
        .Add "复核(&F)109"    '5
        .Add "|"    '6
        .Add "预览报告(&Y)108"    '7
        .Add "|"    '8
        .Add "打印报告(&P)107"    '9
        .Add "|"    '10
        .Add "导出为PDF(&O)102"    '11
        .Add "|"    '12
       
      '2015/1/16 这里的预览会出错，预览在报告管理里面
       .Add "预览(&V)102"    '13
        .Add "|"    '14
        .Add "取消结论(&E)111"    '15
        .Add "|"    '16
        .Add "退出"
    End With
    frmProcess.proPercent.Value = 2
    DoEvents
    '为需要通过界面通用对象控制的各种控件赋初始值。
    With mobjGUI
        Set .Form = Me
        Set .c工具栏 = ctlb工具栏
    End With
    
    resql = "0"
    
    '调用界面通用对象提供的方法，对界面控件进行初始化。
    mobjGUI.subInitialize lcol工具栏按钮, ""
    frmProcess.proPercent.Value = 3
    DoEvents
    '变量初始化
    Set pobj体检 = CreateObject("职业病对象.clsMedicalExam")
    Set pobj体检表模板 = CreateObject("职业病对象.clsMedicalExamTemplate")
    Set pobj体检结果业务 = CreateObject("职业病体检结果录入.clsCommon")
    Set pobj科室 = pobjDict.Fetch("职业病体检科室字典")
    pstrPerson = ""
    frmProcess.proPercent.Value = 4
    DoEvents
    '界面权限控制
    '控制的权限仅限于，复核、保存结论、打印报告、导出为PDF。
    Dim lobjTmp As Object
    Set lobjTmp = CreateObject("职业病设置.clsPermissionConfigure")
    If lobjTmp.func科室操作权限(um用户编号, "职业病体检_最终结论录入_保存结论") = False Then
        ctlb工具栏.Buttons(3).Visible = False
    End If
    If lobjTmp.func科室操作权限(um用户编号, "职业病体检_最终结论录入_复核通过") = False Then
        ctlb工具栏.Buttons(5).Visible = False
        ctlb工具栏.Buttons(6).Visible = False
    End If
    If lobjTmp.func科室操作权限(um用户编号, "职业病体检_最终结论录入_打印报告") = False Then
        ctlb工具栏.Buttons(9).Visible = False
    End If
    
    'If lobjTmp.func科室操作权限(um用户编号, "职业病体检_最终结论录入_另存PDF") = False Then
        ctlb工具栏.Buttons(10).Visible = False '注意：另存PDF权限未写入可用操作信息表
        ctlb工具栏.Buttons(11).Visible = False
    'End If
    Set lobjTmp = Nothing
    frmProcess.proPercent.Value = 5
    DoEvents
    '界面控件初始化
'    llabDoctor = llabDoctor & um用户名
    fraFinal.Caption = fraFinal.Caption & "(医师姓名：" & um用户名 & ")"
    cdtpDateTo.Value = DateAdd("m", -1, Now())
    cdtpDateFrom.Value = Now
    cdtpConclusion.Value = Now
    coptType(1).Value = True
    With cgrdInfo
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "体检条码号"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "姓名"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "性别"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "年龄"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "单位名称"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "体检类型"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "体检日期"
        '.Cols = .Cols + 1: .TextMatrix(0, .Cols-1) = "体检表名称"
        .AutoSize 0, .cols - 1, 0, 0
        .SelectionMode = flexSelectionListBox
    End With
    
    If coptType(0).Value = True Or coptType(1).Value = True Then
        ctlb工具栏.Buttons(15).Visible = False
        ctlb工具栏.Buttons(16).Visible = False
    Else
        ctlb工具栏.Buttons(15).Visible = True
        ctlb工具栏.Buttons(16).Visible = True
    End If
    '2012-06-21 于登淼 ↓
    '添加工具栏初始化状态。主要用于流程控制。
    '初始不能复核，选中某个体检人员判断是否复核。复核后打印，导出为PDF。
    ctlb工具栏.Buttons(5).Enabled = False   '复核
    ctlb工具栏.Buttons(7).Enabled = False   '预览报告
    ctlb工具栏.Buttons(9).Enabled = False   '打印报告
    ctlb工具栏.Buttons(11).Enabled = False  '导出为PDF
    ctlb工具栏.Buttons(13).Visible = True  '放开屏蔽预览  2015-11-6 by lanchao
    '2012-06-21 于登淼 ↑
    frmProcess.proPercent.Value = 6
    DoEvents
    '体检表模板列表初始化
    sub加载体检表模板
    frmProcess.proPercent.Value = 7
    DoEvents
    '功能：加载结论模板
    '作者：翁乔
    '时间：2012-05-31
    Dim lobj结论模板 As Object
    Dim lobj结论表 As Object
    Dim i As Integer
    Dim sql As String
    
    sql = "select * from 系统管理_字典_体检结论模板表 where 科室编号='16'"
    
    Set lobj结论模板 = CreateObject("职业病对象.clsConclusionSet")
    Set lobj结论表 = lobj结论模板.func读取所有最终结论模板(sql)
    If lobj结论表.RecordCount > 0 Then
        For i = 1 To lobj结论表.RecordCount
            cbox结论模板.AddItem lobj结论表("结论模板")
            lobj结论表.MoveNext
        Next i
    End If
    frmProcess.proPercent.Value = 8
    DoEvents
    '时间：2012-05-31
    
    '2012-06-21 于登淼 ↓
    '初始化体检状态
    mstrState = ""
    '2012-06-21 于登淼 ↑
    
    '2012-07-03 于登淼 ↓
    '界面控件初始化
    cchk标准(0).Value = True
    '2012-07-03 于登淼 ↑
    
    '2012-08-22 于登淼 ↓
    '添加科室变量
    Set pobjDept = pobjDict.Fetch("职业病体检科室字典")
    '2012-08-22 于登淼 ↑
    Unload frmProcess
'    Exit Sub
errHandler:
    Me.Enabled = True
    MousePointer = 0
    If Err.Number <> 0 Then
        Dim lstrError As String
        lstrError = func错误处理(Err.Number, Err.Description)
        sfsub错误处理 "职业病界面", "frmFinalConclusion", "Form_Load", 6666, lstrError, False
        Exit Sub
        Resume
    End If
End Sub

'2012-04-10 于登淼
'退出窗体时，清空部分变量
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    mblnInUse = False
    Set mobjGUI = Nothing
    Set mobjQueryResult = Nothing
End Sub

'窗体自适应分辨率大小
'2012-10-19 刘云乐
Private Sub Form_Resize()
    On Error Resume Next
    Picture1.Width = Me.ScaleWidth - Picture1.Left
    Picture1.Height = Me.ScaleHeight - Picture1.Top
    Frame1.Width = Picture1.Width - Frame1.Left
    Frame1.Height = Picture1.Height - Frame1.Top
    ctlb工具栏.Width = Frame1.Width - ctlb工具栏.Left
    fraFinal.Left = Frame1.Width - fraFinal.Width - 80
    fraPerson.Width = Frame1.Width - fraPerson.Left - fraFinal.Width - 160
    cgrdInfo.Width = fraPerson.Width - cgrdInfo.Left * 2
    Label5.Width = cgrdInfo.Width
    
    fraDeptItem.Width = Frame1.Width - fraDeptItem.Left - 80
    cgrdDept.Width = fraDeptItem.Width * 2 / 3
'    ctxtDetpConclusion.Width = cgrdDept.Width
    ctxtDetpConclusion.Width = cgrdDept.Width - 3000   '将ctxtDetpConclusion文本框宽度调小  2015-11-5
    cgrdItem.Left = cgrdDept.Width + cgrdDept.Left + 80
    cgrdItem.Width = fraDeptItem.Width - cgrdItem.Left - 80
    cchkAbnormal.Left = cgrdItem.Left
    cchkUnfilled.Left = cchkAbnormal.Left + cchkAbnormal.Width + 200
    
    fraDeptItem.Height = Frame1.Height - fraDeptItem.Top - 40
    cgrdDept.Height = fraDeptItem.Height * 2 / 3
    ctxtDetpConclusion.Top = cgrdDept.Height + cgrdDept.Top + 40
    ctxtDetpConclusion.Height = IIf(fraDeptItem.Height - ctxtDetpConclusion.Top - 40 <= 0, 1, fraDeptItem.Height - ctxtDetpConclusion.Top - 40)
    cchkAbnormal.Top = fraDeptItem.Height - cchkAbnormal.Height - 40
    cchkUnfilled.Top = cchkAbnormal.Top
    cgrdItem.Height = cchkAbnormal.Top - cgrdItem.Top - 100
    
End Sub

'2012-04-12 于登淼
'界面工具栏按钮操作设定
Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = True
    
    Dim lcolID As New Collection
    Dim lobj体检类型 As Object
    Dim lstrStatus As String '当前体检状态
    Dim mcolIndex As New Collection
     
    lcolID.Add pstrPerson
    Set lobj体检类型 = CreateObject("职业病对象.clsMedicalExam")
    lobj体检类型.系统编号 = pstrPerson
    
    Select Case Operate
    Case "清空界面"
        subClear
    Case "移除人员"
        cgrdInfo_KeyPress (8)
        '修改人：张令 2012.12.05
        'bug号：0000070
        '说明：此处移除人员以后跳出，否则会查询一次，将会显示数据。    ↓↓
        Exit Sub
        '2012.12.05         ↑↑
    '2012-06-21 于登淼 ↓
    '添加复核通过功能
    Case "复核"
        pobj业务对象.func写入单人当前体检状态 pstrPerson, 6 '"已复核"
        ctxtConclusion.Text = ""
        ctxtDiagnose = ""    '2015-10-16
        cgrdDept.rows = 1
    '2012-06-21 于登淼 ↑
    Case "预览报告"
       
        pobj业务对象.Sub打印文书 "职业病体检_最终结论", lcolID, False, True, False
    Case "打印报告"
        pobj业务对象.Sub打印文书 "职业病体检_最终结论", lcolID, True, False, False
        '2012-07-03 于登淼 ↓
        '控制体检状态
        pobj业务对象.func写入单人当前体检状态 pstrPerson, 7 '"已发报告"
        '2012-07-03 于登淼 ↑
    '2012-05-30 陶露
    Case "导出为PDF"
        pobj业务对象.Sub打印文书 "职业病体检_最终结论", lcolID, True, False, True
        '2012-07-03 于登淼 ↓
        '控制体检状态
        pobj业务对象.func写入单人当前体检状态 pstrPerson, 7 '"已发报告"
        '2012-07-03 于登淼 ↑
    '2012-05-30
    Case "保存结论"
    
    '8023,涉核，铀矿等单独处理，即保存结论后不到待复核，而是到特有的未下建议  2016-4-20 by 牟俊↓
    Dim resql As Object
'    mstr系统编号 = cgrdInfo.TextMatrix(cgrdInfo.Row, mcolIndex("系统编号"))
    Set resql = dafuncGetData("select 体检表类型 From 职业病体检_体检人员基本信息表 where 系统编号='" & mstr系统编号 & "'")
'    If resql("体检表类型") = "8023部队"  Then
    If resql("体检表类型") = "8023部队" Or resql("体检表类型") = "涉核部队" Or resql("体检表类型") = "涉核部队YK" Then
            subSaveConclusion
        '2012-07-03 于登淼 ↓
        '控制体检状态
        If cchk标准(0).Value Then
            pobj业务对象.func写入单人当前体检状态 pstrPerson, 11 '命状态11为8023，涉核，铀矿特有的"未下建议"
        Else
            pobj业务对象.func写入单人当前体检状态 pstrPerson, 11
            pobj业务对象.func写入复查简单信息 pstrPerson
        End If
        cgrdDept.rows = 1
    Else
    '2016-4-20 ↑
    
        subSaveConclusion
        '2012-07-03 于登淼 ↓
        '控制体检状态
        If cchk标准(0).Value Then
            pobj业务对象.func写入单人当前体检状态 pstrPerson, 5 '"已下结论"
        Else
            pobj业务对象.func写入单人当前体检状态 pstrPerson, 5 '以前为8"待复查".复检与打体检表没关系
            pobj业务对象.func写入复查简单信息 pstrPerson
        End If
        cgrdDept.rows = 1
        '2012-07-03 于登淼 ↑
    End If
        
        
        
    '2012-08-20 于登淼 ↓
    '添加word模板功能
    Case "word报告"
        With cgrdInfo
            If .Row < 1 Or .Row > .rows - 1 Then Exit Sub
            mstr体检表名称 = .TextMatrix(indX, 7)
'            mstr体检表名称 = .TextMatrix(.Row, mcolIndex("体检表名"))
        End With
    
           '作者：罗李奎 时间2013-1-9 ↓
             '显示进度。
            frmProcess.proPercent.max = 4
            frmProcess.Label1.Caption = "正在加载，请等待..."
            frmProcess.proPercent.Value = 0
            frmProcess.Show 0, Me
            DoEvents
         '作者：罗李奎 时间2013-1-9 ↑
            
        '获取体检系统编号
        If cgrdInfo.SelectedRow(0) = -1 Then Exit Sub
        If coptType(1).Value = True Or coptType(2).Value = True Then
            sub编辑word文档 Me, pstrPerson, mstr体检表名称, True
        Else
            sub编辑word文档 Me, pstrPerson, mstr体检表名称, False
        End If
          Unload frmProcess
        
               
        '2012-08-23 于登淼 ↓
        '控制体检状态（word的处理并不好，需要把数据库功能添加之后，这步骤才完善）
        If pstrFilename = "" Then Exit Sub
        If coptType(3).Value = True Then pobj业务对象.func写入单人当前体检状态 pstrPerson, 7   '"已发报告"
        '2012-08-23 于登淼 ↑
    '2012-08-20 于登淼 ↑
    
    
    
    
    '暂时先把这个隐藏，会报错，预览在报告管理里面有。
    
    Case "预览"
     subPrint True  '是否预览
    
     
    
    
    
    Case "取消结论"
        If MsgBox("你确认要取消结论并退回该条记录吗？", vbYesNo + vbQuestion + vbDefaultButton2, "系统提示") = vbYes Then
            pobj业务对象.func写入单人当前体检状态 pstrPerson, 4 '"已下结论"
            ccmdQuery_Click
            cgrdDept.rows = 1
            MsgBox "已成功取消结论，该体检信息已退回。"
        End If
    Case "退出"
        '修改人：  2012.12.05
        'bug号：0000062
        '说明：添加函数     ↓↓
        Dim isSave As Integer
        '2012.12.05    ↑↑
        Set frmFinalConclusion = Nothing
        '修改人：  2012.12.05
        'bug号：0000062
        '说明：当权限标志为true时，退出提示是否保存     ↓↓
        If mstr权限标志 = True And cgrdInfo.SelectedRows > 0 Then
'            isSave = MsgBox("是否保存已修改结果？", vbYesNoCancel)
'            If isSave = vbCancel Then Exit Sub
             Unload Me
'            If isSave = vbNo Then
'                mobjGUI_BeforeOperate "清空界面", True
'                Exit Sub
'            End If
            If isSave = vbYes Then
                mobjGUI_BeforeOperate "保存结论", False
                mstr权限标志 = False
                Unload Me
            Else
                Unload Me
            End If
        Else
            Unload Me
        End If
        '2012.12.05    ↑↑
        Cancel = True
    End Select
    
    '2012-07-03 于登淼 ↓
    '由于每次操作都可能改变体检状态，所以，每次操作完后重新查询结果。
    '2012.12.11 张令
    '说明：用"<>"不能达到预期效果，改成"="。↓↓
'    If Operate <> "清空界面" Or Operate <> "移除人员" Or Operate <> "预览报告" Or Operate <> "退出" Then ccmdQuery_Click
    If Operate = "复核" Or Operate = "打印报告" Or Operate = "导出为PDF" Or Operate = "保存结论" Or Operate = "word报告" Then ccmdQuery_Click
    '2012.12.11  ↑↑
    '2012-07-03 于登淼 ↑
    Set lobj体检类型 = Nothing
    
    Exit Sub
errHandler:
    sfsub错误处理 "职业病界面", "frmFinalConclusion", "mobjGUI_BeforeOperate", Err.Number, Err.Description, False
    Cancel = True
    Exit Sub
    Resume
End Sub

'2012-04-12 于登淼
Sub sub加载体检表模板()
    Dim i As Integer
    Dim lobjRec As Object
    On Error GoTo errHandler

    '将体检类别加入组合框中
    Set lobjRec = pobjDict.FetchEx("体检类型字典")
    Ccmb体检人类别.Clear
    'Ccmb体检人类别.AddItem ""
    For i = 1 To lobjRec.RecordCount
        Ccmb体检人类别.AddItem lobjRec("名称")
        Ccmb体检人类别.ItemData(Ccmb体检人类别.NewIndex) = lobjRec("编号")
        lobjRec.MoveNext
    Next
    Ccmb体检人类别.ListIndex = 0
   
    Set lobjRec = pobjDict.FetchEx("体检人类别字典")
    ccmb体检人类型.Clear
    'ccmb体检人类型.AddItem ""
    For i = 1 To lobjRec.RecordCount
        ccmb体检人类型.AddItem lobjRec("名称")
        ccmb体检人类型.ItemData(ccmb体检人类型.NewIndex) = lobjRec("编号")
        lobjRec.MoveNext
    Next
    ccmb体检人类型.ListIndex = 0
   
    If ccmbTemplate.ListCount > 0 Then
        ccmbTemplate.Text = ccmbTemplate.List(0)
        subChangeTemplate
    Else
        ccmb体检人类型_Click
    End If
        
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病界面", "frmFinalConclusion", "sub加载体检表模板", 6666, lstrError, False
    Exit Sub
    Resume
End Sub

'2012-04-12 于登淼
'选择体检表模板下拉列表
Private Sub ccmbTemplate_Click()
    On Error GoTo errHandler
    MousePointer = 11
    subChangeTemplate       '选择体检表
    MousePointer = 0
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病界面", "frmFinalConclusion", "ccmbTemplate_Click", 6666, lstrError, False
    Exit Sub
    Resume
End Sub

'2012-04-12 于登淼
'改变体检表模板时的操作。
Private Sub subChangeTemplate()
    On Error GoTo errHandler
    
    If pobj体检.体检表.体检表名 <> ccmbTemplate.Text Then
        pobj体检.体检表.体检表名 = ccmbTemplate.Text
        '根据体检表模板获取该体检表所有可用的字母。
        pobj体检表模板.体检表名 = ccmbTemplate.Text
    End If
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病界面", "frmFinalConclusion", "subChangeTemplate", 6666, lstrError, True
    Exit Sub
    Resume
End Sub

'2012-04-12 于登淼
Private Sub sub列出各科室结论()
    On Error GoTo errHandler
    Dim lobjRec As Object
    Dim lstrCon As String
    Dim strArray
    Dim i As Integer
    
    '2012-05-24 于登淼 ↓
    '每次查询，清空当前已有结果表格、科室结论、最终结论、建议
    cgrdDept.Clear
    cgrdDept.rows = 1
    cgrdItem.Clear
    cgrdItem.rows = 1
    ctxtDetpConclusion.Text = ""
    ctxtConclusion.Text = ""
    ctxtDiagnose.Text = ""

    
    'vsflexgrid列宽度按内容自动调整；点击表头按表头下内容排序
    cgrdInfo.AutoSize 0, cgrdInfo.cols - 1, 0, 0
    cgrdInfo.ExplorerBar = flexExSort
    cgrdInfo.DataMode = flexDMFree
    
    '2012-05-24 于登淼 ↑
    
    Set lobjRec = pobj体检结果业务.func获取体检人员科室结论(pstrPerson)  '传回的科室为编号
    
    If lobjRec.RecordCount > 0 Then
        Set cgrdDept.DataSource = lobjRec
        With cgrdDept
            lobjRec.MoveFirst
            For i = 1 To lobjRec.RecordCount
                pobj科室.Filter = "编号=" & lobjRec("科室")
                .TextMatrix(i, 0) = lobjRec("科室") & " " & pobj科室("名称")       '固定第0列为科室名称，datasource决定
                If pobj科室("名称") = "最终结论录入" Then
                    .RowHidden(i) = True
                End If
                lobjRec.MoveNext
            Next
            .Col = 0
            .Sort = flexSortStringAscending
            .AutoSize 0, .cols - 1, 0, 0
            .SelectionMode = flexSelectionListBox
            .AllowSelection = False
        End With
    End If
    lstrCon = pobj业务对象.func返回科室结论(pstrPerson, "最终结论录入")
    
    ' "_00_" 没有什么特殊含义，只是作为分割结论和诊断意见的分隔符
    strArray = Split(lstrCon, "_00_", -1, vbBinaryCompare)
    If UBound(strArray) = 1 Then
        ctxtConclusion.Text = strArray(0)
        ctxtDiagnose.Text = strArray(1)
    End If
    Set lobjRec = dafuncGetData("select 复查原因,复查项目 from 职业病体检_体检基本信息表 where 系统编号='" & pstrPerson & "'")
    If Not (lobjRec.BOF Or lobjRec.EOF) Then
        ctxtReview.Text = IIf(IsNull(lobjRec("复查原因")), "", lobjRec("复查原因"))
        ctxtReviewItem.Text = IIf(IsNull(lobjRec("复查项目")), "", lobjRec("复查项目"))
        '注释：下面判断条件永远不成立，重新修改判断条件  2016-5-16 by 牟俊
'        If IsNull(lobjRec("复查原因")) And IsNull(lobjRec("复查项目")) Then
'            cchk标准(0).Value = True
'        Else
'            cchk标准(1).Value = True
'        End If
        If Me.cgrdInfo.TextMatrix(Me.cgrdInfo.Row, 12) = "" And Me.cgrdInfo.TextMatrix(Me.cgrdInfo.Row, 13) = "" Then
        cchk标准(0).Value = True
        Else
        cchk标准(1).Value = True
        End If
    End If
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病界面", "frmFinalConclusion", "sub列出各科室结论", 6666, lstrError, True
    Exit Sub
    Resume
End Sub

'2012-04-12 于登淼
Private Sub sub列出单科室所有体检结果(ByVal paraDeptName As String)
    Set pobjItem = pobj体检结果业务.func获取体检人员单科室体检结果(pstrPerson, paraDeptName)
    With cgrdItem
        Set .DataSource = pobjItem
        '修改人：张令 2012.12.12   ↓↓
        '取消设置行为0，否则没有数据。
'        .Col = 0
        '修改人：张令 2012.12.12   ↑↑
        .Sort = flexSortStringAscending
        .AutoSize 0, .cols - 1, 0, 0
        .SelectionMode = flexSelectionListBox
        .AllowSelection = False
    End With
    '修改人：张令 2012.12.12     ↓↓
    '事件重复执行，没有意义。
'    cchkUnfilled_Click
'    cchkAbnormal_Click
    '修改人：张令 2012.12.12     ↑↑
End Sub

'2012-04-12 于登淼
'清空界面上控件内容
Sub subClear()
    
    '界面控件初始化
    cchkBarCode.Value = 0
    ctxtBarCode.Text = ""
    cchkCompanyName.Value = 0
    ctxtCompanyName.Text = ""
    cgrdInfo.rows = 1
    cgrdDept.rows = 1
    cgrdItem.rows = 1
    ctxtConclusion.Text = ""
    ctxtDiagnose.Text = ""
    ctxtReview.Text = ""
    ctxtReviewItem.Text = ""
    ctxtDetpConclusion.Text = ""
    '修改人：罗李奎 2012-12-10 ↓
    'bug号：0000060
    '体检表模版还原
    cchkTemplate.Value = 0
    sub加载体检表模板
     '时间还原为当前电脑时间
    cdtpDateTo.Value = Date
    cchkDate.Value = 0
     '体检条码号还原
    ctxtBarCode.Text = ""
    cchkBarCode.Value = 0
    '单位名称还原
    ctxtCompanyName.Text = ""
    cchkCompanyName.Value = 0
      '修改人：罗李奎 2012-12-10 ↑
End Sub

'2012-04-12 于登淼
'保存结论和诊断意见。
'将“最终结论录入”看做一个科室，结论和诊断意见合为一个字符串，当做一个总的结论存入“科室结论表”中
Sub subSaveConclusion()
    Dim i, selRow As Integer
    For i = 0 To cgrdInfo.SelectedRows - 1
        selRow = cgrdInfo.SelectedRow(i)
        ' "_00_" 没有什么特殊含义，只是作为分割结论和诊断意见的分隔符
        pobj业务对象.sub单个填写体检结论 cgrdInfo.TextMatrix(selRow, 0), "最终结论录入", ctxtConclusion.Text & "_00_" & Trim(ctxtDiagnose.Text), um用户编号, Trim(ctxtReview.Text), Trim(ctxtReviewItem.Text)
    Next
    If cgrdInfo.SelectedRows > 0 Then
        MsgBox ("保存成功！")
        '清空评价信息
        ctxtConclusion.Text = ""
        ctxtDiagnose = ""
    End If
End Sub

'2012-06-25 于登淼
'添加函数自动在结论text中，添加自动填入不合格体检项目
Sub sub自动填入不合格项目(ByVal paraSysNo As String)
    Dim lobjRec As Object
    Dim strSQL As String
    Dim i As Integer
    strSQL = "select distinct b.名称 from 职业病体检_体检结果视图 a,职业病体检_体检项目设置表 b where a.系统编号='" & paraSysNo & "' and a.体检项目=b.编码 and a.单项结论='不合格'"
    Set lobjRec = dafuncGetData(strSQL)
    If lobjRec.RecordCount > 0 Then
        ctxtConclusion.Text = "该次体检不合格项目："
        lobjRec.MoveFirst
        For i = 1 To lobjRec.RecordCount
            If i <> 1 Then
                ctxtConclusion.Text = ctxtConclusion.Text & "、" & lobjRec("名称")
            Else
                ctxtConclusion.Text = ctxtConclusion.Text & lobjRec("名称")
            End If
            lobjRec.MoveNext
        Next
        ctxtConclusion.Text = ctxtConclusion.Text & "。" & vbCrLf
    End If
End Sub

'2012-08-22 于登淼
'找出所有未体检完科室与项目
Private Function sub未体检完科室与体检项目(ByVal paraSysNo As String) As String 'vbcrlf
    Dim strSQL As String
    Dim lobjRec As Object
    Dim i As Integer, j As Integer
    Dim resultStrDept, resultStrItem As String
    
    strSQL = "select distinct a.体检项目,b.名称 from 职业病体检_体检结果视图 a, 职业病体检_体检项目设置表 b where a.系统编号='" & paraSysNo & "' and (a.体检结果='' or a.体检结果 is null) and a.体检项目=b.编码"
    Set lobjRec = dafuncGetData(strSQL)
    If lobjRec.RecordCount > 0 Then
        lobjRec.MoveFirst
        resultStrItem = "未体检项目有："
        For i = 0 To lobjRec.RecordCount - 1
'            resultStrItem = IIf(i = 0, resultStrItem & lobjRec("体检项目") & lobjRec("名称"), resultStrItem & "，" & lobjRec("体检项目") & lobjRec("名称"))
            resultStrItem = IIf(i = 0, resultStrItem & lobjRec("名称"), resultStrItem & "，" & lobjRec("名称"))
            lobjRec.MoveNext
        Next i
        resultStrItem = resultStrItem & "。"
    End If
    sub未体检完科室与体检项目 = resultStrItem
End Function
'供WORD宏调用，用于保存WORD至数据库
Public Sub subSave(ByVal paraFile As String, ByVal paraNo As Integer, ByVal para系统编号 As String)
    subSaveDoc paraFile, paraNo, para系统编号
End Sub
'打印体检表 2015-11-6 by lanchao update print
Private Sub subPrintold(ByVal para预览 As Boolean)
    Dim i As Integer
    Dim lobj文书 As Object
    Dim lcolSysNo As Collection
    On Error GoTo errHandler
    Set lobj文书 = CreateObject("职业病文书.cls文书")
'    sum = 0
    With cgrdInfo
        Set lcolSysNo = New Collection
'        For i = 1 To .rows
'             If .Cell(flexcpChecked, i, 0) = flexChecked Then
                    lcolSysNo.Add .TextMatrix(.Row, 0)
                    If Right(lcolSysNo(1), 1) = "F" Then
                    '这里只能看职业体检表的，
                        lobj文书.Sub打印文书 "职业健康体检_" & .TextMatrix(.Row, 5) & "F", lcolSysNo, para预览
                    Else
                        lobj文书.Sub打印文书 "职业健康体检_" & .TextMatrix(.Row, 5), lcolSysNo, para预览
                    End If
                    If para预览 = False Then
                        dafuncGetData "update 职业病体检_体检基本信息表 set 体检状态='7' where 系统编号='" & Trim(.TextMatrix(i, 0)) & "'"
                        .RowHidden(i) = True
                    End If
'            End If
'        Next i
'         If lcolSysNo.Count < 1 And .rows > 1 Then
'            MsgBox "请勾选要打印或预览的体检表！", vbInformation, "系统提示"
'            Exit Sub
'        End If
   
    End With
errHandler:
    
End Sub
            
'打印体检表
'2015-11-9 牟俊
Private Sub subPrint(ByVal para预览 As Boolean)
    Dim sql As String
    Dim lobjet As Object
    Dim mstr体检表类型 As String
    Dim i As Integer
    Dim lobj文书 As Object
    Dim lcolSysNo As Collection
    On Error GoTo errHandler
  
    Set lobj文书 = CreateObject("职业病文书.cls文书")
'    sum = 0
    With cgrdInfo
'        For i = 1 To .rows - 1
'            If .Cell(flexcpChecked, i, 0) = flexChecked Then
             Set lcolSysNo = New Collection
                lcolSysNo.Add .TextMatrix(cgrdInfo.Row, 0)
                 
'                 lobj文书.Sub打印文书 "职业健康体检_" & .TextMatrix(i, mcolIndex("体检类型")), lcolSysNo, para预览

                mstr系统编号 = cgrdInfo.TextMatrix(cgrdInfo.Row, mcolIndex("系统编号"))    '取出系统编号
                
                '取出体检类型，因为打印报告是根据体检类型来判断的打印哪张表
                sql = "select 体检类型 from 职业病体检_体检基本信息表 where 系统编号='" & mstr系统编号 & "'"
                Set lobjet = dafuncGetData(sql)
                mstr体检表类型 = lobjet(0)
'                mstr体检表类型 = cgrdInfo.TextMatrix(cgrdInfo.Row, mcolIndex("体检表编号"))
'                Set lcolSysNo = New Collection
'                lcolSysNo.Add .TextMatrix(i, 0)
'                 Dim tst As String
'                 tst = "职业健康体检_" + mstr体检表类型
                 lobj文书.Sub打印文书 "职业健康体检_" + mstr体检表类型, lcolSysNo, para预览
'                 lobj文书.Sub打印文书 "职业健康体检_" & .TextMatrix(i, mstr体检表类型), mstr系统编号, para预览
'                lobj文书.Sub打印文书 "职业健康体检_" & .TextMatrix(i, mcolIndex("体检类型")), mstr系统编号, para预览
                
                If para预览 = False Then
'                    dafuncGetData "update 职业病体检_体检基本信息表 set 体检状态='7' where 系统编号='" & Trim(.TextMatrix(i, 0)) & "'"
                    dafuncGetData "update 职业病体检_体检基本信息表 set 体检状态='7' where 系统编号='" & mstr系统编号 & "'"
                    .RowHidden(i) = True
                End If
'            End If
'        Next i
        If lcolSysNo.Count < 1 And .rows > 1 Then
            MsgBox "请勾选要打印或预览的体检表！", vbInformation, "系统提示"
            Exit Sub
        End If
    End With
errHandler:
   
End Sub
