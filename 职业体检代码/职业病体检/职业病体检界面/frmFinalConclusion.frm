VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmFinalConclusion 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "���ս���¼�봰��"
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
            Caption         =   "��������Ŀ���������"
            Height          =   5415
            Left            =   360
            TabIndex        =   18
            Top             =   6480
            Width           =   16215
            Begin VB.TextBox Text������� 
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
               Caption         =   "����δ������"
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
               Caption         =   "�������ʾ��������"
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
               FormatString    =   "��Ŀ���|�����Ŀ|�����|���ҽʦ|�������"
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
               ToolTipText     =   "˫���鿴�������"
               Top             =   240
               Width           =   11175
               _cx             =   2088783103
               _cy             =   2088769345
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
               FormatString    =   "����|���ֽ���|ҽʦ����"
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
               Caption         =   "���ҽ��������"
               Height          =   255
               Left            =   7920
               TabIndex        =   56
               Top             =   3840
               Visible         =   0   'False
               Width           =   2415
            End
         End
         Begin VB.Frame fraFinal 
            Caption         =   "��д���ս���"
            Height          =   5895
            Left            =   12480
            TabIndex        =   24
            Top             =   600
            Width           =   3855
            Begin VB.CommandButton cmdѡ�񸴲���Ŀ 
               Caption         =   "ѡ�񸴲���Ŀ"
               Height          =   375
               Left            =   1800
               TabIndex        =   63
               Top             =   4440
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.CommandButton comd������� 
               Caption         =   "�������"
               Height          =   375
               Left            =   2640
               TabIndex        =   58
               Top             =   2040
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.CommandButton Command1 
               Caption         =   "���ѡ�������������"
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
            Begin VB.CommandButton Cmd����ģ�� 
               Caption         =   "�������ģ��"
               Height          =   345
               Index           =   1
               Left            =   1200
               TabIndex        =   46
               Top             =   2040
               Width           =   1335
            End
            Begin VB.CommandButton Cmd����ģ�� 
               Caption         =   "����ģ��"
               Height          =   315
               Index           =   0
               Left            =   1920
               TabIndex        =   45
               Top             =   720
               Width           =   1455
            End
            Begin VB.OptionButton cchk��׼ 
               Caption         =   "�踴��"
               ForeColor       =   &H000000FF&
               Height          =   255
               Index           =   1
               Left            =   2400
               TabIndex        =   42
               Top             =   240
               Width           =   975
            End
            Begin VB.OptionButton cchk��׼ 
               Caption         =   "������"
               Height          =   255
               Index           =   0
               Left            =   1200
               TabIndex        =   41
               Top             =   240
               Width           =   855
            End
            Begin VB.ComboBox cbox����ģ�� 
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
               Caption         =   "������Ŀ��"
               Height          =   255
               Left            =   120
               TabIndex        =   50
               Top             =   4560
               Width           =   1095
            End
            Begin VB.Label Label6 
               BackColor       =   &H00FFC0FF&
               Caption         =   "����ԭ��"
               Height          =   255
               Left            =   120
               TabIndex        =   48
               Top             =   3600
               Width           =   1095
            End
            Begin VB.Label Label2 
               BackColor       =   &H00FFC0FF&
               Caption         =   "ģ��ɸѡ��"
               Height          =   255
               Left            =   120
               TabIndex        =   35
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label4 
               BackColor       =   &H00FFC0FF&
               Caption         =   "���������"
               Height          =   255
               Left            =   120
               TabIndex        =   31
               Top             =   2040
               Width           =   1095
            End
            Begin VB.Label Label3 
               BackColor       =   &H00FFC0FF&
               Caption         =   "�������"
               Height          =   255
               Left            =   120
               TabIndex        =   30
               Top             =   720
               Width           =   975
            End
            Begin VB.Label Label1 
               BackColor       =   &H00FFC0FF&
               Caption         =   "�������ڣ�"
               Height          =   255
               Left            =   120
               TabIndex        =   28
               Top             =   3360
               Width           =   1095
            End
            Begin VB.Label llabDoctor 
               BackColor       =   &H00FFC0FF&
               Caption         =   "����ҽʦ��"
               Height          =   255
               Left            =   120
               TabIndex        =   29
               Top             =   3480
               Visible         =   0   'False
               Width           =   1935
            End
         End
         Begin VB.Frame fraPerson 
            Caption         =   "��ѯ��Ա��Ϣ"
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
               Caption         =   "��������"
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
               Caption         =   "�����棬ѡ���ж�����Ϊ��ǰ�Ľ��ۺʹ����������BackSpace�Ƴ�ѡ����"
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   120
               TabIndex        =   17
               ToolTipText     =   "˫���鿴���ҽ���"
               Top             =   240
               Width           =   8175
            End
         End
         Begin VB.Frame fraQuery 
            Caption         =   "ɸѡ�����Ա"
            Height          =   5535
            Left            =   240
            TabIndex        =   2
            Top             =   840
            Width           =   3735
            Begin VB.CommandButton Com���� 
               Caption         =   "����"
               Height          =   375
               Left            =   120
               TabIndex        =   62
               Top             =   3480
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.CommandButton Com�˻� 
               Caption         =   "�˻ظ���"
               Height          =   375
               Left            =   2640
               TabIndex        =   61
               Top             =   3480
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.OptionButton coptType 
               BackColor       =   &H00FFFFC0&
               Caption         =   "δ�½���"
               Height          =   255
               Index           =   6
               Left            =   2160
               TabIndex        =   57
               Top             =   4080
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.CommandButton Command2 
               Caption         =   "�鿴��Ϣ"
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
               Caption         =   "�� ѯ"
               Height          =   375
               Left            =   480
               TabIndex        =   3
               Top             =   4920
               Width           =   1095
            End
            Begin VB.CheckBox cchkTemplate 
               BackColor       =   &H00C0FFC0&
               Caption         =   "����ģ��"
               Height          =   300
               Left            =   240
               TabIndex        =   43
               Top             =   960
               Width           =   1335
            End
            Begin VB.OptionButton coptType 
               BackColor       =   &H00FFFFC0&
               Caption         =   "������"
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
               Caption         =   "�ѷ�����"
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
               Caption         =   "�Ѹ���"
               Height          =   255
               Index           =   3
               Left            =   480
               TabIndex        =   38
               Top             =   4560
               Width           =   1095
            End
            Begin VB.OptionButton coptType 
               BackColor       =   &H00FFFFC0&
               Caption         =   "������"
               Height          =   255
               Index           =   2
               Left            =   2160
               TabIndex        =   37
               Top             =   4560
               Width           =   1095
            End
            Begin VB.OptionButton coptType 
               BackColor       =   &H00FFFFC0&
               Caption         =   "δ�½���"
               Height          =   255
               Index           =   1
               Left            =   2160
               TabIndex        =   36
               Top             =   4080
               Width           =   1095
            End
            Begin VB.ComboBox Ccmb�������� 
               Enabled         =   0   'False
               Height          =   300
               Left            =   2040
               TabIndex        =   13
               Text            =   "������"
               Top             =   360
               Width           =   1575
            End
            Begin VB.ComboBox ccmb��������� 
               Enabled         =   0   'False
               Height          =   300
               Left            =   120
               TabIndex        =   12
               Text            =   "�����Ա����"
               Top             =   360
               Width           =   1695
            End
            Begin VB.OptionButton coptType 
               BackColor       =   &H00FFFFC0&
               Caption         =   "�����"
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
               Caption         =   "������� ��"
               Height          =   255
               Left            =   240
               TabIndex        =   9
               Top             =   1440
               Value           =   1  'Checked
               Width           =   1335
            End
            Begin VB.CheckBox cchkCompanyName 
               BackColor       =   &H00C0FFC0&
               Caption         =   "��λ����"
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
               Caption         =   "��������"
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
               Caption         =   "��λ��λ"
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
               Caption         =   "��"
               Height          =   180
               Left            =   1320
               TabIndex        =   52
               Top             =   1920
               Width           =   180
            End
         End
         Begin MSComctlLib.Toolbar ctlb������ 
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
'2012-04-10 �ڵ��
'���� ���ս���¼�봰�壬����Ӧ��������
'��Ͳ���Ȩ�ޣ�1����ѯ����������Ա�ĸ������ҽ��ۺ͸��Ƹ����������
'              2�����޸Ĳ��ܱ��棬Ҳ���ܴ�ӡ���档

Option Explicit
Public mblnInUse As Boolean
Private WithEvents mobjGUI As cls����ͨ�ö���
Attribute mobjGUI.VB_VarHelpID = -1
Private mstr��쵥�� As String
'Private mstrϵͳ��� As String
Private mlobjRec As Object
Private mstrȨ�ޱ�־ As Boolean

'��ѯ���
Private mstrDoctorName As String
Private mobjQueryResult As Object
Private mcolIndex As New Collection
Private indX, indY As Integer       '��¼�����vsflexgrid�����ꡣ
Private resql As String     '��¼ÿ�β�ѯ��sql
 
'�ý��湲�ö���
Private pobj����ģ�� As Object
Private pobj��� As Object
Private pobj�����ҵ�� As Object
Private pobj���� As Object
Private pstrPerson As String        '��ǰ���������Աϵͳ���,cgrdInfo˫�������
Private pobjItem As Object

Private mstrSearchString As String
Private mstr�������� As String

'2012-06-21 �ڵ��
'��ǵ�ǰѡ�е������Ա���״̬��
'��Ҫ�����ж� δ�½��ۡ����½��ۡ�����ˡ��ѷ��ű��桢�����顣
Private mstrState As String

'2012-08-22 �ڵ�� ��
'��ӿ��ұ���
Private pobjDept As Object
'2012-08-22 �ڵ�� ��

'2012-04-10 �ڵ��
'���ܣ����ص�ǰ�����Ƿ��Ѿ����ر�־������ϵͳƽ̨��Ҫ��ġ�
Public Property Get pblnInUse() As Boolean
    On Error Resume Next
    pblnInUse = mblnInUse
End Property
'���ܣ�ѡ�����ģ��
'���ߣ�����
'ʱ�䣺2012-06-01
Private Sub cbox����ģ��_Click()
    '2012-06-25 �ڵ�� ��
    'ģ�������Ҫ�س������ĩβ
    ctxtConclusion.Text = ctxtConclusion.Text & cbox����ģ��.Text
    ctxtDiagnose.Text = ctxtDiagnose.Text & cbox����ģ��.Text    '2015-10-16
    '2012-06-25 �ڵ�� ��
End Sub

'2012-04-12 �ڵ��
'����ʾ���ϸ��Ϊ�˷�������ҽʦ�����ս���ʱ���鿴�����
Private Sub cchkAbnormal_Click()
    Dim i As Integer
    If cchkAbnormal.Value = 1 Then      '����ʾ��������
        For i = 1 To cgrdItem.rows - 1: cgrdItem.RowHidden(i) = False: Next
        For i = 1 To cgrdItem.rows - 1  'Ĭ�ϵ�4��Ϊ�������
            If cgrdItem.TextMatrix(i, 4) = "�ϸ�" Then cgrdItem.RowHidden(i) = True
        Next
    Else
        '�޸��ˣ����� 2012.12.12        ����
        '�����������ʾ�������δ��ѡʱ��������ʾ���ݡ�
'        If cchkUnfilled.Value = 0 Then Exit Sub
'        cgrdItem.Clear
'        Set cgrdItem.DataSource = pobjItem
'        For i = 1 To cgrdItem.rows - 1
'            If cgrdItem.TextMatrix(i, 4) = "���ϸ�" Then cgrdItem.RowHidden(i) = False
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
    '�޸��ˣ����� 2012.12.12       ����
    '��������δ�����Ϊ��ѡ�����������ʾ�������Ϊδ��ѡʱִ��cchkUnfilled_Click
    If cchkUnfilled.Value = 1 And cchkAbnormal.Value = 0 Then
        cchkUnfilled_Click
    End If
    '�޸��ˣ����� 2012.12.12       ����
End Sub


Private Sub cchkTemplate_Click()
    If cchkTemplate.Value = 1 Then
        ccmb���������.Enabled = True
        Ccmb��������.Enabled = True
    Else
        ccmb���������.Enabled = False
        Ccmb��������.Enabled = False
    End If
End Sub

'2012-04-12 �ڵ��
'ȥ��û����д������Ϊ�˷�������ҽʦ�����ս���ʱ���鿴�����
Private Sub cchkUnfilled_Click()
    Dim i As Integer
    If cchkUnfilled.Value = 1 Then      '����δ��д��
        For i = 1 To cgrdItem.rows - 1  'Ĭ�ϵ�1��Ϊ�����
            If cgrdItem.TextMatrix(i, 2) = "" Or (IsNull(cgrdItem.TextMatrix(i, 2)) = True) Then cgrdItem.RowHidden(i) = True
        Next
    Else
        '�޸��ˣ����� 2012.12.12        ����
        '��������δ�����δ��ѡʱ��������ʾ���ݡ�
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
        '�޸��ˣ����� 2012.12.12       ����
    End If
    '�޸��ˣ����� 2012.12.12       ����
    '��������δ�����Ϊδ��ѡ�����������ʾ�������Ϊ��ѡʱִ��cchkAbnormal_Click
    If cchkAbnormal.Value = 1 And cchkUnfilled.Value = 0 Then
        cchkAbnormal_Click
    End If
    '�޸��ˣ����� 2012.12.12       ����
End Sub

'���ܣ�ʵ�ֽ���ģ���ɸѡ
'���ߣ�����
'ʱ�䣺2012-05-31
Private Sub cchk��׼_Click(Index As Integer)
    
    Dim lobj���� As Object
    Dim pub���� As Object
    Dim i As Integer
    Dim sql As String 'func��ȡ�������ս���ģ��
    Set pub���� = CreateObject("ְҵ������.clsConclusionSet")
    
    '2012-07-03 �ڵ�� ��
    '����ѡ���Ϊ��ѡ���ж�ֵ���ж�������΢�Ķ�
    'If cchk��׼(0).Value = 1 And cchk��׼(1).Value = 0 Then
    If cchk��׼(0).Value = True Then
        sql = "select * from ϵͳ����_�ֵ�_������ģ��� where ���۱�׼='�ϸ�' and ���ұ��=(select ��� from ϵͳ����_�ֵ�_�ֵ����ݱ� where ID='84' and ����='���ս���¼��')"
        Set lobj���� = pub����.func��ȡ�������ս���ģ��(sql)
        cbox����ģ��.Clear
        'lobj����.MoveFirst
        For i = 1 To lobj����.RecordCount
            cbox����ģ��.AddItem lobj����("����ģ��")
            lobj����.MoveNext
        Next i
        ctxtReview.Text = ""
        ctxtReviewItem.Text = ""
    End If
    'If cchk��׼(0).Value = 0 And cchk��׼(1).Value = 1 Then
    If cchk��׼(1).Value = True Then
        sql = "select * from ϵͳ����_�ֵ�_������ģ��� where ���۱�׼='���ϸ�' and ���ұ��=(select ��� from ϵͳ����_�ֵ�_�ֵ����ݱ� where ID='84' and ����='���ս���¼��')"
        Set lobj���� = pub����.func��ȡ�������ս���ģ��(sql)
        cbox����ģ��.Clear
        'lobj����.MoveFirst
        For i = 1 To lobj����.RecordCount
            cbox����ģ��.AddItem lobj����("����ģ��")
            lobj����.MoveNext
        Next i
    End If
'''    If cchk��׼(0).Value = 1 And cchk��׼(1).Value = 1 Then
'''        sql = "select * from ϵͳ����_�ֵ�_������ģ��� and ���ұ��=(select ��� from ϵͳ����_�ֵ�_�ֵ����ݱ� where id='84' and ����='���ս���¼��')"
'''        Set lobj���� = pub����.func��ȡ�������ս���ģ��(sql)
'''        cbox����ģ��.Clear
'''        'lobj����.MoveFirst
'''        For i = 1 To lobj����.RecordCount
'''            cbox����ģ��.AddItem lobj����("����ģ��")
'''            lobj����.MoveNext
'''        Next i
'''    End If
'''    If cchk��׼(0).Value = 0 And cchk��׼(1).Value = 0 Then
'''        cbox����ģ��.Clear
'''    End If
    If cchk��׼(0).Value = True Then
        ctxtReview.Enabled = False
        ctxtReviewItem.Enabled = False
        cmdѡ�񸴲���Ŀ.Enabled = False
    Else
        ctxtReview.Enabled = True
        ctxtReviewItem.Enabled = True
        cmdѡ�񸴲���Ŀ.Enabled = True
    End If
    '2012-07-03 �ڵ�� ��
End Sub

'2012-04-11 �ڵ��
'�������б������������ҳ���Ӧ����ģ��
Private Sub Ccmb��������_Click()
    Dim lobj����ģ�弯 As Object
    Dim lobj������ As Object
    Dim lcolInfo As New Collection
    Dim lcol������ As Collection
    Dim i As Integer
    On Error GoTo errHandler
    
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
    sfsub������ "ְҵ������", "frmFinalConclusion", "ccmb��������_click", Err.Number, Err.Description, True
End Sub

'2012-04-11 �ڵ��
'�����Ա����б���������ͬʱ����������б�������
Private Sub ccmb���������_Click()
    Dim lobj������� As Object
    On Error GoTo errHandler
    Set lobj������� = CreateObject("ְҵ������.clsmedicalexam")
    lobj�������.������� = ccmb���������.ItemData(ccmb���������.ListIndex)
    Call Ccmb��������_Click
    Exit Sub
errHandler:
    sfsub������ "ְҵ������", "frmFinalConclusion", "Private Sub ccmb���������_Click", Err.Number, Err.Description, True
End Sub

'2012-04-10 �ڵ��
'��λ��λ��Ϊ�˷����ѯĳ����λ��Ա���������Ϣ
Private Sub ccmdLocate_Click()
    On Error GoTo errHandler
    Dim lobjRec As Object                       '��λ��λ���صĽ����¼��
    Set lobjRec = pobjҵ�����.func��λ��λ     '������λ��λ���档
    
    '��ȡ��λ�ĵ�λ����ʾ�ڡ���λ���ơ�¼����С�(��ʱֻ��ʾ����λ���ơ�)
    '-----��֪�������費��Ҫ������ģ�������趨��˲��ӡ�
    If Not lobjRec Is Nothing Then
        If lobjRec.RecordCount > 0 Then
            ctxtCompanyName.Text = IIf(IsNull(lobjRec("��λ����")), "", lobjRec("��λ����"))
        End If
    End If
    Set lobjRec = Nothing
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ������", "frmFinalConclusion", "ccmdLocate_Click", 6666, lstrError, False
End Sub

''''2012-04-11 �ڵ��
''''����ѯ���ֱ�Ӽӵ�cgrdInfo�б�ĺ��档
'''Private Sub ccmdAdd_Click()
'''    On Error GoTo errHandler
'''
'''
'''    Dim lobjTmp, lobjRec As Object
'''    Dim i As Integer, j As Integer
'''    Dim lstrWhere As String
'''
'''    lstrWhere = " and ������='" & ccmbTemplate.Text & "'"      '
'''    lstrWhere = " and �������='" & ccmb��������� & "'"
'''
'''    '��װ��ѯ����
'''    If cchkDate.Value = 1 Then          '�������
'''        lstrWhere = lstrWhere & " and �������='" & Format(cdtpDateTo.Value, "yyyy-mm-dd hh:mm:ss") & "'"
'''    End If
'''
'''    If cchkBarCode.Value = 1 Then                         '��������
'''        lstrWhere = lstrWhere & " and ϵͳ���='" & Trim(ctxtBarCode) & "'"
'''    End If
'''
'''    If cchkCompanyName.Value = 1 Then                         '��λ����
'''        lstrWhere = lstrWhere & " and ��λ����='" & Trim(ctxtCompanyName) & "'"
'''    End If
'''
'''    If coptConclusion(0).Value = True Then                          '���½��ۡ�δ�½���(�ܽ���)
'''        lstrWhere = lstrWhere & " and ((������ is null) or ������='')"
'''    Else
'''        lstrWhere = lstrWhere & " and (������ is not null)"
'''    End If
'''
'''    If reSql = "0" Or reSql <> lstrWhere Then
'''        reSql = lstrWhere
'''        Set lobjTmp = CreateObject("ְҵ�������¼��.clscommon")
'''        Set lobjRec = lobjTmp.func��ȡ���޸Ľ��۵�_�ض����ҵ�_�����Ա������Ϣ(lstrWhere, "")
'''         '�����л����������Щ�У�����������Ϣ��ʾ
'''        If lobjRec.RecordCount > 0 Then
'''            lobjRec.MoveFirst
'''            With cgrdInfo
'''                For i = 1 To lobjRec.RecordCount
'''                    .AddItem ("")
'''                    For j = 0 To .Cols - 1
'''                        If .TextMatrix(0, j) = "��������" Then
'''                            .TextMatrix(.Rows - 1, j) = lobjRec("ϵͳ���")
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
'''    lstrError = func������(Err.Number, Err.Description)
'''    sfsub������ "ְҵ������", "frmFinalConclusion", "ccmdQuery_Click", 6666, lstrError, False
'''End Sub

'2012-07-03 �ڵ��
'����ѯ�������cgrdinfo�б��У�����֮ǰ�Ľ����
Private Sub ccmdQuery_Click()
    Dim lstr��������, lstr��λ����, lstrϵͳ��� As String
    Dim lstr��ʼ����, lstr�������� As Date
    Dim lstr�������� As String           '2015-11-9 Ĳ�� ������������
    Dim i As Integer
    
    If cchkDate.Value = 1 Then
        '�޸��ˣ����� 2012.12.05
        'bug�ţ�0000071
        '˵������ʼ�������������һ�������ѯ�����ڲ�������Ϊһ���0�㵽23�㡣  ����
'        lstr��ʼ���� = cdtpDateTo.Value: lstr�������� = cdtpDateTo.Value
        lstr��ʼ���� = CStr(Format(cdtpDateTo.Value, "yyyy/mm/dd"))
        lstr�������� = CStr(Format(cdtpDateFrom.Value, "yyyy/mm/dd"))
        '2012.12.05    ����
    Else
        lstr��ʼ���� = "1900-01-01 00:00:00": lstr�������� = "3000-01-01 00:00:00"
    End If
    
    If cchkTemplate.Value = 1 Then lstr�������� = ccmbTemplate.Text
    If cchkBarCode.Value = 1 Then lstrϵͳ��� = ctxtBarCode.Text
    If cchkCompanyName.Value = 1 Then lstr��λ���� = ctxtCompanyName.Text
    
    Set mobjQueryResult = pobjҵ�����.func����������ѯ(lstr��ʼ����, lstr��������, lstr��������, lstr��λ����, "", "", "", lstrϵͳ���, "")
'    Set mobjQueryResult = pobjҵ�����.func����������ѯ(lstr��ʼ����, lstr��������, lstr��������, lstr��λ����, "", "", "", lstrϵͳ���, "", lstr��������)
'    For i = 0 To 5
'        If coptType(i).Value = True Then mobjQueryResult.Filter = "���״̬='" & coptType(i).Caption & "'": Exit For
    '�޸��ˣ������ 2012-12-12 ��
'        If coptType(i).Index = 0 Then
        If coptType(0).Value = True Then
            mobjQueryResult.Filter = "���״̬='" & coptType(i).Caption & "'  or ���״̬='δ¼���ܼ��߸�����Ϣ'"
'        Else
'            If coptType(i).Value = True Then mobjQueryResult.Filter = "���״̬='" & coptType(i).Caption & "'": Exit For
        ElseIf coptType(1).Value = True Then
            mobjQueryResult.Filter = "���״̬='δ�½���'"
        ElseIf coptType(2).Value = True Then
            mobjQueryResult.Filter = "���״̬='������'"
        ElseIf coptType(3).Value = True Then
            mobjQueryResult.Filter = "���״̬='�Ѹ���' or ���״̬='�ѷ�����' or ���״̬='������'"
        ElseIf coptType(6).Value = True Then     '2016-4-20 by Ĳ��
            mobjQueryResult.Filter = "���״̬='δ�½���'"
        End If
        '�޸��ˣ������ 2012-12-12 ��
'    Next
    Set cgrdInfo.DataSource = mobjQueryResult
    Set mcolIndex = New Collection
    For i = 0 To cgrdInfo.cols - 1
        mcolIndex.Add i, cgrdInfo.TextMatrix(0, i)
    Next
    cgrdInfo.ColHidden(mcolIndex("�Թܱ��")) = True
    cgrdInfo.ColHidden(mcolIndex("������")) = True
    
    'vsflexgrid�п�Ȱ������Զ������������ͷ����ͷ����������
    cgrdInfo.AutoSize 0, cgrdInfo.cols - 1, 0, 0
    cgrdInfo.ExplorerBar = flexExSort
    cgrdInfo.DataMode = flexDMFree
    cgrdInfo.Col = 0
    cgrdInfo.Sort = flexSortGenericDescending
    
    ctxtConclusion.Text = ""  '��ս��ۺͽ���,���ر��������ť�ͽ��������,�˻ظ��˰�ť��  2016-4-20 by Ĳ��
    ctxtDiagnose.Text = ""
    comd�������.Visible = False
    Text�������.Visible = False
    Com�˻�.Visible = False
    Label10.Caption = cgrdInfo.rows - 1
End Sub

'2012-04-12 �ڵ��
'�����鿴����������������
Private Sub cgrdDept_Click()
    Dim strTmp As String
    If cgrdDept.Row = 0 Then Exit Sub
    strTmp = cgrdDept.TextMatrix(cgrdDept.Row, 0)
    strTmp = Right(strTmp, Len(strTmp) - 3)
    sub�г���������������� (strTmp)     '�̶���0��Ϊ��������
    '�޸��ˣ����� 2012.12.12    ����
    '����ť��ѡʱִ�е����¼���
    If cchkUnfilled.Value = 1 Then
        cchkUnfilled_Click
    ElseIf cchkAbnormal.Value = 1 Then
        cchkAbnormal_Click
    End If
    '�޸��ˣ����� 2012.12.12    ����
    
    '���� ���ƽ������   2015-11-5
    Dim Jlstr As String
    Dim JstrTmp As String
    Dim Jsql As Object
    JstrTmp = cgrdDept.TextMatrix(cgrdDept.Row, 0)
    JstrTmp = Left(JstrTmp, 2)
    Jlstr = "select ���ֽ��� from ְҵ�����_���ҽ��۱� where ����='" & JstrTmp & " ' and ϵͳ���='" & mstrϵͳ��� & "' "
    Set Jsql = dafuncGetData(Jlstr)
    Text�������.Text = Jsql("���ֽ���")
    Text�������.Visible = True
'    Text�������.Top = 3000
    Text�������.Top = ctxtDetpConclusion.Top
    Text�������.Left = ctxtDetpConclusion.Left + ctxtDetpConclusion.Width
    Text�������.Width = cgrdDept.Width - ctxtDetpConclusion.Width
    Label10.Caption = cgrdInfo.rows
End Sub

'˫��ѡ�У��οƳ�Ҫ�󣬴Ӵ˴��������
Private Sub cgrdDept_DblClick()
    cgrdDept.EditCell
    Text�������.Visible = False
End Sub

'�޸��ˣ����� 2012.12.12           ����
'�����뵥���¼�һ�����ظ��ˡ�

'2012-04-12 �ڵ��
'˫���鿴ÿ�����ҵĵ�����
'Private Sub cgrdDept_DblClick()
'    Dim strTmp As String
'    If cgrdDept.Row = 0 Then Exit Sub
'    strTmp = cgrdDept.TextMatrix(cgrdDept.Row, 0)
'    strTmp = Right(strTmp, Len(strTmp) - 3)
'    sub�г���������������� (strTmp)     '�̶���0��Ϊ��������
'    cchkAbnormal_Click
'End Sub


'2012-04-12 �ڵ��
'˫���鿴���ҽ���
Private Sub cgrdInfo_DblClick()
        cgrdDept.Clear
        cgrdItem.Clear
        pstrPerson = cgrdInfo.TextMatrix(cgrdInfo.Row, 0)
        sub�г������ҽ���
        '���ܣ������ʼ����ʱ������
        '���ߣ�����
        'ʱ�䣺2012-06-01
'        cbox����ģ��.Visible = True
        cchk��׼(1).Visible = True
        cchk��׼(0).Visible = True
        'ʱ�䣺2012-06-01
        
'''        '2012-06-21 �ڵ�� ��
'''        '��ѯ��ǰ�����Ա�����״̬���ж��Ƿ������ˣ����ĵ�ǰ����������״̬
'''        Dim lobjRec As Object
'''        Set lobjRec = pobjҵ�����.func����������ѯ("", "", "", "", "", "", "", pstrPerson)
'''        mstrState = lobjRec("���״̬")
'''        ctlb������.Buttons(5).Enabled = (mstrState = "���½���")
'''        ctlb������.Buttons(7).Enabled = True   'Ԥ������
'''        Set lobjRec = Nothing
'''        '2012-06-21 �ڵ�� ��
        
        '2012-06-25 �ڵ�� ��
        '��첻�ϸ���Ŀ�Զ���������ܽ����
         If coptType(1).Value = True Then
            sub�Զ����벻�ϸ���Ŀ pstrPerson
        End If
       '2012-06-25 �ڵ�� ��
       
        '˫��ʱ������ģ��Ĭ���С����飺������       2015-10-19
        If ctxtDiagnose = "" Then
       ctxtDiagnose.Text = "���飺"
        End If
        
       'ѡ���˺󡰸�����Ϣ����ť������   2015-10-22
         Command2.Enabled = True
         Command2.Visible = True
         
    '˫��ʱ8023���������Ҫ��Ҫ��˫���͵���Ч��һ����      2015-10-26   Ĳ��
    Dim resql As Object
    mstrϵͳ��� = cgrdInfo.TextMatrix(cgrdInfo.Row, mcolIndex("ϵͳ���"))
    Set resql = dafuncGetData("select �������� From ְҵ�����_�����Ա������Ϣ�� where ϵͳ���='" & mstrϵͳ��� & "'")
'    If resql("��������") = "8023����" Or resql("��������") = "��˲���" Then
    If resql("��������") = "8023����" Or resql("��������") = "��˲���" Or resql("��������") = "��˲���YK" Then
        If mobjQueryResult.Filter = "���״̬='δ�½���'" Then
        ctxtDiagnose.Enabled = False
        Cmd����ģ��(1).Enabled = False
        Else
        ctxtDiagnose.Enabled = True
        Cmd����ģ��(1).Enabled = True
        End If
    Else
        ctxtDiagnose.Enabled = True
        Cmd����ģ��(1).Enabled = True

    End If
    '��˫���͵���Ч��һ������8023�ǡ������ˡ�״̬ʱ����������ۡ���ť���ã�8023�������ڸ���ʱ�µģ�  2015-10-26 Ĳ��
    
    If mobjQueryResult.Filter = "���״̬='������'" Then
        If resql("��������") = "8023����" Or resql("��������") = "��˲���" Or resql("��������") = "��˲�YK" Then
'        If ctlb������.Buttons(5).Visible = True And (resql("��������") = "8023����" Or resql("��������") = "��˲���" Or resql("��������") = "��˲���YK") Then
'        If resql("��������") = "8023����" Then
        ctxtDiagnose.Enabled = True
        Cmd����ģ��(1).Enabled = True
        ctlb������.Buttons(3).Enabled = False
        Else
        ctxtDiagnose.Enabled = True
'        ctxtDiagnose.Enabled = False
        ctlb������.Buttons(3).Enabled = False
        End If
    End If
     Text�������.Visible = False  '�������������
      'ֻ���û����οƲ���״̬Ҫ��δ�½������ʾ���潨�鰴ť  2016-4-20 by Ĳ��
     If (um�û���� = "8827" Or um�û���� = "0001") And mobjQueryResult.Filter = "���״̬='δ�½���'" Then
     comd�������.Visible = True
     ctlb������.Buttons(13).Enabled = True
     End If
    '���ο�һ���˻ظ��˵İ�ť����Ϊ��ʱ����������µ��Ѹ��ˣ���Ҫ�����˻ش�����  2016-5-19 by Ĳ��
    If (um�û���� = "8827" Or um�û���� = "0001") And coptType(3).Value = True Then
    Com�˻�.Visible = True
    End If
End Sub

'2012-06-21 �ڵ��
'�����鿴���ҽ��ۣ���˫���Ĳ�����ȫ��ͬ
'��˫����������������������Ӷ���Щ�ߣ�
Private Sub cgrdInfo_Click()
    indX = cgrdInfo.MouseRow
    indY = cgrdInfo.MouseCol
    If indX < 0 Or indY < 0 Then
        Exit Sub
    ElseIf indX >= 0 And indX < cgrdInfo.rows And indY >= 0 And indY < cgrdInfo.cols Then
        cgrdDept.Clear
        cgrdItem.Clear
        Text�������.Text = ""  'ÿ�ε����Ա��¼ʱ���ԭ���Ľ������  2015-11-5
        pstrPerson = cgrdInfo.TextMatrix(indX, 0)
        sub�г������ҽ���
        '���ܣ������ʼ����ʱ������
        '���ߣ�����
        'ʱ�䣺2012-06-01
''        cbox����ģ��.Visible = True
        cchk��׼(1).Visible = True
        cchk��׼(0).Visible = True
        'ʱ�䣺2012-06-01
        
''�Զ����벻�ϸ���Ŀ  2016-5-13 by Ĳ��
'        If coptType(1).Value = True Then
'            sub�Զ����벻�ϸ���Ŀ pstrPerson
'        End If

        
        '2012-08-22 �ڵ�� ��
        '�����趨��ctxtDetpConclusion.text������ʾδ�����Ŀ��Һ�δ�����
        'ctxtDetpConclusion.Text = cgrdDept.TextMatrix(cgrdDept.Row, 1)      '�̶���1��Ϊ���ҽ���
        ctxtDetpConclusion.Text = subδ���������������Ŀ(pstrPerson)
        '2012-08-22 �ڵ�� ��
    
    End If
    '����ʱ������ģ��Ĭ���С����飺������       2015-10-19
    If ctxtDiagnose = "" Then
       ctxtDiagnose.Text = "���飺"
    End If
    
    'ѡ���˺󡰸�����Ϣ����ť������   2015-10-22
    Command2.Enabled = True
    Command2.Visible = True
    
    
    '����ʱ8023���������Ҫ��Ҫ�õ�����˫��Ч��һ����      2015-10-26
    Dim resql As Object
    mstrϵͳ��� = cgrdInfo.TextMatrix(cgrdInfo.Row, mcolIndex("ϵͳ���"))
    Set resql = dafuncGetData("select �������� From ְҵ�����_�����Ա������Ϣ�� where ϵͳ���='" & mstrϵͳ��� & "'")
'    If resql("��������") = "8023����"  Then
    If resql("��������") = "8023����" Or resql("��������") = "��˲���" Or resql("��������") = "��˲���YK" Then
        If mobjQueryResult.Filter = "���״̬='δ�½���'" Then
        ctxtDiagnose.Enabled = False
        Cmd����ģ��(1).Enabled = False
        Else
        ctxtDiagnose.Enabled = True
        Cmd����ģ��(1).Enabled = True
        End If
    Else
        ctxtDiagnose.Enabled = True
        Cmd����ģ��(1).Enabled = True

    End If
    '��˫���͵���Ч��һ������8023�ǡ������ˡ�״̬ʱ����������ۡ���ť���ã�8023�������ڸ���ʱ�µģ�  2015-10-26
    
    If mobjQueryResult.Filter = "���״̬='������'" Then
        If resql("��������") = "8023����" Or resql("��������") = "��˲���" Or resql("��������") = "��˲�YK" Then
'        If ctlb������.Buttons(5).Visible = True And (resql("��������") = "8023����" Or resql("��������") = "��˲���" Or resql("��������") = "��˲�YK") Then
'        If resql("��������") = "8023����" Then
        ctxtDiagnose.Enabled = True
        Cmd����ģ��(1).Enabled = True
        ctlb������.Buttons(3).Enabled = False
        Else
        ctxtDiagnose.Enabled = True
'        ctxtDiagnose.Enabled = False
        ctlb������.Buttons(3).Enabled = False
        End If
     End If
     Text�������.Visible = False  '�������������
     'ֻ���û����ο�(���Ϊ8827)����״̬Ҫ��δ�½������ʾ���潨�鰴ť  2016-4-20 by Ĳ��
     If (um�û���� = "8827" Or um�û���� = "0001") And mobjQueryResult.Filter = "���״̬='δ�½���'" Then
     comd�������.Visible = True
     ctlb������.Buttons(13).Enabled = True    'Ԥ��
     End If
    '���ο�һ���˻ظ��˵İ�ť����Ϊ��ʱ����������µ��Ѹ��ˣ���Ҫ�����˻ش�����  2016-5-19 by Ĳ��
    If (um�û���� = "8827" Or um�û���� = "0001") And coptType(3).Value = True Then
    Com�˻�.Visible = True
    End If
End Sub

'2012-04-12 �ڵ��
'����BackSpace����ascii��Ϊ8���ͽ�ѡ���д��б���ɾ����
'ɾ�������������һ��ѡ������ǰɾ������Ǳ���ģ�
'��Ϊÿ��ɾһ�У������еı�Ż����·���
Private Sub cgrdInfo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        Dim i, selRow As Integer
        For i = cgrdInfo.SelectedRows - 1 To 0 Step -1
            selRow = cgrdInfo.SelectedRow(i)
            cgrdInfo.RemoveItem (selRow)
        Next
        '�޸��ˣ����� 2012.12.05
        'bug�ţ�0000070
        '˵�����Ƴ���Ա��������������ı�������   ����
        cgrdDept.rows = 1
        ctxtConclusion.Text = ""
        '2012.12.05  ����
    End If
End Sub

Private Sub cgrdItem_DblClick()
If Left(cgrdItem.TextMatrix(cgrdItem.RowSel, 0), 2) = "08" Then
formdct.Label2.Caption = cgrdInfo.TextMatrix(cgrdInfo.RowSel, 0)
formdct.Show 1
End If


End Sub

'�������еĽ���ģ�� �ɽ���ѡ��
Private Sub Cmd����ģ��_Click(Index As Integer)
    If Index = 0 Then
         frmConclusion.lobj���� = "����ģ��"
    Else
         frmConclusion.lobj���� = "���ģ��"
     End If
    frmConclusion.lobj���ÿ��� = Me.name
'    frmConclusion.lobj���� = priDeptName
'    frmConclusion.lobj���ұ�� = priDeptNo
    frmConclusion.lobjҽ����� = um�û����
    frmConclusion.lobjʱ�� = Now
    frmConclusion.Show
End Sub


Private Sub cmdѡ�񸴲���Ŀ_Click()
    frmSelectItem.Show 1
End Sub

'���潨��ͽ���   2016-4-20 by Ĳ��
Private Sub comd�������_Click()
        Dim teststyle As String
        Dim i, selRow As Integer
        selRow = cgrdInfo.SelectedRow(i)
        teststyle = cgrdInfo.TextMatrix(selRow, 8)
        Dim sjjl As String
        Dim SyNo As String
        sjjl = ctxtConclusion.Text & "_00_" & Trim(ctxtDiagnose.Text)
        SyNo = cgrdInfo.TextMatrix(selRow, 0)
        Dim lob As Object
        Set lob = dafuncGetData("select ���� from ϵͳ����_Ա��������Ϣ�� where ���='" & um�û���� & "'")
        dafuncGetData "update ְҵ�����_���ҽ��۱� set ���ֽ���='" & sjjl & "',ҽ�����='" & lob("����") & "',��������=getdate() where ϵͳ���='" & SyNo & "' and ����=(select ��� from ϵͳ����_�ֵ�_�ֵ����ݱ� where ����='���ս���¼��' and ����='ְҵ�����_����');"
        dafuncGetData ("update ְҵ�����_��������Ϣ�� set ������='" & ctxtConclusion.Text & "', ��Ϻʹ������='" & Trim(ctxtDiagnose.Text) & "',���״̬='5'  where ϵͳ���='" & SyNo & "'")
        ccmdQuery_Click
        cgrdDept.rows = 1
        comd�������.Visible = False
End Sub

Private Sub Command1_Click()
If cgrdItem.TextMatrix(cgrdItem.RowSel, 0) <> "" Then
       'ȥ���ַ�����ߵ����֣� 2015-10-19 Ĳ��
         Dim stmp As String
         stmp = cgrdDept.TextMatrix(cgrdDept.RowSel, 0)
         Dim slen As Integer
         slen = Len(stmp)
         Dim clen As Integer
         clen = slen - 2
         Dim ctemp As String
         ctemp = Right(stmp, clen)
    If ctxtConclusion.Text = "" Then
'        ctxtConclusion.Text = ctemp + "-" + cgrdItem.TextMatrix(cgrdItem.RowSel, 1) + "-" + cgrdItem.TextMatrix(cgrdItem.RowSel, 2) + "��"
        ctxtConclusion.Text = ctemp + "-" + cgrdItem.TextMatrix(cgrdItem.RowSel, 2) + "��" 'ֻҪ�������Ҫǰ�����Ŀ 2016-4-18 by Ĳ��
    Else
'        ctxtConclusion.Text = ctxtConclusion.Text + ctemp + "-" + cgrdItem.TextMatrix(cgrdItem.RowSel, 1) + "-" + cgrdItem.TextMatrix(cgrdItem.RowSel, 2) + "��"
        ctxtConclusion.Text = ctxtConclusion.Text + cgrdItem.TextMatrix(cgrdItem.RowSel, 2) + "��" 'ֻҪ�������Ҫǰ�����Ŀ  2016-4-18 by Ĳ��
    End If
Else
End If
End Sub


'����һ���ܲ鿴����������Ϣ�İ�ť 2015-10-22��
Private Sub Command2_Click()
If cgrdInfo.Row < 1 Then
            MsgBox "û����Ҫ�޸ĵļ�¼��", vbOKOnly + vbExclamation, "ϵͳ��ʾ"
            Exit Sub
        End If

'        ���ʼǺ� = 1
        mstrϵͳ��� = cgrdInfo.TextMatrix(cgrdInfo.Row, mcolIndex("ϵͳ���"))
        'frmCareerHstRegt.ctxtsysno = cgrdMain.TextMatrix(cgrdMain.Row, mcolIndex("ϵͳ���"))
        frmInfromation.Show 1, Me
End Sub


'�������а�ť 2016-5-30 by Ĳ��
Private Sub Com����_Click()
frmAssessDeformity.Show 1
End Sub

Private Sub Com�˻�_Click()
If MsgBox("��ȷ��Ҫ�˻ظ�����¼����������", vbYesNo + vbQuestion + vbDefaultButton2, "ϵͳ��ʾ") = vbYes Then
            pobjҵ�����.funcд�뵥�˵�ǰ���״̬ pstrPerson, 5 '"���½���"
            ccmdQuery_Click
            cgrdDept.rows = 1
            MsgBox "�ѳɹ��˻ش����ˣ���鿴��"
End If
End Sub

'2012-07-03 �ڵ��
'cgrdinfo���������״̬�ı���ı䣬ͬʱ������Ȩ��Ҳ�����仯��
Private Sub coptType_Click(Index As Integer)
    cgrdDept.rows = 1
    ctxtConclusion.Text = ""
    ctxtReview.Text = ""
    ctxtReviewItem.Text = ""
    '2013-1-17 ������
    'ֱ�Ӵ�����ѯ�¼�
'    If Not mobjQueryResult Is Nothing Then
'        If coptType(Index).Caption <> "�����" Then
'            mobjQueryResult.Filter = "���״̬='" & coptType(Index).Caption & "'"
'
'        Else
'            mobjQueryResult.Filter = "���״̬='�����' or ���״̬='δ¼���ܼ��߸�����Ϣ'"
'        End If
'    End If
'    Set cgrdInfo.DataSource = mobjQueryResult
    ccmdQuery_Click
    '2013-1-17 ������
    If cgrdInfo.rows > 1 Then
        cgrdInfo.AutoSize 0, cgrdInfo.cols - 1, 0, 0
        cgrdInfo.ExplorerBar = flexExSort
        cgrdInfo.DataMode = flexDMFree
        cgrdInfo.Col = 0
        cgrdInfo.Sort = flexSortGenericDescending
    End If
        
    ctlb������.Buttons(6).Visible = False
    ctlb������.Buttons(7).Visible = False
    ctlb������.Buttons(8).Visible = False
    ctlb������.Buttons(9).Visible = False
    ctlb������.Buttons(10).Visible = False
    ctlb������.Buttons(3).Enabled = (coptType(Index).Caption = "δ�½���") Or (coptType(Index).Caption = "���½���") Or (coptType(Index).Caption = "������")
    ctlb������.Buttons(5).Enabled = (coptType(Index).Caption = "���½���") Or (coptType(Index).Caption = "������") Or (coptType(Index).Caption = "������")
    ctlb������.Buttons(7).Enabled = (coptType(Index).Caption = "���½���") Or (coptType(Index).Caption = "�Ѹ���") Or (coptType(Index).Caption = "�ѷ�����") Or (coptType(Index).Caption = "������")
    ctlb������.Buttons(9).Enabled = (coptType(Index).Caption = "�Ѹ���") Or (coptType(Index).Caption = "�ѷ�����") Or (coptType(Index).Caption = "������")
    ctlb������.Buttons(11).Enabled = (coptType(Index).Caption = "�Ѹ���") Or (coptType(Index).Caption = "�ѷ�����") Or (coptType(Index).Caption = "������")
     
    '������״̬���Դ�ӡ��챨�棻���ǣ�2012-10-24
    '����״̬����Ԥ�����档 �޸��ˣ����� 2013.03.01
'    ctlb������.Buttons(13).Enabled = (coptType(Index).Caption = "�Ѹ���") Or (coptType(Index).Caption = "�ѷ�����") Or (coptType(Index).Caption = "������")

    '��δ�½��ۣ������ˣ��Ѹ��� ������Ԥ��  Ĳ��  2015-11-9
    ctlb������.Buttons(13).Enabled = (coptType(Index).Caption = "�Ѹ���") Or (coptType(Index).Caption = "�ѷ�����") Or (coptType(Index).Caption = "������") Or (coptType(Index).Caption = "δ�½���") Or (coptType(Index).Caption = "������") Or (coptType(Index).Caption = "δ�½���")
    '�޸��ˣ����� 2012.12.05
    'bug�ţ�0000062
    '˵�������������Ϊ������ʱ��mstrȨ�ޱ�־��false   ����
    If ctlb������.Buttons(3).Enabled = True Then
        mstrȨ�ޱ�־ = True
    Else
        mstrȨ�ޱ�־ = False
    End If
    '2012.12.05    ����
    If coptType(2).Value = True Then
        ctlb������.Buttons(15).Visible = True
        ctlb������.Buttons(16).Visible = True
    Else
        ctlb������.Buttons(15).Visible = False
        ctlb������.Buttons(16).Visible = False
    End If
    'δ�½���(6)��������(2)���Ѹ���(3)������ ȡ������ ��ť������
    If coptType(2).Value = True Or coptType(6).Value = True Then
'    If coptType(2).Value = True Or coptType(6).Value = True Or coptType(3).Value = True Then
        ctlb������.Buttons(15).Visible = True
    End If
'δ�½���Ҳ����Ԥ��  2016-7-25 by Ĳ��
    If coptType(0).Value = True Then
        ctlb������.Buttons(13).Enabled = True
    End If
    cchk��׼(0).Value = True
End Sub





'2012-04-10 �ڵ��
Private Sub Form_Load()
    Dim lcol��������ť As New Collection '�������ϵİ�ť��ʼ�����ϡ�
    
    On Error GoTo errHandler
    
    '��������Ѿ���ʼ�������ٽ��г�ʼ����
    If mblnInUse Then Exit Sub
    mstrȨ�ޱ�־ = True     'Ĭ����Ȩ��
    '��ʾ���ȡ�
    frmProcess.proPercent.max = 8
    frmProcess.Label1.Caption = "���ڳ�ʼ�����棬��ȴ�..."
    frmProcess.proPercent.Value = 1
    frmProcess.Show
    DoEvents
    Me.Enabled = False
    MousePointer = 11
    '���ô�������ʹ�õı�־��
    mblnInUse = True
    
    '��ʼ������ͨ�ö���ÿ�������Ӧһ������ͨ�ö���ʵ�������ɻ��ã��м��мǣ���
    Set mobjGUI = New cls����ͨ�ö���
    
    '���ù�����������Ҫ�ĸ��ְ�ť��
    With lcol��������ť
        .Add "��ս���(&C)110"    '1
        .Add "�Ƴ���Ա(&M)104"    '2
        .Add "�������(&S)101"    '3
        .Add "|"    '4
        .Add "����(&F)109"    '5
        .Add "|"    '6
        .Add "Ԥ������(&Y)108"    '7
        .Add "|"    '8
        .Add "��ӡ����(&P)107"    '9
        .Add "|"    '10
        .Add "����ΪPDF(&O)102"    '11
        .Add "|"    '12
       
      '2015/1/16 �����Ԥ�������Ԥ���ڱ����������
       .Add "Ԥ��(&V)102"    '13
        .Add "|"    '14
        .Add "ȡ������(&E)111"    '15
        .Add "|"    '16
        .Add "�˳�"
    End With
    frmProcess.proPercent.Value = 2
    DoEvents
    'Ϊ��Ҫͨ������ͨ�ö�����Ƶĸ��ֿؼ�����ʼֵ��
    With mobjGUI
        Set .Form = Me
        Set .c������ = ctlb������
    End With
    
    resql = "0"
    
    '���ý���ͨ�ö����ṩ�ķ������Խ���ؼ����г�ʼ����
    mobjGUI.subInitialize lcol��������ť, ""
    frmProcess.proPercent.Value = 3
    DoEvents
    '������ʼ��
    Set pobj��� = CreateObject("ְҵ������.clsMedicalExam")
    Set pobj����ģ�� = CreateObject("ְҵ������.clsMedicalExamTemplate")
    Set pobj�����ҵ�� = CreateObject("ְҵ�������¼��.clsCommon")
    Set pobj���� = pobjDict.Fetch("ְҵ���������ֵ�")
    pstrPerson = ""
    frmProcess.proPercent.Value = 4
    DoEvents
    '����Ȩ�޿���
    '���Ƶ�Ȩ�޽����ڣ����ˡ�������ۡ���ӡ���桢����ΪPDF��
    Dim lobjTmp As Object
    Set lobjTmp = CreateObject("ְҵ������.clsPermissionConfigure")
    If lobjTmp.func���Ҳ���Ȩ��(um�û����, "ְҵ�����_���ս���¼��_�������") = False Then
        ctlb������.Buttons(3).Visible = False
    End If
    If lobjTmp.func���Ҳ���Ȩ��(um�û����, "ְҵ�����_���ս���¼��_����ͨ��") = False Then
        ctlb������.Buttons(5).Visible = False
        ctlb������.Buttons(6).Visible = False
    End If
    If lobjTmp.func���Ҳ���Ȩ��(um�û����, "ְҵ�����_���ս���¼��_��ӡ����") = False Then
        ctlb������.Buttons(9).Visible = False
    End If
    
    'If lobjTmp.func���Ҳ���Ȩ��(um�û����, "ְҵ�����_���ս���¼��_���PDF") = False Then
        ctlb������.Buttons(10).Visible = False 'ע�⣺���PDFȨ��δд����ò�����Ϣ��
        ctlb������.Buttons(11).Visible = False
    'End If
    Set lobjTmp = Nothing
    frmProcess.proPercent.Value = 5
    DoEvents
    '����ؼ���ʼ��
'    llabDoctor = llabDoctor & um�û���
    fraFinal.Caption = fraFinal.Caption & "(ҽʦ������" & um�û��� & ")"
    cdtpDateTo.Value = DateAdd("m", -1, Now())
    cdtpDateFrom.Value = Now
    cdtpConclusion.Value = Now
    coptType(1).Value = True
    With cgrdInfo
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "��������"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "����"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "�Ա�"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "����"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "��λ����"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "�������"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "�������"
        '.Cols = .Cols + 1: .TextMatrix(0, .Cols-1) = "��������"
        .AutoSize 0, .cols - 1, 0, 0
        .SelectionMode = flexSelectionListBox
    End With
    
    If coptType(0).Value = True Or coptType(1).Value = True Then
        ctlb������.Buttons(15).Visible = False
        ctlb������.Buttons(16).Visible = False
    Else
        ctlb������.Buttons(15).Visible = True
        ctlb������.Buttons(16).Visible = True
    End If
    '2012-06-21 �ڵ�� ��
    '��ӹ�������ʼ��״̬����Ҫ�������̿��ơ�
    '��ʼ���ܸ��ˣ�ѡ��ĳ�������Ա�ж��Ƿ񸴺ˡ����˺��ӡ������ΪPDF��
    ctlb������.Buttons(5).Enabled = False   '����
    ctlb������.Buttons(7).Enabled = False   'Ԥ������
    ctlb������.Buttons(9).Enabled = False   '��ӡ����
    ctlb������.Buttons(11).Enabled = False  '����ΪPDF
    ctlb������.Buttons(13).Visible = True  '�ſ�����Ԥ��  2015-11-6 by lanchao
    '2012-06-21 �ڵ�� ��
    frmProcess.proPercent.Value = 6
    DoEvents
    '����ģ���б��ʼ��
    sub��������ģ��
    frmProcess.proPercent.Value = 7
    DoEvents
    '���ܣ����ؽ���ģ��
    '���ߣ�����
    'ʱ�䣺2012-05-31
    Dim lobj����ģ�� As Object
    Dim lobj���۱� As Object
    Dim i As Integer
    Dim sql As String
    
    sql = "select * from ϵͳ����_�ֵ�_������ģ��� where ���ұ��='16'"
    
    Set lobj����ģ�� = CreateObject("ְҵ������.clsConclusionSet")
    Set lobj���۱� = lobj����ģ��.func��ȡ�������ս���ģ��(sql)
    If lobj���۱�.RecordCount > 0 Then
        For i = 1 To lobj���۱�.RecordCount
            cbox����ģ��.AddItem lobj���۱�("����ģ��")
            lobj���۱�.MoveNext
        Next i
    End If
    frmProcess.proPercent.Value = 8
    DoEvents
    'ʱ�䣺2012-05-31
    
    '2012-06-21 �ڵ�� ��
    '��ʼ�����״̬
    mstrState = ""
    '2012-06-21 �ڵ�� ��
    
    '2012-07-03 �ڵ�� ��
    '����ؼ���ʼ��
    cchk��׼(0).Value = True
    '2012-07-03 �ڵ�� ��
    
    '2012-08-22 �ڵ�� ��
    '��ӿ��ұ���
    Set pobjDept = pobjDict.Fetch("ְҵ���������ֵ�")
    '2012-08-22 �ڵ�� ��
    Unload frmProcess
'    Exit Sub
errHandler:
    Me.Enabled = True
    MousePointer = 0
    If Err.Number <> 0 Then
        Dim lstrError As String
        lstrError = func������(Err.Number, Err.Description)
        sfsub������ "ְҵ������", "frmFinalConclusion", "Form_Load", 6666, lstrError, False
        Exit Sub
        Resume
    End If
End Sub

'2012-04-10 �ڵ��
'�˳�����ʱ����ղ��ֱ���
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    mblnInUse = False
    Set mobjGUI = Nothing
    Set mobjQueryResult = Nothing
End Sub

'��������Ӧ�ֱ��ʴ�С
'2012-10-19 ������
Private Sub Form_Resize()
    On Error Resume Next
    Picture1.Width = Me.ScaleWidth - Picture1.Left
    Picture1.Height = Me.ScaleHeight - Picture1.Top
    Frame1.Width = Picture1.Width - Frame1.Left
    Frame1.Height = Picture1.Height - Frame1.Top
    ctlb������.Width = Frame1.Width - ctlb������.Left
    fraFinal.Left = Frame1.Width - fraFinal.Width - 80
    fraPerson.Width = Frame1.Width - fraPerson.Left - fraFinal.Width - 160
    cgrdInfo.Width = fraPerson.Width - cgrdInfo.Left * 2
    Label5.Width = cgrdInfo.Width
    
    fraDeptItem.Width = Frame1.Width - fraDeptItem.Left - 80
    cgrdDept.Width = fraDeptItem.Width * 2 / 3
'    ctxtDetpConclusion.Width = cgrdDept.Width
    ctxtDetpConclusion.Width = cgrdDept.Width - 3000   '��ctxtDetpConclusion�ı����ȵ�С  2015-11-5
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

'2012-04-12 �ڵ��
'���湤������ť�����趨
Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = True
    
    Dim lcolID As New Collection
    Dim lobj������� As Object
    Dim lstrStatus As String '��ǰ���״̬
    Dim mcolIndex As New Collection
     
    lcolID.Add pstrPerson
    Set lobj������� = CreateObject("ְҵ������.clsMedicalExam")
    lobj�������.ϵͳ��� = pstrPerson
    
    Select Case Operate
    Case "��ս���"
        subClear
    Case "�Ƴ���Ա"
        cgrdInfo_KeyPress (8)
        '�޸��ˣ����� 2012.12.05
        'bug�ţ�0000070
        '˵�����˴��Ƴ���Ա�Ժ�������������ѯһ�Σ�������ʾ���ݡ�    ����
        Exit Sub
        '2012.12.05         ����
    '2012-06-21 �ڵ�� ��
    '��Ӹ���ͨ������
    Case "����"
        pobjҵ�����.funcд�뵥�˵�ǰ���״̬ pstrPerson, 6 '"�Ѹ���"
        ctxtConclusion.Text = ""
        ctxtDiagnose = ""    '2015-10-16
        cgrdDept.rows = 1
    '2012-06-21 �ڵ�� ��
    Case "Ԥ������"
       
        pobjҵ�����.Sub��ӡ���� "ְҵ�����_���ս���", lcolID, False, True, False
    Case "��ӡ����"
        pobjҵ�����.Sub��ӡ���� "ְҵ�����_���ս���", lcolID, True, False, False
        '2012-07-03 �ڵ�� ��
        '�������״̬
        pobjҵ�����.funcд�뵥�˵�ǰ���״̬ pstrPerson, 7 '"�ѷ�����"
        '2012-07-03 �ڵ�� ��
    '2012-05-30 ��¶
    Case "����ΪPDF"
        pobjҵ�����.Sub��ӡ���� "ְҵ�����_���ս���", lcolID, True, False, True
        '2012-07-03 �ڵ�� ��
        '�������״̬
        pobjҵ�����.funcд�뵥�˵�ǰ���״̬ pstrPerson, 7 '"�ѷ�����"
        '2012-07-03 �ڵ�� ��
    '2012-05-30
    Case "�������"
    
    '8023,��ˣ��˿�ȵ���������������ۺ󲻵������ˣ����ǵ����е�δ�½���  2016-4-20 by Ĳ����
    Dim resql As Object
'    mstrϵͳ��� = cgrdInfo.TextMatrix(cgrdInfo.Row, mcolIndex("ϵͳ���"))
    Set resql = dafuncGetData("select �������� From ְҵ�����_�����Ա������Ϣ�� where ϵͳ���='" & mstrϵͳ��� & "'")
'    If resql("��������") = "8023����"  Then
    If resql("��������") = "8023����" Or resql("��������") = "��˲���" Or resql("��������") = "��˲���YK" Then
            subSaveConclusion
        '2012-07-03 �ڵ�� ��
        '�������״̬
        If cchk��׼(0).Value Then
            pobjҵ�����.funcд�뵥�˵�ǰ���״̬ pstrPerson, 11 '��״̬11Ϊ8023����ˣ��˿����е�"δ�½���"
        Else
            pobjҵ�����.funcд�뵥�˵�ǰ���״̬ pstrPerson, 11
            pobjҵ�����.funcд�븴�����Ϣ pstrPerson
        End If
        cgrdDept.rows = 1
    Else
    '2016-4-20 ��
    
        subSaveConclusion
        '2012-07-03 �ڵ�� ��
        '�������״̬
        If cchk��׼(0).Value Then
            pobjҵ�����.funcд�뵥�˵�ǰ���״̬ pstrPerson, 5 '"���½���"
        Else
            pobjҵ�����.funcд�뵥�˵�ǰ���״̬ pstrPerson, 5 '��ǰΪ8"������".�����������û��ϵ
            pobjҵ�����.funcд�븴�����Ϣ pstrPerson
        End If
        cgrdDept.rows = 1
        '2012-07-03 �ڵ�� ��
    End If
        
        
        
    '2012-08-20 �ڵ�� ��
    '���wordģ�幦��
    Case "word����"
        With cgrdInfo
            If .Row < 1 Or .Row > .rows - 1 Then Exit Sub
            mstr�������� = .TextMatrix(indX, 7)
'            mstr�������� = .TextMatrix(.Row, mcolIndex("������"))
        End With
    
           '���ߣ������ ʱ��2013-1-9 ��
             '��ʾ���ȡ�
            frmProcess.proPercent.max = 4
            frmProcess.Label1.Caption = "���ڼ��أ���ȴ�..."
            frmProcess.proPercent.Value = 0
            frmProcess.Show 0, Me
            DoEvents
         '���ߣ������ ʱ��2013-1-9 ��
            
        '��ȡ���ϵͳ���
        If cgrdInfo.SelectedRow(0) = -1 Then Exit Sub
        If coptType(1).Value = True Or coptType(2).Value = True Then
            sub�༭word�ĵ� Me, pstrPerson, mstr��������, True
        Else
            sub�༭word�ĵ� Me, pstrPerson, mstr��������, False
        End If
          Unload frmProcess
        
               
        '2012-08-23 �ڵ�� ��
        '�������״̬��word�Ĵ������ã���Ҫ�����ݿ⹦�����֮���ⲽ������ƣ�
        If pstrFilename = "" Then Exit Sub
        If coptType(3).Value = True Then pobjҵ�����.funcд�뵥�˵�ǰ���״̬ pstrPerson, 7   '"�ѷ�����"
        '2012-08-23 �ڵ�� ��
    '2012-08-20 �ڵ�� ��
    
    
    
    
    '��ʱ�Ȱ�������أ��ᱨ��Ԥ���ڱ�����������С�
    
    Case "Ԥ��"
     subPrint True  '�Ƿ�Ԥ��
    
     
    
    
    
    Case "ȡ������"
        If MsgBox("��ȷ��Ҫȡ�����۲��˻ظ�����¼��", vbYesNo + vbQuestion + vbDefaultButton2, "ϵͳ��ʾ") = vbYes Then
            pobjҵ�����.funcд�뵥�˵�ǰ���״̬ pstrPerson, 4 '"���½���"
            ccmdQuery_Click
            cgrdDept.rows = 1
            MsgBox "�ѳɹ�ȡ�����ۣ��������Ϣ���˻ء�"
        End If
    Case "�˳�"
        '�޸��ˣ�  2012.12.05
        'bug�ţ�0000062
        '˵������Ӻ���     ����
        Dim isSave As Integer
        '2012.12.05    ����
        Set frmFinalConclusion = Nothing
        '�޸��ˣ�  2012.12.05
        'bug�ţ�0000062
        '˵������Ȩ�ޱ�־Ϊtrueʱ���˳���ʾ�Ƿ񱣴�     ����
        If mstrȨ�ޱ�־ = True And cgrdInfo.SelectedRows > 0 Then
'            isSave = MsgBox("�Ƿ񱣴����޸Ľ����", vbYesNoCancel)
'            If isSave = vbCancel Then Exit Sub
             Unload Me
'            If isSave = vbNo Then
'                mobjGUI_BeforeOperate "��ս���", True
'                Exit Sub
'            End If
            If isSave = vbYes Then
                mobjGUI_BeforeOperate "�������", False
                mstrȨ�ޱ�־ = False
                Unload Me
            Else
                Unload Me
            End If
        Else
            Unload Me
        End If
        '2012.12.05    ����
        Cancel = True
    End Select
    
    '2012-07-03 �ڵ�� ��
    '����ÿ�β��������ܸı����״̬�����ԣ�ÿ�β���������²�ѯ�����
    '2012.12.11 ����
    '˵������"<>"���ܴﵽԤ��Ч�����ĳ�"="������
'    If Operate <> "��ս���" Or Operate <> "�Ƴ���Ա" Or Operate <> "Ԥ������" Or Operate <> "�˳�" Then ccmdQuery_Click
    If Operate = "����" Or Operate = "��ӡ����" Or Operate = "����ΪPDF" Or Operate = "�������" Or Operate = "word����" Then ccmdQuery_Click
    '2012.12.11  ����
    '2012-07-03 �ڵ�� ��
    Set lobj������� = Nothing
    
    Exit Sub
errHandler:
    sfsub������ "ְҵ������", "frmFinalConclusion", "mobjGUI_BeforeOperate", Err.Number, Err.Description, False
    Cancel = True
    Exit Sub
    Resume
End Sub

'2012-04-12 �ڵ��
Sub sub��������ģ��()
    Dim i As Integer
    Dim lobjRec As Object
    On Error GoTo errHandler

    '�������������Ͽ���
    Set lobjRec = pobjDict.FetchEx("��������ֵ�")
    Ccmb��������.Clear
    'Ccmb��������.AddItem ""
    For i = 1 To lobjRec.RecordCount
        Ccmb��������.AddItem lobjRec("����")
        Ccmb��������.ItemData(Ccmb��������.NewIndex) = lobjRec("���")
        lobjRec.MoveNext
    Next
    Ccmb��������.ListIndex = 0
   
    Set lobjRec = pobjDict.FetchEx("���������ֵ�")
    ccmb���������.Clear
    'ccmb���������.AddItem ""
    For i = 1 To lobjRec.RecordCount
        ccmb���������.AddItem lobjRec("����")
        ccmb���������.ItemData(ccmb���������.NewIndex) = lobjRec("���")
        lobjRec.MoveNext
    Next
    ccmb���������.ListIndex = 0
   
    If ccmbTemplate.ListCount > 0 Then
        ccmbTemplate.Text = ccmbTemplate.List(0)
        subChangeTemplate
    Else
        ccmb���������_Click
    End If
        
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ������", "frmFinalConclusion", "sub��������ģ��", 6666, lstrError, False
    Exit Sub
    Resume
End Sub

'2012-04-12 �ڵ��
'ѡ������ģ�������б�
Private Sub ccmbTemplate_Click()
    On Error GoTo errHandler
    MousePointer = 11
    subChangeTemplate       'ѡ������
    MousePointer = 0
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ������", "frmFinalConclusion", "ccmbTemplate_Click", 6666, lstrError, False
    Exit Sub
    Resume
End Sub

'2012-04-12 �ڵ��
'�ı�����ģ��ʱ�Ĳ�����
Private Sub subChangeTemplate()
    On Error GoTo errHandler
    
    If pobj���.����.������ <> ccmbTemplate.Text Then
        pobj���.����.������ = ccmbTemplate.Text
        '��������ģ���ȡ���������п��õ���ĸ��
        pobj����ģ��.������ = ccmbTemplate.Text
    End If
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ������", "frmFinalConclusion", "subChangeTemplate", 6666, lstrError, True
    Exit Sub
    Resume
End Sub

'2012-04-12 �ڵ��
Private Sub sub�г������ҽ���()
    On Error GoTo errHandler
    Dim lobjRec As Object
    Dim lstrCon As String
    Dim strArray
    Dim i As Integer
    
    '2012-05-24 �ڵ�� ��
    'ÿ�β�ѯ����յ�ǰ���н����񡢿��ҽ��ۡ����ս��ۡ�����
    cgrdDept.Clear
    cgrdDept.rows = 1
    cgrdItem.Clear
    cgrdItem.rows = 1
    ctxtDetpConclusion.Text = ""
    ctxtConclusion.Text = ""
    ctxtDiagnose.Text = ""

    
    'vsflexgrid�п�Ȱ������Զ������������ͷ����ͷ����������
    cgrdInfo.AutoSize 0, cgrdInfo.cols - 1, 0, 0
    cgrdInfo.ExplorerBar = flexExSort
    cgrdInfo.DataMode = flexDMFree
    
    '2012-05-24 �ڵ�� ��
    
    Set lobjRec = pobj�����ҵ��.func��ȡ�����Ա���ҽ���(pstrPerson)  '���صĿ���Ϊ���
    
    If lobjRec.RecordCount > 0 Then
        Set cgrdDept.DataSource = lobjRec
        With cgrdDept
            lobjRec.MoveFirst
            For i = 1 To lobjRec.RecordCount
                pobj����.Filter = "���=" & lobjRec("����")
                .TextMatrix(i, 0) = lobjRec("����") & " " & pobj����("����")       '�̶���0��Ϊ�������ƣ�datasource����
                If pobj����("����") = "���ս���¼��" Then
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
    lstrCon = pobjҵ�����.func���ؿ��ҽ���(pstrPerson, "���ս���¼��")
    
    ' "_00_" û��ʲô���⺬�壬ֻ����Ϊ�ָ���ۺ��������ķָ���
    strArray = Split(lstrCon, "_00_", -1, vbBinaryCompare)
    If UBound(strArray) = 1 Then
        ctxtConclusion.Text = strArray(0)
        ctxtDiagnose.Text = strArray(1)
    End If
    Set lobjRec = dafuncGetData("select ����ԭ��,������Ŀ from ְҵ�����_��������Ϣ�� where ϵͳ���='" & pstrPerson & "'")
    If Not (lobjRec.BOF Or lobjRec.EOF) Then
        ctxtReview.Text = IIf(IsNull(lobjRec("����ԭ��")), "", lobjRec("����ԭ��"))
        ctxtReviewItem.Text = IIf(IsNull(lobjRec("������Ŀ")), "", lobjRec("������Ŀ"))
        'ע�ͣ������ж�������Զ�������������޸��ж�����  2016-5-16 by Ĳ��
'        If IsNull(lobjRec("����ԭ��")) And IsNull(lobjRec("������Ŀ")) Then
'            cchk��׼(0).Value = True
'        Else
'            cchk��׼(1).Value = True
'        End If
        If Me.cgrdInfo.TextMatrix(Me.cgrdInfo.Row, 12) = "" And Me.cgrdInfo.TextMatrix(Me.cgrdInfo.Row, 13) = "" Then
        cchk��׼(0).Value = True
        Else
        cchk��׼(1).Value = True
        End If
    End If
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ������", "frmFinalConclusion", "sub�г������ҽ���", 6666, lstrError, True
    Exit Sub
    Resume
End Sub

'2012-04-12 �ڵ��
Private Sub sub�г����������������(ByVal paraDeptName As String)
    Set pobjItem = pobj�����ҵ��.func��ȡ�����Ա�����������(pstrPerson, paraDeptName)
    With cgrdItem
        Set .DataSource = pobjItem
        '�޸��ˣ����� 2012.12.12   ����
        'ȡ��������Ϊ0������û�����ݡ�
'        .Col = 0
        '�޸��ˣ����� 2012.12.12   ����
        .Sort = flexSortStringAscending
        .AutoSize 0, .cols - 1, 0, 0
        .SelectionMode = flexSelectionListBox
        .AllowSelection = False
    End With
    '�޸��ˣ����� 2012.12.12     ����
    '�¼��ظ�ִ�У�û�����塣
'    cchkUnfilled_Click
'    cchkAbnormal_Click
    '�޸��ˣ����� 2012.12.12     ����
End Sub

'2012-04-12 �ڵ��
'��ս����Ͽؼ�����
Sub subClear()
    
    '����ؼ���ʼ��
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
    '�޸��ˣ������ 2012-12-10 ��
    'bug�ţ�0000060
    '����ģ�滹ԭ
    cchkTemplate.Value = 0
    sub��������ģ��
     'ʱ�仹ԭΪ��ǰ����ʱ��
    cdtpDateTo.Value = Date
    cchkDate.Value = 0
     '�������Ż�ԭ
    ctxtBarCode.Text = ""
    cchkBarCode.Value = 0
    '��λ���ƻ�ԭ
    ctxtCompanyName.Text = ""
    cchkCompanyName.Value = 0
      '�޸��ˣ������ 2012-12-10 ��
End Sub

'2012-04-12 �ڵ��
'������ۺ���������
'�������ս���¼�롱����һ�����ң����ۺ���������Ϊһ���ַ���������һ���ܵĽ��۴��롰���ҽ��۱���
Sub subSaveConclusion()
    Dim i, selRow As Integer
    For i = 0 To cgrdInfo.SelectedRows - 1
        selRow = cgrdInfo.SelectedRow(i)
        ' "_00_" û��ʲô���⺬�壬ֻ����Ϊ�ָ���ۺ��������ķָ���
        pobjҵ�����.sub������д������ cgrdInfo.TextMatrix(selRow, 0), "���ս���¼��", ctxtConclusion.Text & "_00_" & Trim(ctxtDiagnose.Text), um�û����, Trim(ctxtReview.Text), Trim(ctxtReviewItem.Text)
    Next
    If cgrdInfo.SelectedRows > 0 Then
        MsgBox ("����ɹ���")
        '���������Ϣ
        ctxtConclusion.Text = ""
        ctxtDiagnose = ""
    End If
End Sub

'2012-06-25 �ڵ��
'��Ӻ����Զ��ڽ���text�У�����Զ����벻�ϸ������Ŀ
Sub sub�Զ����벻�ϸ���Ŀ(ByVal paraSysNo As String)
    Dim lobjRec As Object
    Dim strSQL As String
    Dim i As Integer
    strSQL = "select distinct b.���� from ְҵ�����_�������ͼ a,ְҵ�����_�����Ŀ���ñ� b where a.ϵͳ���='" & paraSysNo & "' and a.�����Ŀ=b.���� and a.�������='���ϸ�'"
    Set lobjRec = dafuncGetData(strSQL)
    If lobjRec.RecordCount > 0 Then
        ctxtConclusion.Text = "�ô���첻�ϸ���Ŀ��"
        lobjRec.MoveFirst
        For i = 1 To lobjRec.RecordCount
            If i <> 1 Then
                ctxtConclusion.Text = ctxtConclusion.Text & "��" & lobjRec("����")
            Else
                ctxtConclusion.Text = ctxtConclusion.Text & lobjRec("����")
            End If
            lobjRec.MoveNext
        Next
        ctxtConclusion.Text = ctxtConclusion.Text & "��" & vbCrLf
    End If
End Sub

'2012-08-22 �ڵ��
'�ҳ�����δ������������Ŀ
Private Function subδ���������������Ŀ(ByVal paraSysNo As String) As String 'vbcrlf
    Dim strSQL As String
    Dim lobjRec As Object
    Dim i As Integer, j As Integer
    Dim resultStrDept, resultStrItem As String
    
    strSQL = "select distinct a.�����Ŀ,b.���� from ְҵ�����_�������ͼ a, ְҵ�����_�����Ŀ���ñ� b where a.ϵͳ���='" & paraSysNo & "' and (a.�����='' or a.����� is null) and a.�����Ŀ=b.����"
    Set lobjRec = dafuncGetData(strSQL)
    If lobjRec.RecordCount > 0 Then
        lobjRec.MoveFirst
        resultStrItem = "δ�����Ŀ�У�"
        For i = 0 To lobjRec.RecordCount - 1
'            resultStrItem = IIf(i = 0, resultStrItem & lobjRec("�����Ŀ") & lobjRec("����"), resultStrItem & "��" & lobjRec("�����Ŀ") & lobjRec("����"))
            resultStrItem = IIf(i = 0, resultStrItem & lobjRec("����"), resultStrItem & "��" & lobjRec("����"))
            lobjRec.MoveNext
        Next i
        resultStrItem = resultStrItem & "��"
    End If
    subδ���������������Ŀ = resultStrItem
End Function
'��WORD����ã����ڱ���WORD�����ݿ�
Public Sub subSave(ByVal paraFile As String, ByVal paraNo As Integer, ByVal paraϵͳ��� As String)
    subSaveDoc paraFile, paraNo, paraϵͳ���
End Sub
'��ӡ���� 2015-11-6 by lanchao update print
Private Sub subPrintold(ByVal paraԤ�� As Boolean)
    Dim i As Integer
    Dim lobj���� As Object
    Dim lcolSysNo As Collection
    On Error GoTo errHandler
    Set lobj���� = CreateObject("ְҵ������.cls����")
'    sum = 0
    With cgrdInfo
        Set lcolSysNo = New Collection
'        For i = 1 To .rows
'             If .Cell(flexcpChecked, i, 0) = flexChecked Then
                    lcolSysNo.Add .TextMatrix(.Row, 0)
                    If Right(lcolSysNo(1), 1) = "F" Then
                    '����ֻ�ܿ�ְҵ����ģ�
                        lobj����.Sub��ӡ���� "ְҵ�������_" & .TextMatrix(.Row, 5) & "F", lcolSysNo, paraԤ��
                    Else
                        lobj����.Sub��ӡ���� "ְҵ�������_" & .TextMatrix(.Row, 5), lcolSysNo, paraԤ��
                    End If
                    If paraԤ�� = False Then
                        dafuncGetData "update ְҵ�����_��������Ϣ�� set ���״̬='7' where ϵͳ���='" & Trim(.TextMatrix(i, 0)) & "'"
                        .RowHidden(i) = True
                    End If
'            End If
'        Next i
'         If lcolSysNo.Count < 1 And .rows > 1 Then
'            MsgBox "�빴ѡҪ��ӡ��Ԥ��������", vbInformation, "ϵͳ��ʾ"
'            Exit Sub
'        End If
   
    End With
errHandler:
    
End Sub
            
'��ӡ����
'2015-11-9 Ĳ��
Private Sub subPrint(ByVal paraԤ�� As Boolean)
    Dim sql As String
    Dim lobjet As Object
    Dim mstr�������� As String
    Dim i As Integer
    Dim lobj���� As Object
    Dim lcolSysNo As Collection
    On Error GoTo errHandler
  
    Set lobj���� = CreateObject("ְҵ������.cls����")
'    sum = 0
    With cgrdInfo
'        For i = 1 To .rows - 1
'            If .Cell(flexcpChecked, i, 0) = flexChecked Then
             Set lcolSysNo = New Collection
                lcolSysNo.Add .TextMatrix(cgrdInfo.Row, 0)
                 
'                 lobj����.Sub��ӡ���� "ְҵ�������_" & .TextMatrix(i, mcolIndex("�������")), lcolSysNo, paraԤ��

                mstrϵͳ��� = cgrdInfo.TextMatrix(cgrdInfo.Row, mcolIndex("ϵͳ���"))    'ȡ��ϵͳ���
                
                'ȡ��������ͣ���Ϊ��ӡ�����Ǹ�������������жϵĴ�ӡ���ű�
                sql = "select ������� from ְҵ�����_��������Ϣ�� where ϵͳ���='" & mstrϵͳ��� & "'"
                Set lobjet = dafuncGetData(sql)
                mstr�������� = lobjet(0)
'                mstr�������� = cgrdInfo.TextMatrix(cgrdInfo.Row, mcolIndex("������"))
'                Set lcolSysNo = New Collection
'                lcolSysNo.Add .TextMatrix(i, 0)
'                 Dim tst As String
'                 tst = "ְҵ�������_" + mstr��������
                 lobj����.Sub��ӡ���� "ְҵ�������_" + mstr��������, lcolSysNo, paraԤ��
'                 lobj����.Sub��ӡ���� "ְҵ�������_" & .TextMatrix(i, mstr��������), mstrϵͳ���, paraԤ��
'                lobj����.Sub��ӡ���� "ְҵ�������_" & .TextMatrix(i, mcolIndex("�������")), mstrϵͳ���, paraԤ��
                
                If paraԤ�� = False Then
'                    dafuncGetData "update ְҵ�����_��������Ϣ�� set ���״̬='7' where ϵͳ���='" & Trim(.TextMatrix(i, 0)) & "'"
                    dafuncGetData "update ְҵ�����_��������Ϣ�� set ���״̬='7' where ϵͳ���='" & mstrϵͳ��� & "'"
                    .RowHidden(i) = True
                End If
'            End If
'        Next i
        If lcolSysNo.Count < 1 And .rows > 1 Then
            MsgBox "�빴ѡҪ��ӡ��Ԥ��������", vbInformation, "ϵͳ��ʾ"
            Exit Sub
        End If
    End With
errHandler:
   
End Sub
