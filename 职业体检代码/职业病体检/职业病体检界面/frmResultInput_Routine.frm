VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmResultInput_Routine 
   BackColor       =   &H80000004&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "��ٿƽ��¼�봰��"
   ClientHeight    =   11160
   ClientLeft      =   1050
   ClientTop       =   315
   ClientWidth     =   18840
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9780
   ScaleMode       =   0  'User
   ScaleWidth      =   18840
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3
      Left            =   5640
      Top             =   360
   End
   Begin VB.PictureBox Picture1 
      Height          =   11655
      Left            =   0
      ScaleHeight     =   11595
      ScaleWidth      =   17715
      TabIndex        =   0
      Top             =   0
      Width           =   17775
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   11535
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   16695
         Begin VB.TextBox Text��ʱ 
            Height          =   375
            Left            =   6720
            MultiLine       =   -1  'True
            TabIndex        =   106
            Text            =   "frmResultInput_Routine.frx":0000
            Top             =   8520
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.CommandButton Command3 
            Caption         =   "�� д �� ��"
            Height          =   375
            Left            =   9240
            TabIndex        =   105
            Top             =   840
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.CommandButton comd�����ն� 
            Caption         =   "�� �� �� ��"
            Height          =   375
            Left            =   7080
            TabIndex        =   104
            Top             =   1320
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.CommandButton Command2 
            Caption         =   "����"
            Height          =   255
            Left            =   11280
            TabIndex        =   98
            Top             =   1560
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.CommandButton ccmdDrawline 
            Caption         =   "�������в���"
            Height          =   495
            Left            =   12600
            TabIndex        =   97
            Top             =   1080
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Ѫ������¼��"
            Height          =   495
            Left            =   12360
            TabIndex        =   96
            Top             =   1080
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.Frame Frame8 
            Caption         =   "��ʷ����"
            Enabled         =   0   'False
            Height          =   2295
            Left            =   11760
            TabIndex        =   79
            Top             =   9120
            Width           =   4815
            Begin VB.TextBox ctxtConclunHistory 
               Height          =   1935
               Left            =   0
               ScrollBars      =   2  'Vertical
               TabIndex        =   81
               Top             =   240
               Width           =   4575
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "�����ʷ���"
            Height          =   6015
            Left            =   11760
            TabIndex        =   78
            Top             =   1920
            Width           =   4815
            Begin VSFlex8Ctl.VSFlexGrid cgrdInputHistory 
               Height          =   5655
               Left            =   120
               TabIndex        =   80
               Top             =   240
               Width           =   4575
               _cx             =   2088771462
               _cy             =   2088773367
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
               BackColorBkg    =   -2147483643
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
               Cols            =   6
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
         Begin VB.CommandButton ccmdEyeDraw 
            Caption         =   "��״�廷�漰����ͼ"
            Height          =   495
            Left            =   8760
            TabIndex        =   20
            Top             =   8520
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.OptionButton coptClasses 
            Caption         =   "���佡��"
            Height          =   255
            Index           =   2
            Left            =   2280
            TabIndex        =   19
            Top             =   960
            Width           =   1095
         End
         Begin VB.OptionButton coptClasses 
            Caption         =   "ְҵ����"
            Height          =   255
            Index           =   1
            Left            =   1200
            TabIndex        =   18
            Top             =   960
            Width           =   1095
         End
         Begin VB.OptionButton coptClasses 
            Caption         =   "��ͨ���"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   17
            Top             =   960
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton coptClasses 
            Caption         =   "��˲���"
            Height          =   255
            Index           =   3
            Left            =   3360
            TabIndex        =   16
            Top             =   960
            Width           =   1095
         End
         Begin VB.OptionButton coptClasses 
            Caption         =   "8023����"
            Height          =   255
            Index           =   4
            Left            =   4440
            TabIndex        =   15
            Top             =   960
            Width           =   1095
         End
         Begin VB.CommandButton Cmd����ģ�� 
            Caption         =   "����ģ��"
            Height          =   495
            Left            =   10320
            TabIndex        =   14
            Top             =   8520
            Width           =   1215
         End
         Begin VB.Frame Frame5 
            Caption         =   "������� (������250������)"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   2295
            Left            =   6600
            TabIndex        =   11
            Top             =   9120
            Width           =   5055
            Begin VB.TextBox ctxtConclun 
               Height          =   1935
               Left            =   120
               MaxLength       =   245
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   12
               Top             =   240
               Width           =   4815
            End
         End
         Begin VB.CommandButton ccmdAutoFull 
            Caption         =   "�� �� �� ��"
            Height          =   375
            Index           =   3
            Left            =   9240
            TabIndex        =   10
            Top             =   1320
            Width           =   1695
         End
         Begin VB.CommandButton ccmdAutoFull 
            Caption         =   "�� �� Ĭ ��"
            Height          =   375
            Index           =   2
            Left            =   7080
            TabIndex        =   9
            Top             =   1320
            Width           =   1695
         End
         Begin VB.CommandButton ccmdAutoFull 
            Caption         =   "ȫ �� �� ��"
            Height          =   375
            Index           =   1
            Left            =   9240
            TabIndex        =   8
            Top             =   840
            Width           =   1695
         End
         Begin VB.CommandButton ccmdAutoFull 
            Caption         =   "δ��д��ȫ������"
            Height          =   375
            Index           =   0
            Left            =   7080
            TabIndex        =   7
            Top             =   840
            Width           =   1695
         End
         Begin VB.Frame Frame3 
            Height          =   580
            Left            =   7200
            TabIndex        =   3
            Top             =   720
            Visible         =   0   'False
            Width           =   4815
            Begin VB.CommandButton WriteConclun 
               Caption         =   "��д����"
               Height          =   375
               Left            =   2760
               TabIndex        =   5
               Top             =   120
               Width           =   1455
            End
            Begin VB.TextBox ctxtDoctor 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   720
               TabIndex        =   4
               Top             =   120
               Width           =   1335
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ҽʦ��"
               Height          =   180
               Index           =   1
               Left            =   120
               TabIndex        =   6
               Top             =   240
               Width           =   540
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "�������д�� "
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   6015
            Left            =   6600
            TabIndex        =   2
            Top             =   1920
            Width           =   5100
            Begin VSFlex8Ctl.VSFlexGrid cgrdInput 
               Height          =   5655
               Left            =   120
               TabIndex        =   21
               Top             =   240
               Width           =   4815
               _cx             =   2088771885
               _cy             =   2088773367
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
               BackColorBkg    =   -2147483643
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
         Begin MSComctlLib.Toolbar ctbMain 
            Height          =   540
            Left            =   0
            TabIndex        =   22
            Top             =   0
            Width           =   12525
            _ExtentX        =   22093
            _ExtentY        =   953
            ButtonWidth     =   820
            ButtonHeight    =   953
            Appearance      =   1
            Style           =   1
            ImageList       =   "cimg��ťͼ��"
            _Version        =   393216
            Begin MSCommLib.MSComm MSComm1 
               Left            =   4320
               Top             =   0
               _ExtentX        =   1005
               _ExtentY        =   1005
               _Version        =   393216
               DTREnable       =   -1  'True
            End
            Begin VB.CheckBox cchkˢ���� 
               Caption         =   "ˢ����"
               Height          =   255
               Left            =   15120
               TabIndex        =   23
               Top             =   100
               Value           =   1  'Checked
               Width           =   1215
            End
         End
         Begin MSComctlLib.ImageList cimg��ťͼ�� 
            Left            =   6120
            Top             =   0
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            MaskColor       =   12632256
            _Version        =   393216
         End
         Begin TabDlg.SSTab ctabPerson 
            Height          =   10095
            Left            =   0
            TabIndex        =   24
            Top             =   1320
            Width           =   6375
            _ExtentX        =   11245
            _ExtentY        =   17806
            _Version        =   393216
            Tabs            =   2
            TabHeight       =   520
            ForeColor       =   -2147483630
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "   ��������  "
            TabPicture(0)   =   "frmResultInput_Routine.frx":0027
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "ccrp����"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "Frame2"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "Frame4"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).ControlCount=   3
            TabCaption(1)   =   "   �������� "
            TabPicture(1)   =   "frmResultInput_Routine.frx":0043
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Ccmb��������"
            Tab(1).Control(1)=   "Timerccrp"
            Tab(1).Control(2)=   "ccmdSelInfo"
            Tab(1).Control(3)=   "cchkCompanyBatch"
            Tab(1).Control(4)=   "ctxtQueyCompanyBatch"
            Tab(1).Control(5)=   "ccmd��ѯ��λ"
            Tab(1).Control(6)=   "fraQueryBatch"
            Tab(1).Control(7)=   "cchkDateBatch"
            Tab(1).Control(8)=   "ccmdClear"
            Tab(1).Control(9)=   "ccmdRemove"
            Tab(1).Control(10)=   "cchkBchResult(0)"
            Tab(1).Control(11)=   "cchkBchResult(1)"
            Tab(1).Control(12)=   "cgrdInfoBatch"
            Tab(1).Control(13)=   "cdtpDateBatch"
            Tab(1).Control(14)=   "Label11(1)"
            Tab(1).Control(15)=   "Label6"
            Tab(1).Control(16)=   "TotalPeopleBatch"
            Tab(1).ControlCount=   17
            Begin VB.ComboBox Ccmb�������� 
               Height          =   300
               Left            =   -73320
               Style           =   2  'Dropdown List
               TabIndex        =   87
               Top             =   3480
               Width           =   2415
            End
            Begin VB.Frame Frame4 
               Caption         =   "�����Ա������Ϣ   "
               ForeColor       =   &H00000000&
               Height          =   2775
               Left            =   120
               TabIndex        =   60
               Top             =   360
               Width           =   6135
               Begin VB.TextBox ctxtweihaiyinsu 
                  Enabled         =   0   'False
                  Height          =   270
                  Left            =   1440
                  Locked          =   -1  'True
                  MaxLength       =   10
                  TabIndex        =   92
                  Top             =   2040
                  Width           =   2655
               End
               Begin VB.TextBox ctxtSingleNo 
                  Height          =   270
                  Left            =   1560
                  MaxLength       =   20
                  TabIndex        =   67
                  Top             =   960
                  Width           =   2655
               End
               Begin VB.PictureBox cpicPhoto 
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   10.5
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1890
                  Index           =   0
                  Left            =   4440
                  ScaleHeight     =   1830
                  ScaleWidth      =   1515
                  TabIndex        =   66
                  Top             =   240
                  Width           =   1575
               End
               Begin VB.TextBox ctxtCompanyName 
                  Height          =   270
                  Left            =   1440
                  MaxLength       =   20
                  TabIndex        =   65
                  Top             =   2400
                  Width           =   3495
               End
               Begin VB.TextBox ctxtAge 
                  Height          =   270
                  Left            =   2880
                  MaxLength       =   3
                  TabIndex        =   64
                  Top             =   1680
                  Width           =   1215
               End
               Begin VB.TextBox ctxtSex 
                  Height          =   270
                  Left            =   960
                  MaxLength       =   1
                  TabIndex        =   63
                  Top             =   1680
                  Width           =   1095
               End
               Begin VB.TextBox ctxtName 
                  Height          =   270
                  Left            =   1440
                  MaxLength       =   10
                  TabIndex        =   62
                  Top             =   1320
                  Width           =   2655
               End
               Begin VB.ComboBox ccmbHistory 
                  Height          =   300
                  Left            =   1440
                  Style           =   2  'Dropdown List
                  TabIndex        =   61
                  Top             =   600
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
                  Left            =   1440
                  TabIndex        =   68
                  Top             =   240
                  Width           =   2655
                  _ExtentX        =   4683
                  _ExtentY        =   450
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
                  Format          =   59965440
                  CurrentDate     =   36951
                  MaxDate         =   73050
                  MinDate         =   17899
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Σ������"
                  Height          =   180
                  Index           =   6
                  Left            =   240
                  TabIndex        =   93
                  Top             =   2040
                  Width           =   720
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "����"
                  Height          =   255
                  Index           =   4
                  Left            =   240
                  TabIndex        =   75
                  Top             =   1320
                  Width           =   975
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "��������"
                  Height          =   255
                  Index           =   3
                  Left            =   240
                  TabIndex        =   74
                  Top             =   960
                  Width           =   975
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "����¼������"
                  Height          =   255
                  Index           =   2
                  Left            =   240
                  TabIndex        =   73
                  Top             =   240
                  Width           =   1080
               End
               Begin VB.Label Label4 
                  Caption         =   "����"
                  Height          =   255
                  Left            =   2280
                  TabIndex        =   72
                  Top             =   1680
                  Width           =   975
               End
               Begin VB.Label Label3 
                  Caption         =   "�Ա�"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   71
                  Top             =   1680
                  Width           =   975
               End
               Begin VB.Label Label5 
                  Caption         =   "��λ����"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   70
                  Top             =   2400
                  Width           =   975
               End
               Begin VB.Label Label13 
                  Caption         =   "���겡��"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   69
                  Top             =   600
                  Width           =   975
               End
            End
            Begin VB.Frame Frame2 
               Caption         =   "��ѯ�����Ա"
               Height          =   6735
               Left            =   120
               TabIndex        =   49
               Top             =   3240
               Width           =   6135
               Begin VB.CheckBox cchkSigResult 
                  Caption         =   "δ����"
                  Height          =   255
                  Index           =   1
                  Left            =   1320
                  TabIndex        =   53
                  Top             =   1440
                  Value           =   1  'Checked
                  Width           =   1215
               End
               Begin MSComCtl2.DTPicker lastcdtpDate 
                  Height          =   255
                  Left            =   1200
                  TabIndex        =   101
                  Top             =   1080
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   450
                  _Version        =   393216
                  Format          =   59965441
                  CurrentDate     =   42461
               End
               Begin MSComCtl2.DTPicker beforecdtpDate 
                  Height          =   255
                  Left            =   1200
                  TabIndex        =   100
                  Top             =   720
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   450
                  _Version        =   393216
                  Format          =   59965441
                  CurrentDate     =   42461
               End
               Begin VB.CheckBox Check��� 
                  Caption         =   "�����"
                  Height          =   255
                  Left            =   1440
                  TabIndex        =   99
                  Top             =   1440
                  Visible         =   0   'False
                  Width           =   975
               End
               Begin VB.CommandButton ccmdSingleQuery 
                  Caption         =   "��ѯ(&Q)"
                  Height          =   375
                  Left            =   4680
                  Style           =   1  'Graphical
                  TabIndex        =   56
                  Top             =   1440
                  Width           =   1065
               End
               Begin VB.TextBox ctxtCheckName 
                  Height          =   270
                  Left            =   4080
                  MaxLength       =   10
                  TabIndex        =   55
                  Top             =   720
                  Width           =   1695
               End
               Begin VB.CheckBox cchkSigResult 
                  Caption         =   "������"
                  Height          =   255
                  Index           =   0
                  Left            =   240
                  TabIndex        =   54
                  Top             =   1440
                  Width           =   1095
               End
               Begin VB.TextBox ctxtcchkWork 
                  Height          =   270
                  Left            =   4080
                  MaxLength       =   20
                  TabIndex        =   52
                  Top             =   360
                  Width           =   1695
               End
               Begin VB.TextBox ctxtcchkNo 
                  Height          =   270
                  Left            =   1200
                  MaxLength       =   20
                  TabIndex        =   51
                  Top             =   360
                  Width           =   1695
               End
               Begin VB.CommandButton ccmdWork 
                  Caption         =   "��λ��λ"
                  Height          =   375
                  Left            =   3120
                  TabIndex        =   50
                  Top             =   1440
                  Width           =   1305
               End
               Begin VSFlex8Ctl.VSFlexGrid cgrdSingleList 
                  Height          =   4575
                  Left            =   120
                  TabIndex        =   57
                  Top             =   2040
                  Width           =   5895
                  _cx             =   10398
                  _cy             =   8070
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
               Begin VB.Label Label9 
                  Caption         =   "��ֹʱ��"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   103
                  Top             =   1080
                  Width           =   855
               End
               Begin VB.Label Label8 
                  Caption         =   "��ʼʱ��"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   102
                  Top             =   720
                  Width           =   735
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "��λ����"
                  Height          =   180
                  Index           =   8
                  Left            =   3120
                  TabIndex        =   85
                  Top             =   360
                  Width           =   720
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "����"
                  Height          =   255
                  Index           =   5
                  Left            =   3120
                  TabIndex        =   84
                  Top             =   720
                  Width           =   855
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "�������"
                  Height          =   180
                  Index           =   0
                  Left            =   240
                  TabIndex        =   83
                  Top             =   360
                  Width           =   720
               End
               Begin VB.Label TotalPeople 
                  AutoSize        =   -1  'True
                  Caption         =   "0"
                  Height          =   180
                  Left            =   840
                  TabIndex        =   59
                  Top             =   1800
                  Width           =   90
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  Caption         =   "������"
                  Height          =   180
                  Left            =   240
                  TabIndex        =   58
                  Top             =   1800
                  Width           =   540
               End
            End
            Begin VB.Timer Timerccrp 
               Left            =   -69240
               Top             =   4560
            End
            Begin VB.CommandButton ccmdSelInfo 
               Caption         =   "�� ѯ"
               Height          =   375
               Left            =   -69720
               TabIndex        =   48
               Top             =   4200
               Width           =   975
            End
            Begin VB.CheckBox cchkCompanyBatch 
               Caption         =   "��λ����"
               Height          =   255
               Left            =   -74760
               TabIndex        =   47
               Top             =   4200
               Width           =   1215
            End
            Begin VB.TextBox ctxtQueyCompanyBatch 
               Height          =   300
               Left            =   -73320
               MaxLength       =   20
               TabIndex        =   46
               Top             =   4200
               Width           =   2415
            End
            Begin VB.CommandButton ccmd��ѯ��λ 
               Caption         =   "��λ��λ"
               Height          =   375
               Left            =   -70800
               TabIndex        =   45
               Top             =   4200
               Width           =   975
            End
            Begin VB.Frame fraQueryBatch 
               Caption         =   "���������Ա��Ϣ"
               Height          =   2895
               Left            =   -74880
               TabIndex        =   30
               Top             =   480
               Width           =   6015
               Begin VB.TextBox ctxtΣ������ 
                  Enabled         =   0   'False
                  Height          =   270
                  Left            =   1560
                  Locked          =   -1  'True
                  MaxLength       =   10
                  TabIndex        =   94
                  Top             =   1800
                  Width           =   2415
               End
               Begin VB.PictureBox Picture4 
                  Height          =   1935
                  Left            =   4200
                  ScaleHeight     =   1875
                  ScaleWidth      =   1515
                  TabIndex        =   37
                  Top             =   360
                  Width           =   1575
               End
               Begin VB.TextBox ctxt��λ���� 
                  Height          =   300
                  Left            =   1560
                  MaxLength       =   20
                  TabIndex        =   36
                  Top             =   2160
                  Width           =   2415
               End
               Begin VB.TextBox ctxt���� 
                  Height          =   300
                  Left            =   3120
                  MaxLength       =   3
                  TabIndex        =   35
                  Top             =   1440
                  Width           =   855
               End
               Begin VB.TextBox ctxt�Ա� 
                  Height          =   300
                  Left            =   1200
                  MaxLength       =   1
                  TabIndex        =   34
                  Top             =   1440
                  Width           =   855
               End
               Begin VB.TextBox ctxt���� 
                  Height          =   300
                  Left            =   1560
                  MaxLength       =   10
                  TabIndex        =   33
                  Top             =   1080
                  Width           =   2415
               End
               Begin VB.TextBox ctxt������� 
                  Height          =   300
                  Left            =   1560
                  MaxLength       =   20
                  TabIndex        =   32
                  Top             =   720
                  Width           =   2415
               End
               Begin VB.CheckBox cchk��������� 
                  Caption         =   "�������Ա�����Ϊ���������¼��"
                  Height          =   255
                  Left            =   360
                  TabIndex        =   31
                  Top             =   2520
                  Value           =   1  'Checked
                  Visible         =   0   'False
                  Width           =   3615
               End
               Begin MSComCtl2.DTPicker DTP¼������ 
                  Height          =   300
                  Left            =   1560
                  TabIndex        =   38
                  Top             =   360
                  Width           =   2415
                  _ExtentX        =   4260
                  _ExtentY        =   529
                  _Version        =   393216
                  Enabled         =   0   'False
                  Format          =   59965440
                  CurrentDate     =   40969
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Σ������"
                  Height          =   180
                  Index           =   7
                  Left            =   360
                  TabIndex        =   95
                  Top             =   1800
                  Width           =   720
               End
               Begin VB.Label Label11 
                  Caption         =   "����¼������"
                  Height          =   255
                  Index           =   0
                  Left            =   360
                  TabIndex        =   44
                  Top             =   360
                  Width           =   1095
               End
               Begin VB.Label Label14 
                  Caption         =   "��λ����"
                  Height          =   255
                  Left            =   360
                  TabIndex        =   43
                  Top             =   2160
                  Width           =   975
               End
               Begin VB.Label Label15 
                  Caption         =   "����"
                  Height          =   255
                  Left            =   2400
                  TabIndex        =   42
                  Top             =   1440
                  Width           =   975
               End
               Begin VB.Label Label16 
                  Caption         =   "�Ա�"
                  Height          =   255
                  Left            =   360
                  TabIndex        =   41
                  Top             =   1440
                  Width           =   975
               End
               Begin VB.Label Label17 
                  Caption         =   "����"
                  Height          =   255
                  Index           =   0
                  Left            =   360
                  TabIndex        =   40
                  Top             =   1080
                  Width           =   975
               End
               Begin VB.Label Label18 
                  Caption         =   "��������"
                  Height          =   255
                  Left            =   360
                  TabIndex        =   39
                  Top             =   720
                  Width           =   975
               End
            End
            Begin VB.CheckBox cchkDateBatch 
               Caption         =   "�������"
               Height          =   255
               Left            =   -74760
               TabIndex        =   29
               Top             =   3840
               Width           =   1215
            End
            Begin VB.CommandButton ccmdClear 
               Caption         =   "�� ��"
               Height          =   375
               Left            =   -70800
               TabIndex        =   28
               Top             =   3720
               Width           =   975
            End
            Begin VB.CommandButton ccmdRemove 
               Caption         =   "�� ��"
               Height          =   375
               Left            =   -69720
               TabIndex        =   27
               Top             =   3720
               Width           =   975
            End
            Begin VB.CheckBox cchkBchResult 
               Caption         =   "������"
               Height          =   255
               Index           =   0
               Left            =   -73320
               TabIndex        =   26
               Top             =   4560
               Width           =   1095
            End
            Begin VB.CheckBox cchkBchResult 
               Caption         =   "δ����"
               Height          =   255
               Index           =   1
               Left            =   -72000
               TabIndex        =   25
               Top             =   4560
               Value           =   1  'Checked
               Width           =   1095
            End
            Begin VSFlex8Ctl.VSFlexGrid cgrdInfoBatch 
               Height          =   4935
               Left            =   -74880
               TabIndex        =   76
               Top             =   5040
               Width           =   6135
               _cx             =   2088774213
               _cy             =   2088772097
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
            Begin MSComCtl2.DTPicker cdtpDateBatch 
               Height          =   300
               Left            =   -73320
               TabIndex        =   77
               Top             =   3840
               Width           =   2415
               _ExtentX        =   4260
               _ExtentY        =   529
               _Version        =   393216
               Format          =   59965440
               CurrentDate     =   40969
            End
            Begin VB.PictureBox ccrp���� 
               Height          =   375
               Left            =   0
               ScaleHeight     =   315
               ScaleWidth      =   4635
               TabIndex        =   86
               Top             =   360
               Visible         =   0   'False
               Width           =   4695
            End
            Begin VB.Label Label11 
               Caption         =   "��������"
               Height          =   255
               Index           =   1
               Left            =   -74760
               TabIndex        =   90
               Top             =   3540
               Width           =   1095
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "������"
               Height          =   180
               Left            =   -74640
               TabIndex        =   89
               Top             =   4560
               Width           =   540
            End
            Begin VB.Label TotalPeopleBatch 
               AutoSize        =   -1  'True
               Caption         =   "0"
               Height          =   180
               Left            =   -74520
               TabIndex        =   88
               Top             =   4800
               Width           =   210
            End
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "Label7"
            ForeColor       =   &H008080FF&
            Height          =   180
            Left            =   7080
            TabIndex        =   91
            Top             =   1680
            Width           =   540
         End
         Begin VB.Label Label17 
            Caption         =   "�б�����"
            Height          =   255
            Index           =   1
            Left            =   -74760
            TabIndex        =   82
            Top             =   4560
            Width           =   735
         End
      End
      Begin VB.Label LabelDoctor 
         Caption         =   "ҽ����"
         Height          =   255
         Left            =   5640
         TabIndex        =   13
         Top             =   960
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmResultInput_Routine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private WithEvents mobj����ͨ�ö��� As cls����ͨ�ö���    '�ṩ��������ʼ�����ȼ�����
Attribute mobj����ͨ�ö���.VB_VarHelpID = -1
'Public InputFlag As String
'Public InputFlagNo As String
Private mobj���ҽʦ  As Object   'clsMedicalExamer    ��ȡ��ǰ���ҽʦ��������ָ�����ԣ�����/���飩�������Ŀ
Private mobjQueryResult As Object
Private mstr�������� As String  '��������ʱ����ǰһ������¼��ʹ�õ�����ģ�����ơ�
Private mstr��������  As String   '��Ӧ����"��������"��
Private mstr�����Ŀ���� As String
Private mstrϵͳ��Ź̶����� As String
Private mstrȨ�ޱ�־ As Boolean
'2012-07-14 �ڵ�� ��
'���ӿ��һ�����Ϣ����
Private priDeptName As String
Private priDeptNo As String
Private priDeptResultName As String
'2012-07-14 �ڵ�� ��

'����Ȩ�ޱ�־
'��¼�ڵ�һ�α��������֮������ٴ��޸Ľ������Ҫ������������޸ģ��Ƿ񱣴桱֮�����ʾ��
'-1����ʾδ��ȡ�������ݿ������������Ϣ��
'0����ʾ���˵Ľ��δ¼�����
'1����ʾ���ݿ������и��˵Ľ�������ڽ�����δ���޸Ĺ���
'2����ʾ���ݿ������и��˵Ľ�������������޸Ĺ���ֻ����Ϊ2��ʱ�򣬲Żᵯ����������޸ģ��Ƿ񱣴桱����
'3����ʾû��Ȩ�޽����޸Ĳ�����
Private ResultChanged As Integer

Private lcolResult As Collection    '��������ϣ�item:[�����Ŀ���ƣ������]��
Private lcolItem As Collection      '���������Ŀ���������[�����Ŀ���ƣ������]��

Private mblnSys As Boolean
Public mblnInUse As Boolean      '��Ӧ����"pblnInUse"
Private mobj���� As cls�û��������� '�޸ģ�2001-12-29�����Ӹö��󣩡�
Private mstrState As String     '��¼��ǰ���״̬
Private mintFixed As Integer
Private mcol�����Ŀ As Collection

Private mstr���ͼƬ���� As String

'���ܣ�������ǰ�����Ƿ�һ���أ��Ա������������жϵ�ǰ�����Ƿ���ִ�й�Form_Load��
Public Property Get pblnInUse() As Boolean
Attribute pblnInUse.VB_Description = "'���ܣ�������ǰ�����Ƿ�һ���أ��Ա������������жϵ�ǰ�����Ƿ���ִ�й�Form_Load��"
    pblnInUse = mblnInUse
End Property

Public Property Let pblnInUse(pblnInUse As Boolean)
    mblnInUse = pblnInUse
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

'2012-07-14 �ڵ�� ���ģ�2012��10��16�� �����
Private Sub cchkBchResult_Click(Index As Integer)
    Dim i, coptIndex As Integer
     If cchkBchResult(0).Value = 0 And cchkBchResult(1).Value = 0 Then
            cgrdInfoBatch.Clear '���cgrdSingleList������
            cgrdInfoBatch.rows = 1
            cgrdInfoBatch.cols = 0
            With cgrdInfoBatch
                .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "ϵͳ���"
                .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "����"
                .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "�Ա�"
                .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "����"
                .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "������"
                .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "�������"
                .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "��λ����"
                .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "��дʱ��"
                .AutoSize 0, .cols - 1, 0, 0
                .SelectionMode = flexSelectionListBox
            End With
            Exit Sub
        End If
    
    For i = 0 To coptClasses.Count - 1
        If coptClasses(i).Value = True Then coptIndex = i
    Next
    sub��ѯ�б���ʾ coptIndex
End Sub

'2012-07-14 �ڵ��  ���ģ������ 2012��10��15��
Private Sub cchkSigResult_Click(Index As Integer)
    Dim i, coptIndex As Integer
        If cchkSigResult(0).Value = 0 And cchkSigResult(1).Value = 0 Then
            cgrdSingleList.Clear '���cgrdSingleList������
            cgrdSingleList.cols = 0
            cgrdSingleList.rows = 1
            TotalPeople.Caption = 0 '������Ϊ0    2015-12-23 by Ĳ��
            With cgrdSingleList
                .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "ϵͳ���"
                .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "����"
                .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "�Ա�"
                .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "����"
                .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "������"
                .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "�������"
                .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "��λ����"
                .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "��дʱ��"
                .AutoSize 0, .cols - 1, 0, 0
                .SelectionMode = flexSelectionListBox
            End With
             Exit Sub
        End If
    For i = 0 To coptClasses.Count - 1
        If coptClasses(i).Value = True Then coptIndex = i
    Next
    sub��ѯ�б���ʾ coptIndex
End Sub

'��ʾѡ�����ڵĲ�����Ϣ
'����
'2012-07-31
Private Sub ccmbHistory_Click()
    Dim lcolInfo As Collection '��ŵ�ǰϵͳ����е�ǰҽʦ��������ָ�����Ե������Ŀ��������
    Dim lobjItem As Variant    'clsFactTestItem,lcolInfo�е�Ԫ�ء�
    Dim lstrEnum As String
    Dim i As Long
    Dim j As Long
'    On Error Resume Next
    
    'ʹ���Ż��㷨��ȡ�����Ŀ��
    Dim lobjRec As Object
    Dim lobjTemp As Object
    
    '�޸��ˣ������  ʱ�䣺2013-1-4 ��
    '˵������Ӳ�ѯ�������ʷ��¼
    'bug�ţ�0000152
    With cgrdInputHistory
        .rows = 1
        .cols = 6
        .TextMatrix(0, 0) = "����"
        .ColWidth(0) = 700
        .TextMatrix(0, 1) = "����"
        .ColWidth(1) = 1600
        .TextMatrix(0, 2) = "�����"
        .ColWidth(2) = 1000
        .ColHidden(3) = True '��Ÿ������Ŀ��ö����Դ��
        .ColWidth(4) = 1200
        .TextMatrix(0, 4) = "��׼ֵ"
        .ColWidth(5) = 600
        .TextMatrix(0, 5) = "��λ"
    End With
    
      '�޸��ˣ������  ʱ�䣺2013-1-4 ��
    If ccmbHistory.Text <> "����" Then
'        ctbMain.Buttons(2).Enabled = False
        Set lobjRec = mobj���ҽʦ.func��ȡָ����ݵ���첡��(Trim(ctxtSingleNo.Text), ccmbHistory.Text, InputFlag)

        If Not lobjRec Is Nothing Then
        
            '������Ŀ�����ʾ����
'            Chk����ģ��.Visible = False
'            Cmd����ģ��.Visible = False
'            Frame5.Visible = False
'            Frame6.Caption = "�����Ա���겡����"
'            Frame6.Height = Frame6.Height + 300
'            cgrdInput.Height = cgrdInput.Height - 300
            
            
            '���ĵ�ǰ¼��״̬
            If IsNull(lobjRec("�������")) Then
                ResultChanged = IIf(ResultChanged <> 3, 0, 3)
            Else
                ResultChanged = IIf(ResultChanged <> 3, 1, 3)
            End If
            
            cgrdInputHistory.rows = lobjRec.RecordCount + 1
            i = 1
            Do While Not lobjRec.EOF
                cgrdInputHistory.TextMatrix(i, 0) = lobjRec!�����Ŀ���
                cgrdInputHistory.TextMatrix(i, 1) = lobjRec!�����Ŀ����
                cgrdInputHistory.TextMatrix(i, 2) = IIf(IsNull(lobjRec!�����), "", lobjRec!�����)
                
                '���ݵ������������ɫ��
                Dim lstr������� As String
                If IIf(IsNull(lobjRec!�������), "", lobjRec!�������) = "" And cgrdInputHistory.TextMatrix(i, 2) <> "" Then
                    '�����µ�����ۡ�
                    lstr������� = pobjҵ�����.func��ȡ�������(lobjRec!�����Ŀ���, IIf(IsNull(lobjRec!�����), "", lobjRec!�����))
                Else
                    lstr������� = IIf(IsNull(lobjRec!�������), "", lobjRec!�������)
                End If
                If lstr������� = "���ϸ�" Then
                    '������ɫ��
                    cgrdInputHistory.Cell(flexcpBackColor, i, 2, i, 2) = &H8A5AFA
                Else
                    '������ɫ��
                    cgrdInputHistory.Cell(flexcpBackColor, i, 2, i, 2) = vbWhite
                End If
                '��ö����Դ��ת��ΪGrid����ʶ���ColComboList�����ԡ�|����������
                lstrEnum = IIf(IsNull(lobjRec!ö����Դ), "", lobjRec!ö����Դ)
                lstrEnum = gffuncStrReplace(lstrEnum, ",", "|")
                lstrEnum = gffuncStrReplace(lstrEnum, "��", "|")
                cgrdInputHistory.TextMatrix(i, 3) = lstrEnum
    
                cgrdInputHistory.TextMatrix(i, 4) = IIf(IsNull(lobjRec!��׼ֵ), "", lobjRec!��׼ֵ)
                cgrdInputHistory.TextMatrix(i, 5) = IIf(IsNull(lobjRec!��λ), "", lobjRec!��λ)
    
                i = i + 1
                lobjRec.MoveNext
            Loop
            '��ӿ����б���ʾ����
            With cgrdInputHistory
                .Col = 0
                .Sort = flexSortGenericAscending
                .AutoSize 0, .cols - 1, 0, 0
                .ExplorerBar = flexExSort
                .DataMode = flexDMFree
            End With
            
            Set lobjRec = mobj���ҽʦ.func��ȡָ����ݵ���첡������(Trim(ctxtSingleNo.Text), InputFlagNo, Trim(ccmbHistory.Text))
            If Not lobjRec Is Nothing Then
                ctxtConclunHistory.Text = lobjRec("���ֽ���")
            End If
    
            cgrdInputHistory.Select 1, 2, 1, 2
'            cgrdInputHistory.Enabled = False
        Else
            cgrdInputHistory.rows = 1
            ctbMain.Buttons(3).Enabled = False
            
        End If
        
'    ElseIf ccmbHistory.Text = "����" Or ccmbHistory.Text = "" Then
'
''        Chk����ģ��.Visible = True
'        Cmd����ģ��.Visible = True
''        Frame5.Visible = True
''        Frame6.Caption = "�����Ŀ�����д��"
''        Frame6.Height = Frame6.Height - 300
''        cgrdInputHistory.Height = cgrdInputHistory.Height - 300
'        cgrdInputHistory.Enabled = True
'        cgrdInputHistory.rows = 1
'
'        subShowInputGridHistory Trim(ctxtSingleNo.Text)
    Else
        cgrdInputHistory.rows = 1
        ctxtConclunHistory.Text = ""
        ctbMain.Buttons(3).Enabled = False
    End If
    
End Sub

Private Sub Ccmb��������_Click()
    Dim i As Integer
    For i = 0 To coptClasses.Count - 1
        If coptClasses(i) Then
            sub��ѯ�б���ʾ i
        End If
    Next
    
End Sub

Private Sub ccmdClear_Click()
    cgrdInfoBatch.rows = 1
    TotalPeopleBatch.Caption = 0
End Sub

Private Sub ccmdDrawline_Click()

frmdrawline.lblno.Caption = ctxtSingleNo.Text
frmdrawline.Show 1
End Sub

'2012-07-15 �ڵ��
'��ٿ���ӻ�ͼ���ܡ�
Private Sub ccmdEyeDraw_Click()
    frmEyeDraw.mstr���ͼƬ���� = mstr���ͼƬ����
    frmEyeDraw.pubSysNo = IIf(ctabPerson.Tab = 0, ctxtSingleNo.Text, ctxt�������.Text)
    frmEyeDraw.Show 1
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
    'lstrWhere = " and �������='" & coptClasses(coptIndex).Caption & "'"

    '��װ��ѯ����
    If cchkDateBatch.Value = 1 Then
        lstrWhere = lstrWhere & " and �������>=''" & Format(cdtpDateBatch.Value, "yyyy-mm-dd 00:00:00") & "'' and �������<=''" & Format(cdtpDateBatch.Value, "yyyy-mm-dd 23:59:59") & "''"
    End If
    
    If cchkCompanyBatch.Value = 1 Then
        lstrWhere = lstrWhere & " and ��λ����=''" & Trim(ctxtQueyCompanyBatch.Text) & "''"
    End If
    
    '2012-07-14 �ڵ�� ��
    '���Ĳ�ѯ����������8/48Сʱ�ж����ݡ������޸�ʱ���ʼ�ղ������ѯ����С�
    '��ѯ���ݱ�����ݷ����ϴ�仯�����޸ģ������⡣

    '���ÿ������������������Ա�޸�ʱ�����¸��¡���������Ϣ���С��������״̬����'2'��Ϊ'3'�ģ���ѯʱ���ԡ�
    sub���¿��޸Ľ����Ա�޸�״̬
    
    Set lobjTmp = CreateObject("ְҵ�������¼��.clscommon")
    Set mobjQueryResult = lobjTmp.func��ȡ���޸Ľ���_������_�����Ա��Ϣ(lstrWhere, priDeptName)
    
    cchkBchResult_Click 1
'    sub��ѯ�б���ʾ coptIndex
    '2012-07-14 �ڵ�� ��
    TotalPeopleBatch.Caption = cgrdInfoBatch.rows - 1
    Set lobjTmp = Nothing
    Set lobjRec = Nothing
    lstrWhere = ""
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�������¼��", InputFlag & "���¼��", "ccmdQuery_Click", 6666, lstrError, False
End Sub

Private Sub ccmdWork_Click()
    On Error GoTo errHandler
    Dim lobjRec As Object  '��λ��λ���صĽ����¼��
    Dim lobj��λ As Object
    Dim lobj��λ��Ϣ As Object
    Dim mstr��λ������ As String
    '������λ��λ���档
    Set lobjRec = pobjҵ�����.func��λ��λ
    '��ȡ��λ�ĵ�λ����ʾ�ڡ���λ���ơ�¼����С�
    If Not lobjRec Is Nothing Then
        If lobjRec.RecordCount > 0 Then
            ctxtcchkWork.Text = IIf(IsNull(lobjRec("��λ����")), "", lobjRec("��λ����"))
'            mstr��λ������ = lobjRec!������
            'Set lobj��λ = CreateObject("ְҵ������.class1")
            'lobj��λ.��λ��Ϣ���� = lobjRec!������
            'Set lobj��λ��Ϣ���� = lobj��λ.��λ��Ϣ
            
            
            
'            If mstr��λ������ <> "" Then
'                '�޸ģ�2001-8-23����ʾ��λ���ԣ���
'                On Error Resume Next
'                'sub��ʾ��λ���� ciptBase, mstr��λ������, mobjGUI
'                func��ȡ��λ��Ϣ lobjRec!������
'            End If
        End If
    End If
    
    '�ѽ���ص���λ¼��򡣱����ܱ����µ�λ��λ��Ϣ��
    ctxtcchkWork.SetFocus
'    SendKeys vbTab
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����沿��", InputFlag & "���¼��", "ccmd��λ��λ_Click", 6666, lstrError, False
End Sub

Private Sub ccmd��ѯ��λ_Click()
    Dim lobjRec As Object                       '��λ��λ���صĽ����¼��
    
    On Error GoTo errHandler
    Set lobjRec = pobjҵ�����.func��λ��λ     '������λ��λ���档
    
    '��ȡ��λ�ĵ�λ����ʾ�ڡ���λ���ơ�¼����С�(��ʱֻ��ʾ����λ���ơ�)
    If Not lobjRec Is Nothing Then
        If lobjRec.RecordCount > 0 Then
            ctxtQueyCompanyBatch.Text = IIf(IsNull(lobjRec("��λ����")), "", lobjRec("��λ����"))
        End If
    End If
    'flag����.Value = 1
    Exit Sub
errHandler:
    'If Err.Number = 0 Then Exit Sub
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����沿��", InputFlag & "���¼��", "ccmd��λ��λ_Click", 6666, lstrError, False
End Sub

'�������������ر�ʱ������
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
                    If cgrdInput.TextMatrix(i, 4) <> "" Then
                        cgrdInput.TextMatrix(i, 2) = cgrdInput.TextMatrix(i, 4)
                    Else
                        cgrdInput.TextMatrix(i, 2) = "����"
                    End If
'                    cgrdInput.TextMatrix(i, 2) = "����"
                End If
            Next
        Case 1
            For i = 1 To cgrdInput.rows - 1
                If cgrdInput.TextMatrix(i, 4) <> "" Then
                    cgrdInput.TextMatrix(i, 2) = cgrdInput.TextMatrix(i, 4)
                Else
                    cgrdInput.TextMatrix(i, 2) = "����"
                End If
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
    
'    ��ʾָ��������ڵ�δ�½��۵������Ա����
'    subShowSingleList
'    ��װ��ѯ����
    
    For i = 0 To coptClasses.Count - 1
        If coptClasses(i).Value = True Then
            coptIndex = i
            Exit For
        End If
    Next
    'lstrWhere = " and �������='" & coptClasses(coptIndex).Caption & "'"
    
    '����ʱ����ƣ���ʼĬ�ϵ��ǵ�ǰһ���ʱ�� 2016-4-1 by Ĳ��
'    lstrWhere = lstrWhere & " and a.�������>=''" & Format(beforecdtpDate.Value, "yyyy-mm-dd 00:00:00") & "'' and a.�������<=''" & Format(lastcdtpDate.Value, "yyyy-mm-dd 23:59:59") & "''"

    '2012-07-24 ���� �޸ģ�����ɸѡ������
    'ϵͳ���
    If Trim(ctxtcchkNo.Text) <> "" And (Len(Trim(ctxtcchkNo.Text)) = 15 Or Len(Trim(ctxtcchkNo.Text)) = 16) Then
        lstrWhere = lstrWhere & " and a.ϵͳ���=''" & Trim(ctxtcchkNo.Text) & "''"
    End If
    '���֤��
'    If Trim(ctxtcchkCardNo.Text) <> "" Then
'        lstrWhere = lstrWhere & " and ������ݺ���='" & Trim(ctxtcchkCardNo.Text) & "'"
'    End If
    '����
    If Trim(ctxtCheckName.Text) <> "" Then
        lstrWhere = lstrWhere & " and a.����=''" & Trim(ctxtCheckName.Text) & "''"
    End If
    '������λ
    If Trim(ctxtcchkWork.Text) <> "" Then
        lstrWhere = lstrWhere & " and a.��λ����=''" & Trim(ctxtcchkWork.Text) & "''"
    End If
    '����ʱ����ƣ���ʼĬ�ϵ��ǵ�ǰһ���ʱ�� 2016-4-1 by Ĳ��
    lstrWhere = lstrWhere & " and (a.�������>=''" & Format(beforecdtpDate.Value, "yyyy-mm-dd 00:00:00") & "'' and a.�������<=''" & Format(lastcdtpDate.Value, "yyyy-mm-dd 23:59:59") & "'')"
    '2012-07-24 ���� �޸ģ�����ɸѡ������

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
    
    
     '20150414 ע�� liuwei
     '  Dim objds As Object
     ' Set objds = dafuncGetData("select �����   from ְҵ�����_�����Ϣ_�ڿ� where �����Ŀ='02002' and ϵͳ���='" & ctxtcchkNo.Text & "'")
     'If InputFlag = "�ڿ�" And objds("�����") <> "" Then
     '  Label8.Visible = True
     ' Label9.Visible = True
     '  Label9.Caption = objds("�����")
     '  Else
     '    Label9.Caption = "��δ��顣"
     '   Label8.Visible = False
     '   Label9.Visible = False
   'End If
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����沿��", InputFlag & "���¼��", "ccmdSingleQuery_Click", 6666, lstrError, False
End Sub

'Private Sub subָ��������ѯ()
'
'    Dim lobjTmp, lobjRec As Object
'    Dim i As Integer, j As Integer
'    Dim lstrWhere As String
'    Dim coptIndex As Integer
'
'    '��ʾָ��������ڵ�δ�½��۵������Ա������
'    'subShowSingleList
'    '��װ��ѯ����
'
'    For i = 0 To coptClasses.Count - 1
'        If coptClasses(i).Value = True Then
'            coptIndex = i
'            Exit For
'        End If
'    Next
'
'    If mobjQueryResult.recordcount > 0 Then
'
'        lstrWhere = lstrWhere & " and �������>='" & Format(cdtpDate.Value, "yyyy-mm-dd 00:00:00") & "' and �������<='" & Format(cdtpDate.Value, "yyyy-mm-dd 23:59:59") & "'"
'
'        '2012-07-24 ���� �޸ģ�����ɸѡ������
'        'ϵͳ���
'        If Trim(ctxtcchkNo.Text) <> "" And Len(Trim(ctxtcchkNo.Text)) = 16 Then
'            mobjQueryResult.Filter = "ϵͳ���='" & Trim(ctxtcchkNo.Text) & "'"
''            lstrWhere = lstrWhere & " and a.ϵͳ���='" & Trim(ctxtcchkNo.Text) & "'"
'        End If
'        '����
'        If Trim(ctxtCheckName.Text) <> "" Then
'            mobjQueryResult.Filter = "����='" & Trim(ctxtCheckName.Text) & "'"
''            lstrWhere = lstrWhere & " and ����='" & Trim(ctxtCheckName.Text) & "'"
'        End If
'        '������λ
'        If Trim(ctxtcchkWork.Text) <> "" Then
'            mobjQueryResult.Filter = "��λ����='" & Trim(ctxtcchkWork.Text) & "'"
''            lstrWhere = lstrWhere & " and ��λ����='" & Trim(ctxtcchkWork.Text) & "'"
'        End If
'
'        If mobjQueryResult.Filter <> "" And mobjQueryResult.Filter <> 0 And mobjQueryResult.Filter <> "ϵͳ���='xxx'" Then
'            mobjQueryResult.Filter = mobjQueryResult.Filter & " and �������='" & coptClasses(coptIndex).Caption & "'"
'        Else
'            mobjQueryResult.Filter = "�������='" & coptClasses(coptIndex).Caption & "'"
'        End If
'
'         '2012-10-26 �����
'        If ctabPerson.Tab = 1 Then
'            mobjQueryResult.Filter = mobjQueryResult.Filter & " and ������='" & Trim(Ccmb��������.Text) & "'"
'        End If
'
'    End If 'mobjQueryResult.recordcount = 0
'
'    If ctabPerson.Tab = 0 Then
'        With cgrdSingleList
'            Set .DataSource = mobjQueryResult
'            .col = 0
'            .Sort = flexSortGenericDescending
'            .AutoSize 0, .cols - 1, 0, 0
'            .ExplorerBar = flexExSort
'            .DataMode = flexDMFree
'            .AllowSelection = True
'            .AllowBigSelection = False
'            .SelectionMode = flexSelectionByRow
'        End With
'        TotalPeople.Caption = IIf(mobjQueryResult.recordcount = 0, "0", mobjQueryResult.recordcount)
'    Else
'        With cgrdInfoBatch
'            Set .DataSource = mobjQueryResult
'            .col = 0
'            .Sort = flexSortGenericDescending
'            .AutoSize 0, .cols - 1, 0, 0
'            .ExplorerBar = flexExSort
'            .DataMode = flexDMFree
'            .AllowSelection = True
'            .AllowBigSelection = True
'            .SelectionMode = flexSelectionListBox
'        End With
'        TotalPeopleBatch.Caption = IIf(mobjQueryResult.recordcount = 0, "0", mobjQueryResult.recordcount)
'    End If
'    cgrdInput.rows = 1
'
'End Sub



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
        ctxt�������.Text = cgrdInfoBatch.TextMatrix(indX, 0)
        ctxt�������_KeyDown 13, 0
    End If
End Sub

Private Sub cgrdInput_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lstr������� As String
    On Error GoTo errHandler
    If Row > 0 Then
        lstr������� = pobjҵ�����.func��ȡ�������(cgrdInput.TextMatrix(Row, 0), cgrdInput.TextMatrix(Row, 2))
        If lstr������� = "���ϸ�" Then
            '������ɫ��
            cgrdInput.Cell(flexcpBackColor, Row, 2, Row, 2) = &H8A5AFA
        Else
            '������ɫ��
            cgrdInput.Cell(flexcpBackColor, Row, 2, Row, 2) = vbWhite
        End If
        
        If cgrdInput.TextMatrix(Row, 2) = "�쳣" Then
            cgrdInput.Cell(flexcpBackColor, Row, 2, Row, 2) = &H8A5AFA
        End If
        
        '2012-06-21 �ڵ�� ��
        '���õ�ǰ¼��״̬(�Ѿ�¼����������޸ĵ�ǰ���)
        If ResultChanged = 1 Then ResultChanged = 2
        '2012-06-21 �ڵ�� ��
        
        cgrdInput.AutoSize 0, cgrdInput.cols - 1, 0, 0
        If cgrdInput.ColWidth(2) < 2000 Then
            cgrdInput.ColWidth(2) = 2000
        End If
    End If
    Exit Sub
errHandler:
End Sub

Private Sub cgrdInput_DblClick()
    '�޸���ɫ��
    On Error Resume Next
    If cgrdInput.Row > 0 Then
        If cgrdInput.Cell(flexcpBackColor, cgrdInput.Row, 2, cgrdInput.Row, 2) = &H8A5AFA Then
            cgrdInput.Cell(flexcpBackColor, cgrdInput.Row, 2, cgrdInput.Row, 2) = vbWhite
        Else
            cgrdInput.Cell(flexcpBackColor, cgrdInput.Row, 2, cgrdInput.Row, 2) = &H8A5AFA
        End If
        '2012-06-21 �ڵ�� ��
        '���õ�ǰ¼��״̬(�Ѿ�¼����������޸ĵ�ǰ���)
        If ResultChanged = 1 Then ResultChanged = 2
        '2012-06-21 �ڵ�� ��
    End If
End Sub

Private Sub cgrdInput_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    On Error GoTo errHandler
    If Col = 2 And KeyCode = 13 Then
        '���С�
        If Row = cgrdInput.rows - 1 Then
            cgrdInput.Row = 1
        Else
            cgrdInput.Row = cgrdInput.Row + 1
        End If
        cgrdInput.Col = 2
    End If
    Exit Sub
errHandler:

End Sub

Private Sub cgrdInput_LostFocus()
     cgrdInput.AutoSize 0, cgrdInput.cols - 1, 0, 0
End Sub

Private Sub cgrdInput_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If cgrdInput.Col = 2 And Button = 1 Then
     '   cgrdInput.EditCell
    End If
End Sub
'(����¼��)
'Private Sub cgrdPerson_AfterEdit(ByVal row As Long, ByVal col As Long)
'    Dim lstr������� As String
'    On Error GoTo errHandler
'    If row > 0 Then
'        lstr������� = pobjҵ�����.func��ȡ�������(cgrdPerson.TextMatrix(row, 0), cgrdPerson.TextMatrix(row, col))
'        If lstr������� = "���ϸ�" Then
'            '������ɫ��
'            cgrdPerson.Cell(flexcpBackColor, row, col, row, col) = &H8A5AFA
'        Else
'            '������ɫ��
'            cgrdPerson.Cell(flexcpBackColor, row, col, row, col) = vbWhite
'        End If
'    End If
'    Exit Sub
'errHandler:
'End Sub

Private Sub cgrdPerson_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Row < 1 Or Col <= mintFixed Then
        Cancel = True
    End If
    
End Sub

'�����б� ������ݣ��Ҳ�����ʾ�����Ŀ����
Private Sub cgrdSingleList_dblClick()
    'If cgrdInput.rows < 2 Then
        If cgrdSingleList.Row > 0 Then
        
            cgrdInputHistory.Clear
            cgrdInputHistory.Enabled = False
            ctxtConclunHistory.Text = ""
            ctbMain.Buttons(3).Enabled = False
            ctxtSingleNo.Text = cgrdSingleList.Cell(flexcpText, cgrdSingleList.Row, 0)
            
            '2012-07-15 �ڵ�� ��
            '������Ϣ�������ƹ��ܲ�ȫ����ֱ�ӵ���ctxtsingleno_keydown

            Cmd����ģ��.Visible = True
            Frame5.Visible = True
            Frame6.Caption = "�������д��"
    '        Frame6.Height = Frame6.Height - 300
    '        cgrdInput.Height = cgrdInput.Height - 300
    
    
            cgrdInput.Enabled = True
            cgrdInput.rows = 1
            
            ctxtSingleNo_KeyDown 13, 0
            
'''            '��ʾ��Ա��Ϣ��
'''            subShowSinglePerson
            '2012-07-15 �ڵ�� ��
            
            If InputFlag = "�������" Then
            ccmdDrawline.Visible = True
            End If
            
            
            
        End If
        
    'Else
    '    MsgBox "���ȱ��浱ǰ�����Ա��Ϣ��"
    'End If
End Sub

Private Sub clblInfo_Click(Index As Integer)
End Sub

'2012-05-11 ��¶
'�������еĽ���ģ�� �ɽ���ѡ��
Private Sub Cmd����ģ��_Click()
    frmConclusion.lobj���ÿ��� = Me.name
    frmConclusion.lobj���� = priDeptName
    frmConclusion.lobj���ұ�� = priDeptNo
    frmConclusion.lobjҽ����� = um�û����
    frmConclusion.lobjʱ�� = Now
    frmConclusion.Show
End Sub
'2012-05-11 ��¶



Private Sub comd�����ն�_Click()
      With MSComm1
       
        .CommPort = 1
        .Settings = "9600,N,8,1"
        .InBufferSize = 1024 'ԭ��Ϊ19
        .RThreshold = 1      '����1�ֽڴ���oncomm�¼�
        .InputMode = comInputModeBinary
        .InputLen = 1 '���볤��Ϊ19
        .InBufferCount = 0      '������ջ�����
    End With
        '�򿪶˿�
    If MSComm1.PortOpen Then MSComm1.PortOpen = False
                
'                MSComm1.CommPort = 6  '�ٶ�����COM5��
                MSComm1.CommPort = 1  '�򳣹�����COM1��
                
                ' �趨�������ʵȣ������������������
                MSComm1.Settings = "9600,N,8,1"
    
                MSComm1.PortOpen = True
                Text��ʱ.Text = ""
End Sub


Private Sub Command1_Click()
frmfresult.Labelҽʦ���.Caption = um�û����
frmfresult.Show 1

End Sub

Private Sub Command2_Click()
frmshenghua.Show 1
End Sub

Private Sub Command3_Click()
'ע�⣺����д������Ϣʱ����Ԫ����Ҫ���ݾ���������
    Dim A, B, C, Temp() As String
    A = Trim(Text��ʱ.Text)
    Temp = Split(A, "-")
    '��ԭURO
'ȡ����ֵ��0��12�ȣ��������ѡһ������ȡ���ֻ��ǽ����������ʱȡ���
'    B = Trim(Temp(5))
'    C = Trim(Split(B, "umol")(0))
'    cgrdInput.TextMatrix(1, 2) = C

'ȡ�����-��+��2+�ȣ�
    Temp = Split(A, "URO")
    B = Trim(Temp(1))
    C = Trim(Split(B, "umol/L")(0))
    C = Trim(Left(C, 2))
    cgrdInput.TextMatrix(1, 2) = C
    
    '������GLU
'ȡ����ֵ��0��12�ȣ��������ѡһ������ȡ���ֻ��ǽ������
'    B = Trim(Temp(10))
'    C = Trim(Split(B, "mmol")(0))
'    cgrdInput.TextMatrix(2, 2) = C

'ȡ�����-��+��2+�ȣ�
    Temp = Split(A, "GLU")
    B = Trim(Temp(1))
    C = Trim(Split(B, "mmol/L")(0))
    C = Trim(Left(C, 3))
    cgrdInput.TextMatrix(2, 2) = C
    
    '��ͪ��KET
'ȡ����ֵ��0��12�ȣ��������ѡһ������ȡ���ֻ��ǽ������
'    B = Trim(Temp(8))
'    C = Trim(Split(B, "mmol")(0))
'    cgrdInput.TextMatrix(3, 2) = C

'ȡ�����-��+��2+�ȣ�
    Temp = Split(A, "KET")
    B = Trim(Temp(1))
    C = Trim(Split(B, "mmol/L")(0))
    C = Trim(Left(C, 3))
    cgrdInput.TextMatrix(3, 2) = C
    
    '������BIL
'ȡ����ֵ��0��12�ȣ��������ѡһ������ȡ���ֻ��ǽ������
'    B = Trim(Temp(9))
'    C = Trim(Split(B, "umol")(0))
'    cgrdInput.TextMatrix(4, 2) = C
    
'ȡ�����-��+��2+�ȣ�
    Temp = Split(A, "BIL")
    B = Trim(Temp(1))
    C = Trim(Split(B, "umol/L")(0))
    C = Trim(Left(C, 3))
    cgrdInput.TextMatrix(4, 2) = C
    
    '������PRO
'ȡ����ֵ��0��12�ȣ��������ѡһ������ȡ���ֻ��ǽ������
'    B = Trim(Temp(6))
'    C = Trim(Split(B, "g/L")(0))
'    cgrdInput.TextMatrix(5, 2) = C
    
'ȡ�����-��+��2+�ȣ�
    Temp = Split(A, "PRO")
    B = Trim(Temp(1))
    C = Trim(Split(B, "g/L")(0))
    C = Trim(Left(C, 3))
    cgrdInput.TextMatrix(5, 2) = C
    
    '��������NIT
'ȡ����ֵ��0��12�ȣ��������ѡһ������ȡ���ֻ��ǽ������
'    B = Trim(Temp(4))
'    C = Trim(Split(B, "URO")(0))
'    cgrdInput.TextMatrix(6, 2) = C
    
'ȡ�����-��+��2+�ȣ�
    Temp = Split(A, "NIT")
    B = Trim(Temp(1))
    C = Trim(Left(B, 3))
    cgrdInput.TextMatrix(6, 2) = C

    'ǱѪ����Ѫ��BLD��ERY��
'ȡ����ֵ��0��12�ȣ��������ѡһ������ȡ���ֻ��ǽ������
'    B = Trim(Temp(7))
'    C = Trim(Split(B, "cells")(0))
'    cgrdInput.TextMatrix(8, 2) = C
    
'ȡ�����-��+��2+�ȣ�
    Temp = Split(A, "BLD")
    B = Trim(Temp(1))
    C = Trim(Split(B, "cells")(0))
    C = Trim(Left(C, 3))
    cgrdInput.TextMatrix(8, 2) = C
    
    '����Ѫ��VC��VitC��
'ȡ����ֵ��0��12�ȣ��������ѡһ������ȡ���ֻ��ǽ������
'    B = Trim(Temp(11))
'    C = Trim(Split(B, "mmol")(0))
'    cgrdInput.TextMatrix(11, 2) = C
    
'ȡ�����-��+��2+�ȣ�
    Temp = Split(A, "VC")
    B = Trim(Temp(1))
    C = Trim(Split(B, "mmol/L")(0))
    C = Trim(Left(C, 3))
    cgrdInput.TextMatrix(11, 2) = C

    '��ϸ����øLEU
'ȡ����ֵ��0��12�ȣ��������ѡһ������ȡ���ֻ��ǽ������
'    B = Trim(Temp(3))
'    C = Trim(Split(B, "cells")(0))
'    cgrdInput.TextMatrix(10, 2) = C
    
'ȡ�����-��+��2+�ȣ�
    Temp = Split(A, "LEU")
    B = Trim(Temp(1))
    C = Trim(Split(B, "cells")(0))
    C = Trim(Left(C, 3))
    cgrdInput.TextMatrix(10, 2) = C
    
    '����PH
    Temp = Split(A, "PH")
    B = Trim(Temp(1))
    C = Trim(Split(B, "ULD")(0))
    cgrdInput.TextMatrix(7, 2) = C
    '����SG
    Temp = Split(A, "SG")
    B = Trim(Temp(1))
    C = Trim(Split(B, "KET")(0))
    cgrdInput.TextMatrix(9, 2) = C
End Sub

'2012-07-14 �ڵ�� 2012-10-26 �����
Private Sub coptClasses_Click(Index As Integer)
    
    coptIndex = Index
    sub��������ģ��
    
    sub��ѯ�б���ʾ coptIndex
End Sub




Private Sub ctabPerson_Click(PreviousTab As Integer)
'    sub��ѯ�б���ʾ coptIndex
    Timer1.Enabled = True
End Sub







'Private Sub ctabPerson_Click(PreviousTab As Integer)
'    If PreviousTab = 0 Then
'        sub��������ģ��
'    End If
'End Sub

Private Sub ctxt����_Change()
    If Not IsNumeric(Trim(ctxt����.Text)) Then
        ctxt����.Text = ""
        Exit Sub
    End If
    
End Sub

Private Sub ctxt�������_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lstrNo As String
    Dim i As Integer
    Dim str���ҽ��� As String
    Dim lcolְҵ������ As Object
    Dim lobjRec As Object
    Dim strSQL, lstrItemNo As String
    
    lstrNo = Trim(ctxt�������.Text)
    
'''    coptClasses(0).Enabled = False
'''    coptClasses(1).Enabled = False
'''    coptClasses(2).Enabled = False
    
    '���������Ƿ����
    Dim mlobjRec As Object
    Dim lobjTmp As Object
    Set lobjTmp = CreateObject("ְҵ�������¼��.clscommon")
    Set mlobjRec = lobjTmp.func��ȡ�����Ա������Ϣ(lstrNo)
    If mlobjRec.RecordCount = 0 Then
        Set mlobjRec = Nothing
        
        '��յ�ǰ������Ϣ
        ctxt�������.Enabled = True
        ctxt����.Text = ""
        ctxt�Ա�.Text = ""
        ctxt����.Text = ""
        ctxt��λ����.Text = ""
        Exit Sub
    End If
    
    LoadPersonalInfoBatch (lstrNo)
    
    '2012-07-15 �ڵ�� ��
    '�ж��Ƿ���Բ�����ͼ����
    If InputFlag = "��ٿ�" Then
        ccmdEyeDraw.Visible = sub�Ƿ��л�ͼ��Ŀ(lstrNo)
    End If
   
    
    
    
    '2012-07-15 �ڵ�� ��
        
    Set lobjTmp = CreateObject("ְҵ�������¼��.clscommon")
    If lobjTmp.func��ȡ�����Ա��������Ϣ(lstrNo, priDeptName) Then
        Set lobjTmp = Nothing
       
        Set lcolְҵ������ = CreateObject("ְҵ������.clsManageMedicalExam")
        str���ҽ��� = lcolְҵ������.func���ؿ��ҽ���(ctxt�������.Text, priDeptName)
        ctxtConclun.Text = str���ҽ���
        
        'һ��ȷ����ǰ�����Ա��ţ��Ͳ��ܸ��ġ����ǣ���ս������ݡ�
        ctxt�������.Enabled = False
        ctxt����.Enabled = False
        ctxt�Ա�.Enabled = False
        ctxt����.Enabled = False
        ctxt��λ����.Enabled = False '��ʵ��λ�ҵ���֮������С���λ��λ����ť�����ǿ��Ըĵġ�
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
        MsgBox ("�������Աû�иÿ��ҵ������Ŀ��")
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
    '�����Ƚ���
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
        '���������롰'����
        KeyAscii = 0
    End If

End Sub

Private Sub Form_Load()
    Dim i As Long
    Dim lobjRec As Object
    On Error GoTo errHandler
    
    '��ʼʱ��ͽ���ʱ��  2016-4-1 by Ĳ��
    beforecdtpDate.Value = Now
    lastcdtpDate.Value = Now
    
    If mblnInUse Then Exit Sub
    '��ʾ���ȡ�
    frmProcess.proPercent.Max = 8
    frmProcess.Label1.Caption = "���ڳ�ʼ�����棬��ȴ�..."
    frmProcess.proPercent.Value = 1
    frmProcess.Show
    DoEvents
    Me.Caption = InputFlag & "���¼��"
    Label7.Caption = "˫��ѡ���У��ɸ��ĵ�����ۣ����Ϊ�쳣��"
    Frame6.FontSize = 10
    '���ô����Ѽ��ر�־��
    mblnInUse = True
    mstrȨ�ޱ�־ = True     'Ĭ����Ȩ��
    
    '�������ʱ�������Ѫ���棬���밴ť��ʾ   2015-12-4 by Ĳ��
    If InputFlag = "Ѫ���滯���" Then
    Command1.Visible = True
    End If
     '�������ʱ����������������Ӱ�ť��ʾ   2015-12-17 by Ĳ��
    If InputFlag = "������" Then
    Command2.Visible = True
    End If
     '�������ʱ��������򳣹棬�����ն˺���д�����ť��ʾ   2016-5 by Ĳ��
    If InputFlag = "�򳣹滯���" Then
        ccmdAutoFull(1).Visible = False
        ccmdAutoFull(2).Visible = False
        comd�����ն�.Visible = True
        Command3.Visible = True
    End If
    
    
    '����mobj����ͨ�ö��󣬳�ʼ����������
    Dim lcol������ As Collection
    Set lcol������ = New Collection
    With lcol������
        .Add "��ս���(&N)110"
        .Add "����"
        .Add "��������(&S)"
        .Add "���(&F)109"    '4  ������� 2016-3-7 by Ĳ��
        .Add "|"
        .Add "�˳�"
    End With
    Set mobj����ͨ�ö��� = New cls����ͨ�ö���
    With mobj����ͨ�ö���
        Set .Form = Me
        Set .c������ = ctbMain
        .subInitialize lcol������, ""
    End With
    frmProcess.proPercent.Value = 2
    DoEvents
    ctbMain.Buttons(2).Enabled = False
    ctbMain.Buttons(3).Enabled = False
    ctbMain.Buttons(4).Enabled = False   '���     2016-3-7 by Ĳ��
    ctbMain.Buttons(4).Visible = False
    '¼����ʱ��Ϊ��ҽ��������
    LabelDoctor.Caption = LabelDoctor.Caption & " " & um�û���
    
    
    '��ʼ����ѯ�б�
    cgrdSingleList.HighLight = flexHighlightWithFocus
    cgrdSingleList.SelectionMode = flexSelectionListBox
    cgrdSingleList.cols = 0
    With cgrdSingleList
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "ϵͳ���"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "����"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "�Ա�"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "����"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "������"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "�������"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "��λ����"
        '�޸��ˣ����� 2012.12.10
        'bug�ţ�0000018
        '˵���������������    ����
'        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "�Ƿ���д"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "��дʱ��"
        '2012.12.10    ����
        .AutoSize 0, .cols - 1, 0, 0
        .SelectionMode = flexSelectionListBox
    End With
    frmProcess.proPercent.Value = 3
    DoEvents
    '������ѯ�б��ʼ��
    cgrdInfoBatch.HighLight = flexHighlightWithFocus
    cgrdInfoBatch.SelectionMode = flexSelectionListBox
    cgrdInfoBatch.cols = 0
    With cgrdInfoBatch
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "ϵͳ���"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "����"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "�Ա�"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "����"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "������"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "�������"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "��λ����"
        '�޸��ˣ����� 2012.12.10
        'bug�ţ�0000018
        '˵���������������    ����
'        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "�Ƿ���д"
        .cols = .cols + 1: .TextMatrix(0, .cols - 1) = "��дʱ��"
        '2012.12.10    ����
        .AutoSize 0, .cols - 1, 0, 0
        .SelectionMode = flexSelectionListBox
    End With
    frmProcess.proPercent.Value = 4
    DoEvents
    '��ʼ�������¼������

    With cgrdInput
        .rows = 1
        .cols = 6
        .TextMatrix(0, 0) = "����"
        .ColWidth(0) = 700
        .TextMatrix(0, 1) = "����"
        .ColWidth(1) = 1600
        .TextMatrix(0, 2) = "�����"
        .ColWidth(2) = 1000
        .ColHidden(3) = True '��Ÿ������Ŀ��ö����Դ��
        .ColWidth(4) = 1200
        .TextMatrix(0, 4) = "��׼ֵ"
        .ColWidth(5) = 600
        .TextMatrix(0, 5) = "��λ"
        .ShowComboButton = True
        .Editable = True
    End With
   
    '2012-08-22 �ڵ�� ��
    '��Ӳ�ѯ�������ʷ��¼
    With cgrdInputHistory
        .rows = 1
        .cols = 6
        .TextMatrix(0, 0) = "����"
        .ColWidth(0) = 700
        .TextMatrix(0, 1) = "����"
        .ColWidth(1) = 1600
        .TextMatrix(0, 2) = "�����"
        .ColWidth(2) = 1000
        .ColHidden(3) = True '��Ÿ������Ŀ��ö����Դ��
        .ColWidth(4) = 1200
        .TextMatrix(0, 4) = "��׼ֵ"
        .ColWidth(5) = 600
        .TextMatrix(0, 5) = "��λ"
    End With
    '2012-08-22 �ڵ�� ��
    
    frmProcess.proPercent.Value = 5
    DoEvents
    '�����������ȫ�ֶ���mobj���ҽʦ��
    Set mobj���ҽʦ = CreateObject("ְҵ������.clsMedicalExaminer")
    mobj���ҽʦ.��� = um�û����
    
    '2012-06-21 �ڵ�� ��
    '����ϵͳ��Ź̶����֡�ʡ������Ҫ���иı�ϵͳ��Ź���
    '��ȡϵͳ��Ź̶����֡�
'''    Dim lobj��� As Object 'ְҵ�����󣬻�ȡϵͳ��ŵĹ̶����֡�
'''    Set lobj��� = CreateObject("ְҵ������.clsMedicalExam")
'''    mstrϵͳ��Ź̶����� = lobj���.ϵͳ��Ź̶�����
'''    ctxtSingleNo.Text = mstrϵͳ��Ź̶�����
'    sub��ȡϵͳ��Ź̶�����
    '2012-06-21 �ڵ�� ��
    frmProcess.proPercent.Value = 6
    DoEvents
    '���ҽʦ"��ʾ����ʾ��ǰ�û�����
    ctxtDoctor.Text = um�û���
    cdtpInputDate.Value = Date
    cdtpDateBatch.Value = Date
'    cdtpDate.Value = Now
    DTP¼������.Value = Date
    cgrdInput.rows = 1
    ctxtSingleNo.TabIndex = 0
'    If InputFlag = "��ٿ�" Then
'        Me.ccmdEyeDraw.Visible = True
'    Else
'        Me.ccmdEyeDraw.Visible = False
'    End If
    
    '2012-04-11
    '���水ť����
    ccmdAutoFull(0).Enabled = False
    ccmdAutoFull(1).Enabled = False
    ccmdAutoFull(2).Enabled = False
    ccmdAutoFull(3).Enabled = False
    Frame5.Enabled = False
    '2012-04-11
    
    '2012-05-22 ���� ������
    '����Ȩ������
    Dim lobjTmp As Object
    Set lobjTmp = CreateObject("ְҵ������.clsPermissionConfigure")
    If lobjTmp.func���Ҳ���Ȩ��(um�û����, "ְҵ�����_" & InputFlag & "���¼��_�޸�") = False Then
        ctbMain.Buttons(2).Visible = False
        mstrȨ�ޱ�־ = False
    End If
    
    If lobjTmp.func���Ҳ���Ȩ��(um�û����, "ְҵ�����_" & InputFlag & "���¼��_�����޸�") = False Then
        ctbMain.Buttons(3).Visible = False
        mstrȨ�ޱ�־ = False
    End If
    
'    '��һ�����Ȩ�ޣ���������Ȩ�ޣ��������ʾ   2016-3-7 by Ĳ��
'    If lobjTmp.func���Ҳ���Ȩ��(um�û����, "ְҵ�����_" & InputFlag & "���¼��_���") = True Then
'    ctbMain.Buttons(4).Visible = True
'    Check���.Visible = True    '�Ӹ�ѡ������ˡ�
'    cchkSigResult(0).Caption = "δ���"   '��ѡ���������������ʱΪ��δ��ˡ�
'    cchkSigResult(0).Value = 1    'ȱʡ��δ��ˡ�
'    cchkSigResult(1).Visible = False  '���ʱ��δ����������
'    mstrȨ�ޱ�־ = False
'    End If
    
    
    Set lobjTmp = Nothing
    frmProcess.proPercent.Value = 7
    DoEvents
    '2012-05-22 ������
    ctabPerson.Tab = 0
    '2012-06-21 �ڵ�� ��
    '��ʼ����ǰ¼��״̬(��ǰ�ж�����Ȩ���޸ģ����ޣ�ֱ�Ӹ�ֵΪ3)
    ResultChanged = IIf(ResultChanged <> 3, -1, 3)
    '2012-06-21 �ڵ�� ��
    
    '2012-07-14 �ڵ�� ��
    '��ʼ����ѯ���棬������ѯ�б��ʽ����ʼ�����һ�����Ϣ��
    ccmbHistory.Enabled = False
    priDeptName = InputFlag
    priDeptNo = InputFlagNo
    priDeptResultName = InputFlag
'    ccmdSingleQuery_Click
'    ctabPerson.Tab = 1: ccmdSelInfo_Click: ctabPerson.Tab = 0
    mstr���ͼƬ���� = "��״�廷�漰����ͼ"
    coptClasses_Click (0)
    '2012-07-14 �ڵ�� ��
    
    '�޸ģ�2001-12-29����ȡ��������ֵ����
    On Error Resume Next
    Set mobj���� = New cls�û���������
    mobj����.�û���� = um�û����
    mobj����.ҵ���� = "������"
    sub��������ģ��
    frmProcess.proPercent.Value = 8
    Unload frmProcess
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����沿��", "FrmResultInput_Routine", "Form_Load", 6666, lstrError, False
    Exit Sub
    Resume
End Sub

'����Ӧ���ڷֱ���
'2012-10-18 ������
Private Sub Form_Resize()
    On Error Resume Next
    Picture1.Width = Me.ScaleWidth - Picture1.Left
    Picture1.Height = Me.ScaleHeight - Picture1.Top - 20
    Frame1.Width = Picture1.Width - Frame1.Left
    Frame1.Height = Picture1.Height - Frame1.Top - 20
    
    ctbMain.Width = Frame1.Width - ctbMain.Left
    
    Frame5.Top = Frame1.Height - Frame5.Height - 50
    Frame8.Height = Frame5.Height
    Frame8.Top = Frame5.Top
    
    Frame6.Height = Frame1.Height - Frame6.Top - Frame5.Height - ccmdEyeDraw.Height - 100
    Frame7.Height = Frame6.Height
    
    ccmdEyeDraw.Top = Frame6.Top + Frame6.Height + 30
    Cmd����ģ��.Top = ccmdEyeDraw.Top
    cgrdInput.Height = Frame6.Height - cgrdInput.Top - 10
    cgrdInputHistory.Height = cgrdInput.Height
    
    ctabPerson.Height = Frame1.Height - ctabPerson.Top - 20
    Frame2.Height = ctabPerson.Height - Frame2.Top - 20
    cgrdSingleList.Height = Frame2.Height - cgrdSingleList.Top - 20
    cgrdInfoBatch.Height = ctabPerson.Height - cgrdInfoBatch.Top - 20
    
    '���
    Frame6.Width = (Frame1.Width - Frame6.Left - 40) * 3 / 5
    Frame7.Width = Frame6.Width * 2 / 3
    Frame7.Left = Frame6.Left + Frame6.Width + 10
    cgrdInput.Width = Frame6.Width - cgrdInput.Left * 2
    cgrdInputHistory.Width = Frame7.Width - cgrdInputHistory.Left * 2
    Frame5.Width = Frame6.Width
    Frame5.Left = Frame6.Left
    Frame8.Left = Frame7.Left
    Frame8.Width = Frame7.Width
    ctxtConclun.Width = Frame5.Width - ctxtConclun.Left * 2
    ctxtConclunHistory.Width = Frame8.Width - ctxtConclunHistory.Left * 2

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    '�޸ģ�2002-9-26����������������ֵ��
    'mobj����.sub���Ǽ���ֵ "¼����ʱˢ����", IIf(cchkˢ����.Value = 1, "��", "��")
    
    '�ͷű������ȫ�ֶ���
    Set mobj����ͨ�ö��� = Nothing
    Set mobj���ҽʦ = Nothing
    Set mobj���� = Nothing
    
    '���ñ�־pblnInUse�����������Ѳ���ʹ�á�
    mblnInUse = False

End Sub

Private Sub cchkˢ����_Click()
    If Not cchkˢ����.Visible Then Exit Sub
    If ctxtSingleNo.Enabled = False Then Exit Sub
    
    If ctabPerson.Tab = 0 Then
        ctxtSingleNo.Text = ""
        If cchkˢ����.Value = 0 Then sub��ȡϵͳ��Ź̶�����
        ctxtSingleNo.Enabled = True
        ctxtSingleNo.SetFocus
        ctxtSingleNo.SelStart = Len(ctxtSingleNo)
        ctxtSingleNo.SelLength = 0
    Else
        ctxt�������.Text = ""
        ctxt�������.SetFocus
    End If
End Sub

Private Sub cgrdInput_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim lstrEnum As String   '��ǰ�������ö����Դ����Ӣ�Ķ��Ż����Ķ��Ÿ�������
    Dim i As Long
    
    On Error GoTo errHandler
    'ֻ��������п���¼�롣
    If Col <> 2 Then
        Cancel = True
    Else
        '������������д�ŵ�ö����Դ���õ�ǰ��Ԫ�������б�
        lstrEnum = cgrdInput.TextMatrix(cgrdInput.Row, 3)
        If lstrEnum = "" Then
            'û��ö����Դ��
            cgrdInput.ColComboList(2) = ""
            
        Else
            '����������е�¼�뷽ʽΪ����ѡ��
            cgrdInput.ColComboList(2) = "|" & lstrEnum
'            cgrdInput.ColComboList(2) = lstrEnum
        End If
        
    End If
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����沿��", InputFlag & "���¼��", "cgrdInput_BeforeEdit", 6666, lstrError, False
    Exit Sub
    Resume
End Sub


'Private Sub ctxtSingleNo_GotFocus()
'    On Error Resume Next
'    '������ϵͳ��ţ�����ϵͳ��ŵĹ̶����֣�����¼�롣
'    If ctxtSingleNo.Text = "" Then
'        ctxtSingleNo.Text = mstrϵͳ��Ź̶�����
'        ctxtSingleNo.SelLength = 0
'        ctxtSingleNo.SelStart = Len(mstrϵͳ��Ź̶�����)
'        ctbMain.Buttons(1).Enabled = False
'
'    End If
'End Sub

Private Sub ctxtSingleNo_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errHandler
    Dim str���ҽ��� As String
    Dim lcolְҵ������ As Object
    Dim lstrNo As String
    
    lstrNo = Trim(ctxtSingleNo.Text)
    If KeyCode = 13 And lstrNo <> "" Then
        '��ʾ��Ա��Ϣ��
        subShowSinglePerson
        
        Set lcolְҵ������ = CreateObject("ְҵ������.clsManageMedicalExam")
        str���ҽ��� = lcolְҵ������.func���ؿ��ҽ���(lstrNo, priDeptName)

        '2016-1-5 ģ��Ĭ��δ���쳣   by Ĳ��
        ctxtConclun.Text = str���ҽ���
        If priDeptName = "B��Ӱ���" Then
            If Trim(ctxtConclun.Text) = "" Then
                ctxtConclun.Text = "���ࣺδ���쳣������" + Chr(13) + Chr(10) + "���ң�δ���쳣������" + Chr(13) + Chr(10) + "���٣�δ���쳣������" + Chr(13) + Chr(10) + "Ƣ�ࣺδ���쳣������" + Chr(13) + Chr(10) + "���ࣺδ���쳣������"
            Else
            ctxtConclun.Text = str���ҽ���
            End If
        Else
            If Trim(ctxtConclun.Text) = "" Then
                ctxtConclun.Text = "δ���쳣"
            Else
            ctxtConclun.Text = str���ҽ���
            End If
        End If

'        ctxtConclun.Text = str���ҽ���

        'ctbMain.Buttons(2).Enabled = True
        ctbMain.Buttons(3).Enabled = False
        If cgrdInput.rows > 1 Then ctbMain.Buttons(1).Enabled = True

        '2012-07-16 �ڵ�� ��
        '��ӿ����б���ʾ����
        With cgrdInput
            .Col = 0
            .Sort = flexSortGenericAscending
            .AutoSize 0, .cols - 1, 0, 0
            .ExplorerBar = flexExSort
            .DataMode = flexDMFree
            If .ColWidth(2) < 2000 Then
                .ColWidth(2) = 2000
            End If
        End With
        '2012-07-16 �ڵ�� ��

'             ' �ж��Ƿ������Ȩ�� 2016-3-8  by Ĳ��
'            If ctbMain.Buttons(4).Visible = True Then
'            cgrdInput.Enabled = False
'            ctbMain.Buttons(2).Enabled = False  '����ȥ��
'            ctbMain.Buttons(4).Enabled = True   '��˷ſ�
'            Else
'            cgrdInput.Enabled = True
'            End If


        '2012-07-15 �ڵ�� ��
        '�ж��Ƿ���Բ�����ͼ����
        ccmdEyeDraw.Visible = sub�Ƿ��л�ͼ��Ŀ(lstrNo)
        '2012-07-15 �ڵ�� ��
    End If
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����沿��", InputFlag & "���¼��", "ctxtSingleNo_KeyDown", 6666, lstrError, False
    Exit Sub
    Resume
End Sub


'���ܣ�����ϵͳ��ţ����������¼������

Private Sub subShowInputGrid(ByVal paraSysNo As String)
    Dim lcolInfo As Collection '��ŵ�ǰϵͳ����е�ǰҽʦ��������ָ�����Ե������Ŀ��������
    Dim lobjItem As Variant    'clsFactTestItem,lcolInfo�е�Ԫ�ء�
    Dim lstrEnum As String
    Dim i As Long
    Dim j As Long
    On Error GoTo errHandler
    
    'ʹ���Ż��㷨��ȡ�����Ŀ��
    Dim lobjRec As Object
    Dim lobjTemp As Object
    Dim mcolIndex As Collection
    '��ȡָ�����ԣ�����/���飩�������Ŀ��clsFactTestItem(�����Ŀ���룬�����Ŀ���ƣ�ȱʡֵ��ö����Դ�������)��
    '��ѡ�������������ȡ���������Ͽ�����Ŀ��
    Set lobjRec = mobj���ҽʦ.Func�Ż��Ļ�ȡ���˿����������Ŀ(paraSysNo, mstr�����Ŀ����, priDeptName)
    
    '��ʾ�����Ŀ��cgrdInput�С�
    cgrdInput.rows = 1
    
    Set mcol�����Ŀ = New Collection
    
    If Not (lobjRec.EOF Or lobjRec.BOF) Then
        '2012-06-21 �ڵ�� ��
        '���ĵ�ǰ¼��״̬
        If IsNull(lobjRec("�������")) Then
            ResultChanged = IIf(ResultChanged <> 3, 0, 3)
        Else
            ResultChanged = IIf(ResultChanged <> 3, 1, 3)
        End If
        '2012-06-21 �ڵ�� ��
        
        cgrdInput.rows = lobjRec.RecordCount + 1
        i = 1
        'If Not (lobjRec.EOF Or lobjRec.bof) Then
'            cgrdInput.TextMatrix(i, 0) = lobjRec!�����Ŀ���
'            cgrdInput.TextMatrix(i, 1) = lobjRec!�����Ŀ����
'            cgrdInput.TextMatrix(i, 2) = IIf(IsNull(lobjRec!�����), "", lobjRec!�����)
            Set cgrdInput.DataSource = lobjRec
            Set mcolIndex = New Collection
            
            For j = 0 To cgrdInput.cols - 1
                mcolIndex.Add j, cgrdInput.TextMatrix(0, j)
            Next
'            cgrdInput.AutoSize 0, cgrdInput.cols - 1, 0, 0
            '������
            cgrdInput.ColHidden(mcolIndex("�����Ŀ���")) = True
            cgrdInput.ColHidden(mcolIndex("ö����Դ")) = True
            cgrdInput.ColHidden(mcolIndex("ȱʡֵ")) = True
'            cgrdInput.ColHidden(mcolIndex("��׼ֵ")) = True
'            cgrdInput.ColHidden(mcolIndex("��λ")) = True
            cgrdInput.ColHidden(mcolIndex("�������")) = True
        'End If
        Do While Not lobjRec.EOF
            '���ݵ������������ɫ��
            Dim lstr������� As String
            If IIf(IsNull(lobjRec!�������), "", lobjRec!�������) = "" And cgrdInput.TextMatrix(i, 2) <> "" Then
                '�����µ�����ۡ�
                lstr������� = pobjҵ�����.func��ȡ�������(lobjRec!�����Ŀ���, IIf(IsNull(lobjRec!�����), "", lobjRec!�����))
            Else
                lstr������� = IIf(IsNull(lobjRec!�������), "", lobjRec!�������)
            End If
            If lstr������� = "���ϸ�" Then
                '������ɫ��
                cgrdInput.Cell(flexcpBackColor, i, 2, i, 2) = &H8A5AFA
            Else
                '������ɫ��
                cgrdInput.Cell(flexcpBackColor, i, 2, i, 2) = vbWhite
            End If
            '��ö����Դ��ת��ΪGrid����ʶ���ColComboList�����ԡ�|����������
            lstrEnum = IIf(IsNull(lobjRec!ö����Դ), "", lobjRec!ö����Դ)
            lstrEnum = gffuncStrReplace(lstrEnum, ",", "|")
            lstrEnum = gffuncStrReplace(lstrEnum, "��", "|")
            cgrdInput.TextMatrix(i, 3) = lstrEnum

            cgrdInput.TextMatrix(i, 4) = IIf(IsNull(lobjRec!��׼ֵ), "", lobjRec!��׼ֵ)
            cgrdInput.TextMatrix(i, 5) = IIf(IsNull(lobjRec!��λ), "", lobjRec!��λ)

            '(Ϊ�˽�������¼�룬��cgrdperson������������ʾ�����Ŀ����)��
'            If ctabPerson.Tab = 1 Then
'                For j = mintFixed + 1 To cgrdPerson.cols - 1
'                    If cgrdPerson.TextMatrix(0, j) = lobjRec!�����Ŀ���� Then Exit For
'                Next
'                If j = cgrdPerson.cols Then
'                    cgrdPerson.cols = cgrdPerson.cols + 1
'                    cgrdPerson.TextMatrix(0, j) = lobjRec!�����Ŀ����
'
'                    'cgrdPerson.ColHidden(j) = IIf(cchkGrid.Value = 0, True, False)
'                    If lstrEnum = "" Then
'                        cgrdPerson.ColComboList(j) = ""
'                    Else
'                        cgrdPerson.ColComboList(j) = "|" & lstrEnum
'                    End If
'                End If
'
'                mcol�����Ŀ.Add lobjRec("�����Ŀ���").Value, lobjRec!�����Ŀ����
'            End If
            i = i + 1
            lobjRec.MoveNext
        Loop
       
        '2012-07-16 �ڵ�� ��
        '��ӿ����б���ʾ����
        With cgrdInput
            .Col = 0
            .Sort = flexSortGenericAscending
            .AutoSize 0, .cols - 1, 0, 0
            .ExplorerBar = flexExSort
            .DataMode = flexDMFree
        End With
        '2012-07-16 �ڵ�� ��
        cgrdInput.Select 1, 2, 1, 2
    Else
        cgrdInput.rows = 1
        
        Err.Raise 6666, , "�Բ��𣬸������Ա�����ϵ�����" & mstr�����Ŀ���� & "�����Ŀ���㶼�����Բ���������������Ա��ʹ�õ���������û������" & mstr�����Ŀ���� & "��Ŀ�������ҵ�����õġ����ҽʦ���á������ɲ�������Ŀ�������롰�������á�������������õ���Ŀ��"
    End If
    
    Exit Sub
errHandler:
    sfsub������ "ְҵ�����沿��", InputFlag & "���¼��", "subShowInputGrid", Err.Number, Err.Description, True
    Exit Sub
    Resume
End Sub

'2012-08-22 �ڵ��
'���ܣ�����ϵͳ��ţ������������ʷ��¼¼������
Private Sub subShowInputGridHistory(ByVal paraSysNo As String)
    Dim lcolInfo As Collection '��ŵ�ǰϵͳ����е�ǰҽʦ��������ָ�����Ե������Ŀ��������
    Dim lobjItem As Variant    'clsFactTestItem,lcolInfo�е�Ԫ�ء�
    Dim lstrEnum As String
    Dim i As Long
    Dim j As Long
    On Error GoTo errHandler
    
    'ʹ���Ż��㷨��ȡ�����Ŀ��
    Dim lobjRec As Object
    Dim lobjTemp As Object
    
    '��ȡָ�����ԣ�����/���飩�������Ŀ��clsFactTestItem(�����Ŀ���룬�����Ŀ���ƣ�ȱʡֵ��ö����Դ�������)��
    '��ѡ�������������ȡ���������Ͽ�����Ŀ��
    Set lobjRec = mobj���ҽʦ.Func�Ż��Ļ�ȡ���˿����������Ŀ(paraSysNo, mstr�����Ŀ����, priDeptName)
    
    '��ʾ�����Ŀ��cgrdInputHistory�С�
    cgrdInputHistory.rows = 1
    
    Set mcol�����Ŀ = New Collection
    
    If lobjRec.RecordCount > 0 Then
        '2012-06-21 �ڵ�� ��
        '���ĵ�ǰ¼��״̬
        If IsNull(lobjRec("�������")) Then
            ResultChanged = IIf(ResultChanged <> 3, 0, 3)
        Else
            ResultChanged = IIf(ResultChanged <> 3, 1, 3)
        End If
        '2012-06-21 �ڵ�� ��
        
        cgrdInputHistory.rows = lobjRec.RecordCount + 1
        i = 1
        Do While Not lobjRec.EOF
            cgrdInputHistory.TextMatrix(i, 0) = lobjRec!�����Ŀ���
            cgrdInputHistory.TextMatrix(i, 1) = lobjRec!�����Ŀ����
            cgrdInputHistory.TextMatrix(i, 2) = IIf(IsNull(lobjRec!�����), "", lobjRec!�����)
            
            '���ݵ������������ɫ��
            Dim lstr������� As String
            If IIf(IsNull(lobjRec!�������), "", lobjRec!�������) = "" And cgrdInputHistory.TextMatrix(i, 2) <> "" Then
                '�����µ�����ۡ�
                lstr������� = pobjҵ�����.func��ȡ�������(lobjRec!�����Ŀ���, IIf(IsNull(lobjRec!�����), "", lobjRec!�����))
            Else
                lstr������� = IIf(IsNull(lobjRec!�������), "", lobjRec!�������)
            End If
            If lstr������� = "���ϸ�" Then
                '������ɫ��
                cgrdInputHistory.Cell(flexcpBackColor, i, 2, i, 2) = &H8A5AFA
            Else
                '������ɫ��
                cgrdInputHistory.Cell(flexcpBackColor, i, 2, i, 2) = vbWhite
            End If
            '��ö����Դ��ת��ΪGrid����ʶ���ColComboList�����ԡ�|����������
            lstrEnum = IIf(IsNull(lobjRec!ö����Դ), "", lobjRec!ö����Դ)
            lstrEnum = gffuncStrReplace(lstrEnum, ",", "|")
            lstrEnum = gffuncStrReplace(lstrEnum, "��", "|")
            cgrdInputHistory.TextMatrix(i, 3) = lstrEnum
            cgrdInputHistory.TextMatrix(i, 4) = IIf(IsNull(lobjRec!��׼ֵ), "", lobjRec!��׼ֵ)
            cgrdInputHistory.TextMatrix(i, 5) = IIf(IsNull(lobjRec!��λ), "", lobjRec!��λ)

            i = i + 1
            lobjRec.MoveNext
        Loop
        
        '2012-07-16 �ڵ�� ��
        '��ӿ����б���ʾ����
        With cgrdInputHistory
            .Col = 0
            .Sort = flexSortGenericAscending
            .AutoSize 0, .cols - 1, 0, 0
            .ExplorerBar = flexExSort
            .DataMode = flexDMFree
        End With
        '2012-07-16 �ڵ�� ��
        cgrdInputHistory.Select 1, 2, 1, 2
    Else
        cgrdInputHistory.rows = 1
        Err.Raise 6666, , "�Բ��𣬸������Ա�����ϵ�����" & mstr�����Ŀ���� & "�����Ŀ���㶼�����Բ���������������Ա��ʹ�õ���������û������" & mstr�����Ŀ���� & "��Ŀ�������ҵ�����õġ����ҽʦ���á������ɲ�������Ŀ�������롰�������á�������������õ���Ŀ��"
    End If
    
    Exit Sub
errHandler:
    sfsub������ "ְҵ�����沿��", InputFlag & "���¼��", "subShowInputGrid", Err.Number, Err.Description, True
    Exit Sub
    Resume
End Sub

'���ܣ����ݵ�������ʽ����ı����ʾ�����Ա��Ϣ�������Ŀ������������������ʽ����ĵ�����Ż�ȡ�����Ա��Ϣ�����������У���ʾ�����Ŀ��������
Private Sub subShowSinglePerson()
    On Error GoTo errHandler
    Dim lobj��� As Object     'ְҵ������
    Dim lobj��켯 As Object   '��켯�������ڸ����Թܱ��+���ڻ�ȡϵͳ��š�
    Dim lobjRec As Object
    
    Dim lstrNo As String       'ϵͳ��Ż��Թܱ�š�
    Dim llngNoType As Long     '������ͣ�0 ϵͳ���/1 �Թܱ�š�
    Dim lstrSysNo As String    'ϵͳ��š�
    Dim i As Long
    
    
    '��ȡ�����ϵͳ��ţ����Թܱ�ţ���
    lstrNo = Trim(ctxtSingleNo.Text)
    
    If lstrNo <> "" Then
        '����ְҵ������
        Set lobj��� = CreateObject("ְҵ������.clsMedicalExam")
        
        lobj���.ϵͳ��� = Trim(ctxtSingleNo.Text)
        
        lstrSysNo = lobj���.ϵͳ���
        'ctxtSingleNo.Text = lstrSysNo
        
        '��ս��档
        If ctabPerson.Tab = 0 Then
            ctxtName = ""
            ctxtSex = ""
            ctxtAge = ""
            ctxtCompanyName = ""
            cpicPhoto(0).Picture = Nothing
        End If
        
        '�ж��Ƿ���ڡ�
        If Not lobj���.�Ƿ��Ѵ��� Then
            Err.Raise 6666, , "�������������ŵ������Ա�����������롣"
        End If
        
        '�ж��Ƿ����������ۡ�
        'If lobj���.���״̬ = P_ENDED_STATUS Then
        '    Err.Raise 6666, , "�������ŵ�����ѱ�ҽʦȷ���������ۣ��������޸������������۵��������" & Chr(13) & Chr(10) & "��ȷʵҪ�޸ģ����½��۵�ҽʦ���롰������¼�롱����������ȡ���½��ۣ��ٻص��˲��������޸ġ�"
        'End If
        
        '��ʾ��Ա��Ϣ��������Ƭ����
        If ctabPerson.Tab = 0 Then
            '��������ʽ��
            With lobj���.�����Ա
                .����������� = lstrSysNo
                ctxtName = .����
                ctxtSex = .�Ա�
                ctxtAge = .����
                ctxtCompanyName = .��λ����
                ctxtweihaiyinsu = .Σ������
                '2012-04-11
                '��ʾ��Ա��Ϣ���ܽ����޸�
                ctxtSingleNo.Enabled = False
                ctxtName.Enabled = False
                ctxtSex.Enabled = False
                ctxtAge.Enabled = False
                ctxtCompanyName.Enabled = False
                '���水ť�ܹ����в���
                ccmdAutoFull(0).Enabled = True
                ccmdAutoFull(1).Enabled = True
                ccmdAutoFull(2).Enabled = True
                ccmdAutoFull(3).Enabled = True
                Frame5.Enabled = True
                '2012-04-11
                
                If llngNoType = 1 Then 'ϵͳ������뷽ʽ����Ҫ��ʾ�Թܱ�š�
                '    clblInfo(4) = lobj���.��쵥��
                    Label1(8).Caption = "��쵥�ţ�"
                Else
                    'clblInfo(4) = lobj���.�Թܱ��
                    'Label1(8).Caption = "�Թܱ�ţ�"
                End If
                
                '��ʾ��Ƭ��
                If Not .��Ƭ Is Nothing Then
                    cpicPhoto(0).Picture = .��Ƭ
                End If
            End With
            
            '��ʾ���˵����겡���������ǣ�2012-07-31��������������������������
            Dim lobjDatecobo As Object
            Set lobjDatecobo = mobj���ҽʦ.func��ȡ�����Ա����첡��(Trim(ctxtSingleNo.Text), InputFlag)
            If Not lobjDatecobo Is Nothing Then
            
                If Not (lobjDatecobo.EOF Or lobjDatecobo.BOF) Then
                    Label3.Visible = True
                    ccmbHistory.Visible = True
                    ccmbHistory.Clear
                    ccmbHistory.AddItem "����"
                    For i = 1 To lobjDatecobo.RecordCount
                        ccmbHistory.AddItem Format(lobjDatecobo("��дʱ��"), "yyyy-mm-dd")
    '                    ccmbHistory.AddItem
                        lobjDatecobo.MoveNext
                    Next i
                    ccmbHistory.Enabled = True
                Else
                    ccmbHistory.Clear
                    ccmbHistory.Enabled = False
                    
    '                Chk����ģ��.Visible = True
                    Cmd����ģ��.Visible = True
                    Frame5.Visible = True
    '                Frame6.Height = Frame6.Height - 300
    '                cgrdInput.Height = cgrdInput.Height - 300
                    cgrdInput.Enabled = True
                    cgrdInput.rows = 1
                    
                End If
            Else
                ccmbHistory.Clear
                ccmbHistory.Enabled = False
                ctxtConclunHistory.Text = ""
                cgrdInputHistory.Clear
                cgrdInputHistory.rows = 1
            End If
'            ccmbHistory.ListIndex = 0
            
            '��ʾ���˵����겡���������ǣ�2012-07-31�� ������������������������
            
            '���������¼������
            subShowInputGrid lstrSysNo
            
            cgrdSingleList.Row = 0
            
        Else
            '�޸ģ�2001-11-2�������Ϊֻ��ѯָ�����������¼�����Բ����ж������Ƿ���ͬ��
            
            '��������ʽ������Ա��Ϣ���뵽cgrdPerson�У�ע������������ͬ�ģ���
'            If cgrdPerson.rows = 1 Then
''                '��cgrdPerson��ԭû�м�¼������mstr�������ơ�
''                mstr�������� = lobj���.����.������
''
''                '���������¼������
'                subShowInputGrid lstrSysNo
'            'Else
''                '�ж������Ա���������Ƿ�һ�¡�
'                '�޸ģ�2002-8-14������������ѡ�����С�
'                'If ccmbSheet.Text <> "<����>" Then
'                '    If ccmbSheet.Text <> lobj���.����.������ Then
'                '        Err.Raise 6666, , "����������������" & lobj���.����.������ & "����ָ������һ�£���������¼��������ͬ���������"
'                '    End If
'                'End If
'            End If
'
'            '�жϸ���Ա�Ƿ����������У�����������Լ�������
'            For i = 1 To cgrdPerson.rows - 1
'                If cgrdPerson.TextMatrix(i, 0) = lstrSysNo Then
'                    '���������д��ڣ����ټ��롣
'                    Exit Sub
'                End If
'            Next
'
'            '����Ա��ӵ������Ա�����С�
'            cgrdPerson.rows = cgrdPerson.rows + 1
'
'            i = cgrdPerson.rows - 1
'            cgrdPerson.TextMatrix(i, 0) = lstrSysNo
'
'            '�޸ģ�2002-10-11�����������ʾ�Թܱ�š�
'            cgrdPerson.TextMatrix(i, 1) = lobj���.�Թܱ��
'            With lobj���.�����Ա
'                cgrdPerson.TextMatrix(i, 2) = .����
'                cgrdPerson.TextMatrix(i, 3) = .�Ա�
'                cgrdPerson.TextMatrix(i, 4) = .��λ����
'                cgrdPerson.TextMatrix(i, 5) = .����
'            End With
            
        End If
        
    End If

    ctbMain.Buttons(1).Enabled = True
    ctbMain.Buttons(2).Enabled = True
    ctbMain.Buttons(3).Enabled = False
    'cstbMain.Panels(1) = ""
    'cgrdInput.row = 1      '''''2012-07-04 �ڵ�� ��ʱע�ͣ�����ԭ�򣬲�����Ա�����Ŀ����ʱ����
    cgrdInput.Col = 2
    cgrdInput.SetFocus
    SendKeys ""
    
    '2012-07-03 �ڵ�� ��
    'ÿ�ζ��������Ϣʱ���ж��Ƿ񳬹��޸�ʱ�䡣
    '�Դ˿��Ʊ��水ť�Ƿ���á�
    If pobjҵ�����.sub�Ƿ����޸�ʱ�䷶Χ��(ctxtSingleNo.Text, priDeptName, 8) = False Then
        ctbMain.Buttons(2).Enabled = False
    End If
    '2012-07-03 �ڵ��
    Exit Sub
errHandler:
    If ctabPerson.Tab = 0 Then
        ctbMain.Buttons(1).Enabled = False
    End If
    sfsub������ "ְҵ�����沿��", InputFlag & "���¼��", "subShowSinglePerson", Err.Number, Err.Description, True
    Exit Sub
    Resume
End Sub




'���ܣ���һ����ָ����ŷ�Χ����������ڣ��������Ա�������񣬲���ʾ�����Ŀ��������
'Private Sub subShowBatchPerson()
'    Dim lobjRec As Object        'ͨ��ҵ������ȡ��ָ����Χ�ڿ���¼�������������¼��
'    Dim llngNoType  As Integer   '��ŷ�ʽ��0ϵͳ���/1�Թܱ�š�
'    Dim llngStartRow As Long     '��ǰ�����Ա����������+1��
'    Dim llngRow As Long          '��ǰ��ӵ��С�
'    Dim i As Long
'    Dim lobjResult As Object
'
'    On Error GoTo errHandler
'    cstbMain.Panels(1) = "���ڻ�ȡ����¼�����Ժ�..."
'
'    '��ȡ������͡�
'    llngNoType = 0 'ϵͳ��š�
'
'
'    '�������������ڡ�
'    '�޸ģ�2001-11-2�����Ӳ�ѯ�������������ƣ���
'    '�޸ģ�2002-8-14������������ѡ�����С�
'    Set lobjRec = pobjҵ�����.Func��ȡ���޸ĵ�����¼(IIf(coptBatchType(2).Value, ctxtBatchNo, ""), IIf(coptBatchType(2).Value, ctxtBatchNo, ""), IIf(coptBatchType(0).Value, Str(cdtpQueryDate.Value), ""), llngNoType, "", IIf(coptBatchType(1).Value, ctxt��λ����.Text, ""))
'
'    If lobjRec.RecordCount > 0 Then
'
'        lobjRec.Filter = ""
'        If cchkUnEnd(1).Value = 1 And cchkEnd(1).Value = 0 Then
'            lobjRec.Filter = "���״̬<>2"
'        ElseIf cchkUnEnd(1).Value = 0 And cchkEnd(1).Value = 1 Then
'            lobjRec.Filter = "���״̬=2"
'        ElseIf cchkUnEnd(1).Value = 0 And cchkEnd(1).Value = 0 Then
'            lobjRec.Filter = "���״̬=-1"
'        End If
'
'        cgrdPerson.Redraw = False
'        mblnSys = True
'        mintFixed = 6
'        If cgrdPerson.rows = 1 And lobjRec.RecordCount > 0 Then
'            '�޸ģ�2001-11-2���������Ҫ�ж��������ơ�
'            '���������¼�������С�
'            subShowInputGrid lobjRec!ϵͳ���
'        End If
'
'        '��ʾ��Ա��Ϣ��cgrdPerson�У�ע������������ͬ�ģ���
'        llngStartRow = cgrdPerson.rows - 1
'        Do While Not lobjRec.EOF
'            '�޸ģ�2001-11-2���������Ҫ�ж��������ơ�
'    '        If lobjRec!�������� = mstr�������� Then
'                '�жϸ���Ա�Ƿ����������У�����������Լ�������
'                For i = 1 To llngStartRow
'                    If cgrdPerson.TextMatrix(i, 0) = lobjRec!ϵͳ��� Then
'                        '���������д��ڣ����ټ��롣
'                        GoTo LabelContinue
'                    End If
'                Next
'                cgrdPerson.AddItem ""
'                llngRow = cgrdPerson.rows - 1
'                With cgrdPerson
'                    .TextMatrix(llngRow, 0) = lobjRec!ϵͳ���
'
'                    .TextMatrix(llngRow, 1) = IIf(IsNull(lobjRec!�Թܱ��), "", lobjRec!�Թܱ��)
'                    .TextMatrix(llngRow, 2) = IIf(IsNull(lobjRec!����), "", lobjRec!����)
'                    .TextMatrix(llngRow, 3) = IIf(IsNull(lobjRec!�Ա�), "", lobjRec!�Ա�)
'                    .TextMatrix(llngRow, 4) = IIf(IsNull(lobjRec!��λ����), "", lobjRec!��λ����)
'                    .TextMatrix(llngRow, 5) = IIf(IsNull(lobjRec!����), "", lobjRec!����)
'                    .TextMatrix(llngRow, 6) = IIf(IsNull(lobjRec!��쵥��), "", lobjRec!��쵥��)
'
'                    If lobjRec!���״̬ = 2 Then
'                        .Cell(flexcpBackColor, llngRow, 0, llngRow, mintFixed) = cchkEnd(1).BackColor
'                    Else
'                        .Cell(flexcpBackColor, llngRow, 0, llngRow, mintFixed) = cchkUnEnd(1).BackColor
'                    End If
'
'                    '2006-6-19(����¼�룩
'                    'If cchkGrid.Value = 1 Then
'                        '��ȡ���˵������������
'                        subShowPersonResult llngRow, lobjRec!ϵͳ���
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
'    sfsub������ "ְҵ�����沿��", "FrmInputTestResult", "subShowBatchPerson", Err.Number, Err.Description, True
'    mblnSys = False
'    Exit Sub
'    Resume
'End Sub

'Private Sub subShowPersonResult(ByVal paraRow As Long, ByVal paraϵͳ��� As String)
'    Dim i As Long
'    Dim lobjResult As Object
'
'
'    Set lobjResult = pobjҵ�����.func��ȡ�����(paraϵͳ���)
'    Do While Not lobjResult.EOF
'        For i = mintFixed + 1 To cgrdPerson.cols - 1
'            If cgrdPerson.TextMatrix(0, i) = lobjResult!�����Ŀ���� Then
'                cgrdPerson.TextMatrix(paraRow, i) = IIf(IIf(IsNull(lobjResult!�����), "", lobjResult!�����) = "", lobjResult!ȱʡֵ, lobjResult!�����)
'                '������ɫ��
'                If IIf(IsNull(lobjResult!�������), "", lobjResult!�������) = "���ϸ�" Then
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


Private Sub mobj����ͨ�ö���_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    On Error GoTo errHandler
    Dim lcolNo As Collection     'ϵͳ��ż��ϡ�
    Dim lcolResult As Collection '��������ϣ�item:[�����Ŀ�������]��
    Dim lcolItem As Collection   '���������Ŀ���������[�����Ŀ�������]��
    Dim lcolDetail As Collection
    Dim lblnNotOver As Boolean
    Dim lcolConclusion As String '���������Ŀ��������
    Dim i As Long
    Dim j As Long
    
    Cancel = True
    Select Case Operate
    Case "��ս���"
        subClear
    Case "��������"
        '2012-07-13 �ڵ�� ��
        '���û�������Ŀ����ֱ���˳��������档
        If cgrdInfoBatch.rows <= 1 Then Exit Sub
        '2012-07-13 �ڵ�� ��
        
        '2012-07-15 �ڵ�� ��
        'û��¼��������ʱ����ʾ�Ҳ����档
        If Len(Trim(ctxtConclun.Text)) = 0 Then
            MsgBox "��¼����������Ϣ��", vbExclamation, "ϵͳ��ʾ"
            GoTo errHandler
        End If
        '2012-07-15 �ڵ�� ��
        
        sub��������
        
        '2012-07-15 �ڵ�� ��
        '������֮�����½��в�ѯ��
        ccmdSingleQuery_Click
        i = ctabPerson.Tab
        ctabPerson.Tab = 1: ccmdSelInfo_Click: ctabPerson.Tab = i
        '2012-07-15 �ڵ�� ��
        
    Case "����"
        '2012-07-03 �ڵ�� ��
        '�ж��Ƿ����޸�ʱ�䷶Χ��
        If pobjҵ�����.sub�Ƿ����޸�ʱ�䷶Χ��(Trim(ctxtSingleNo.Text), priDeptName, 8) = False Then
            '���� 2012.12.28   ����
            '˵����X��Ӱ���48Сʱ���ڿ����޸ġ�
            If frmResultInput_Routine.Caption = "X��Ӱ��ƽ��¼��" Then
                MsgBox ("���ϴ��޸��Ѿ�����48Сʱ���������Ա��ϵ����޸�Ȩ�޺��ټ�����")
            Else
                MsgBox ("���ϴ��޸��Ѿ�����8Сʱ���������Ա��ϵ����޸�Ȩ�޺��ټ�����")
            End If
            '���� 2012.12.28  ����
            Exit Sub
        End If
        '2012-07-03 �ڵ�� ��
        
        MousePointer = 11
        'cstbMain.Panels(1) = "���ڱ��棬���Ժ�..."
        
        '��ʱ���治�ܲ�����
        ctbMain.Enabled = False
        ctabPerson.Enabled = False
        Frame1.Enabled = False
        cgrdInput.Select 1, 0, 1, 0
        
        Set lcolNo = New Collection
        Set lcolResult = New Collection
        
        '��ȡ��������ϣ�[�����Ŀ�������]��
        lblnNotOver = False
                
        If ctabPerson.Tab = 0 Then
            '���ǵ���¼�뷽ʽ����ctxtSingleNo.text��Ӧ��ϵͳ��ż��뼯��lcolNo�У�
            '�������ϵͳ��š�
            lcolNo.Add ctxtSingleNo.Text
        Else
            '��������¼�뷽ʽ����cgrdPerson�����������е�ϵͳ��ż���lcolNo�С�
'            For i = 1 To cgrdPerson.rows - 1
'                lcolNo.Add cgrdPerson.TextMatrix(i, 0)
'
'                '2006-6-19(����¼��)
'                Set lcolDetail = New Collection
''                'If cchkGrid.Value = 1 Then
''
''                    For j = mintFixed + 1 To cgrdPerson.cols - 1
''                        Set lcolItem = New Collection
''                        lcolItem.Add mcol�����Ŀ(cgrdPerson.TextMatrix(0, j)), "�����Ŀ"
''                        lcolItem.Add cgrdPerson.TextMatrix(i, j), "�����"
''
''                        If cgrdPerson.TextMatrix(i, j) = "" Then
''                            lblnNotOver = True
''                            lcolItem.Add "", "�������"
''                        ElseIf cgrdPerson.Cell(flexcpBackColor, i, j, i, j) = &H8A5AFA Then
''                            lcolItem.Add "���ϸ�", "�������"
''                        Else
''                            lcolItem.Add "�ϸ�", "�������"
''                        End If
''
''                        lcolDetail.Add lcolItem, lcolItem("�����Ŀ")
''                    Next
''                    lcolResult.Add lcolDetail, cgrdPerson.TextMatrix(i, 0)
''                End If
'            Next
        End If
        
        If lcolNo.Count = 0 Then
            Err.Raise 6666, , "��ѡ�������Ա����¼����������ٰ������桱��"
        End If
        
        If ctabPerson.Tab = 0 Then
            
            For i = 1 To cgrdInput.rows - 1
                Set lcolItem = New Collection
                lcolItem.Add cgrdInput.TextMatrix(i, 0), "�����Ŀ"
                lcolItem.Add cgrdInput.TextMatrix(i, 2), "�����"
                
                '��¼û��¼�ꡣ
                If cgrdInput.TextMatrix(i, 2) = "" Then
                    lblnNotOver = True
                    lcolItem.Add "", "�������"
                ElseIf Trim(cgrdInput.TextMatrix(i, 2)) = "�쳣" Then
                    lcolItem.Add "���ϸ�", "�������"
                Else
                    lcolItem.Add "�ϸ�", "�������"
                End If
                lcolResult.Add lcolItem, lcolItem("�����Ŀ")
            Next
        End If
        
        '��û��¼�꣬������ʾ��
        If lblnNotOver Then
            If mstrȨ�ޱ�־ = True Then
                If Not sffuncMsg("��û��¼�����������Ŀ����������Ƿ���Ҫ���棿", sfѯ��) Then
                    '�û�ѡ�񲻱��档
                    GoTo errHandler
                End If
            End If
        End If
        
        '���ƽ���¼�������
        If Len(Trim(ctxtConclun.Text)) >= 250 Then
            Err.Raise -2147217833, "�������ݹ���������󣩣��ѳ���ϵͳ�涨���ȣ����С����"
        End If
        
        If Len(Trim(ctxtConclun.Text)) > 0 Then
            lcolConclusion = ctxtConclun.Text
            '���浥����Ŀ��ҽ������
            pobjҵ�����.sub������д������ lcolNo.Item(1), priDeptName, lcolConclusion, um�û����
            
            '2012-07-03 �ڵ�� ��
            '����һ���ֶ�"�޸���ʼʱ��"���޸ġ�ͬʱ�޸ĸÿ��ҵ������¼��״̬��
            pobjҵ�����.sub�޸���ʼʱ�� Trim(ctxtSingleNo.Text), priDeptName
            pobjҵ�����.sub�޸Ľ��¼��״̬ Trim(ctxtSingleNo.Text), priDeptNo, "2"
            pobjҵ�����.sub���¼���޸����״̬ Trim(ctxtSingleNo.Text), "4"
            '2012-07-03 �ڵ�� ��
        Else
            MsgBox "����д���������Ϣ��", vbExclamation, "ϵͳ��ʾ"
            GoTo errHandler
        End If
        
        'ʹ���Ż����㷨�����������
        If ctabPerson.Tab = 0 Then
            pobjҵ�����.Sub�Ż�����д����� lcolNo, lcolResult, um�û����, cdtpInputDate.Value
        Else
            pobjҵ�����.Sub������д����� lcolNo, lcolResult, um�û����, cdtpInputDate.Value
        End If
        
        '�ָ����档
        ctbMain.Buttons(1).Enabled = False
        ctbMain.Enabled = True
        ctabPerson.Enabled = True
        Frame1.Enabled = True
        
        '��ս��档
        subClear
        
        '2012-07-15 �ڵ�� ��
        '������֮�����½��в�ѯ��
'        ccmdSingleQuery_Click      '����֮�󲻽������²�ѯ  2015-12-25 by Ĳ��
        i = ctabPerson.Tab
        ctabPerson.Tab = 0: ccmdSingleQuery_Click: ctabPerson.Tab = i
        '2012-07-15 �ڵ�� ��
        
        MousePointer = 0
        If MSComm1.PortOpen = True Then
            MSComm1.PortOpen = False    '�����ص��ն�����  2016-7-6 by Ĳ��
        End If
        MsgBox "����ɹ�"    '��ʾ��Ϊ�˷����û�֪���ǵ��ˡ����桱�����������������ť 2016-4-1
        'cstbMain.Panels(1) = "����ɹ���"
    '2012-06-21 �ڵ�� ��
    '�˳�ʱ�����ж��Ƿ񱣴�
    
'    ' ��ˣ��޸����״̬�͸��������״̬����ʱ�Ȳ��޸�״̬�����Ǹ�������ɵ�״̬һ����   2016-3-7 by Ĳ��
'    Case "���"
'        pobjҵ�����.sub�޸���ʼʱ�� Trim(ctxtSingleNo.Text), priDeptName
'        pobjҵ�����.sub�޸Ľ��¼��״̬ Trim(ctxtSingleNo.Text), priDeptNo, "2"
'        pobjҵ�����.sub���¼���޸����״̬ Trim(ctxtSingleNo.Text), "4"
        
    
    Case "�˳�"
        ctxtSingleNo.Enabled = True
        ctxtSingleNo.SetFocus
        ctxtSingleNo.Enabled = False
        Dim isSave As Integer
        
        If Not mstrȨ�ޱ�־ Then
            Unload Me
            Exit Sub
        End If
        If ResultChanged = 2 Or ResultChanged = 0 Then
            '�޸ģ�������ڲ����鿴�����˳������ѡ������ǣ�2012-08-01��
            If Trim(Frame6.Caption) <> "�����Ŀ�����д��" Then
                Unload Me
                Exit Sub
            End If
            isSave = MsgBox("�Ƿ񱣴����޸Ľ����", vbYesNoCancel)
            If isSave = vbCancel Then Exit Sub
            If isSave = vbNo Then
                mobj����ͨ�ö���_BeforeOperate "��ս���", True
                '�޸��ˣ����� 2012.12.07
                'bug�ţ�0000022
                '˵�������˳���ʾ��񲻱���ʱֱ�ӹرմ��ڡ�����
'                Exit Sub
                Unload Me
                '2012.12.07    ����
            End If
            If isSave = vbYes Then
                '�޸��ˣ����� 2012.12.06
                'bug�ţ�0000037
                '˵�����ж��Ƿ�Ϊ�������档����
                If ctabPerson.Tab = 0 Then
                    mobj����ͨ�ö���_BeforeOperate "����", False
                Else
                    mobj����ͨ�ö���_BeforeOperate "��������", False
                End If
                '2012.12.06  ����
                '�޸��ˣ����� 2012.12.06
                'bug�ţ�0000015
                '˵������������������˳��᲻��ʾ��  ����
'                mstrȨ�ޱ�־ = False
                '2012.12.06   ����
                
                '�޸��ˣ����� 2012.12.07
                'bug�ţ�0000022
                '˵�������˳���ʾ���Ǳ������Ժ�����sub��ֱ�ӹرմ��ڡ�����
'                Exit Sub
                Unload Me
                '2012.12.07   ����
            Else
                Unload Me
            End If
        Else
            Unload Me
        End If
        
        
        Set frmResultInput_Routine = Nothing
    '2012-06-21 �ڵ�� ��
    End Select
    Exit Sub
    
errHandler:
    If Err.Number <> 0 Then
        Dim lstrError As String
        lstrError = func������(Err.Number, Err.Description)
        sfsub������ "ְҵ�����沿��", "FrmInputTestResult", "mobj����ͨ�ö���_BeforeOperate", 6666, lstrError, False
    End If
    If Operate = "����" Then
        '�ָ�������Բ�����
        ctbMain.Enabled = True
        ctabPerson.Enabled = True
        Frame1.Enabled = True
    End If
    MousePointer = 0
    'cstbMain.Panels(1) = ""
    Cancel = True
    Exit Sub
    Resume
End Sub

Private Sub MSComm1_OnComm()
    If MSComm1.InBufferCount Then
        ' ͨѶ���м��������ϵĻ�, ���ȡ����
           Dim InStringB() As Byte
           Dim instring As String
          InStringB = MSComm1.Input
          instring = StrConv(InStringB, vbUnicode)
          Text��ʱ.Text = Text��ʱ.Text & instring
          InStringB = ""
    End If
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    If ctabPerson.Tab = 0 Then
        cgrdSingleList.rows = 1
        cgrdInfoBatch.rows = 1
        TotalPeople.Caption = cgrdSingleList.rows - 1
    Else
        cgrdSingleList.rows = 1
        cgrdInfoBatch.rows = 1
        TotalPeopleBatch.Caption = cgrdInfoBatch.rows - 1
    End If
End Sub

'������д���۴���
Private Sub WriteConclun_Click()
    frmConclunInput_Write.Show 1
End Sub

Sub LoadPersonalInfoBatch(ByVal paraSysNo As String)
    On Error GoTo errHandler
    Dim lobj��� As Object     'ְҵ������
    Dim lobj��켯 As Object   '��켯�������ڸ����Թܱ��+���ڻ�ȡϵͳ��š�
    Dim lobjRec As Object
    
    Dim lstrNo As String       'ϵͳ��Ż��Թܱ�š�
    Dim llngNoType As Long     '������ͣ�0 ϵͳ���/1 �Թܱ�š�
    Dim lstrSysNo As String    'ϵͳ��š�
    Dim i As Long
    
    
    '��ȡ�����ϵͳ��ţ����Թܱ�ţ���
    lstrNo = paraSysNo
    
    If lstrNo <> "" Then
        '����ְҵ������
        Set lobj��� = CreateObject("ְҵ������.clsMedicalExam")
        
        lobj���.ϵͳ��� = lstrNo
        
        lstrSysNo = lobj���.ϵͳ���
        'ctxtSingleNo.Text = lstrSysNo
        
        '��ս��档
        If ctabPerson.Tab = 1 Then
            ctxt���� = ""
            ctxt�Ա� = ""
            ctxt���� = ""
            ctxt��λ���� = ""
            cpicPhoto(0).Picture = Nothing
        End If
        
        '�ж��Ƿ���ڡ�
        If Not lobj���.�Ƿ��Ѵ��� Then
            Err.Raise 6666, , "�������������ŵ������Ա�����������롣"
        End If
        
        '�ж��Ƿ����������ۡ�
        'If lobj���.���״̬ = P_ENDED_STATUS Then
        '    Err.Raise 6666, , "�������ŵ�����ѱ�ҽʦȷ���������ۣ��������޸������������۵��������" & Chr(13) & Chr(10) & "��ȷʵҪ�޸ģ����½��۵�ҽʦ���롰������¼�롱����������ȡ���½��ۣ��ٻص��˲��������޸ġ�"
        'End If
        
        '��ʾ��Ա��Ϣ��������Ƭ����
        With lobj���.�����Ա
            .����������� = lstrSysNo
            ctxt���� = .����
            ctxt�Ա� = .�Ա�
            ctxt���� = .����
            ctxt��λ���� = .��λ����
            ctxtΣ������ = .Σ������
            Picture4.Enabled = True
            Picture4.Visible = True
            Picture4.Picture = .��Ƭ
                
            '2012-04-11
            '��ʾ��Ա��Ϣ���ܽ����޸�
            ctxt�������.Enabled = False
            ctxt����.Enabled = False
            ctxt�Ա�.Enabled = False
            ctxt����.Enabled = False
            ctxt��λ����.Enabled = False
            '���水ť�ܹ����в���
            ccmdAutoFull(0).Enabled = True
            ccmdAutoFull(1).Enabled = True
            ccmdAutoFull(2).Enabled = True
            ccmdAutoFull(3).Enabled = True
            Frame5.Enabled = True
            '2012-04-11
                
            If llngNoType = 1 Then 'ϵͳ������뷽ʽ����Ҫ��ʾ�Թܱ�š�
            '    clblInfo(4) = lobj���.��쵥��
                    Label1(8).Caption = "��쵥�ţ�"
            Else
                'clblInfo(4) = lobj���.�Թܱ��
                'Label1(8).Caption = "�Թܱ�ţ�"
            End If
                
            '��ʾ��Ƭ��
            If Not .��Ƭ Is Nothing Then
                cpicPhoto(0).Picture = .��Ƭ
            End If
        End With
            
        '���������¼������
        If cchk���������.Value = 1 Then
            Exit Sub
        End If
        subShowInputGrid lstrSysNo
            
        cgrdSingleList.Row = 0
        
    End If

    ctbMain.Buttons(1).Enabled = True
    'cstbMain.Panels(1) = ""
    cgrdInput.Row = 1
    cgrdInput.Col = 2
    cgrdInput.SetFocus
    SendKeys ""
    Exit Sub
errHandler:
    If ctabPerson.Tab = 0 Then
        ctbMain.Buttons(1).Enabled = False
    End If
    sfsub������ "ְҵ�������¼��", InputFlag & "���¼��", "subShowSinglePerson", Err.Number, Err.Description, True
    Exit Sub
    Resume
End Sub

'���ܣ���ս�������
'���ߣ�����
'ʱ�䣺2012-05-08
Sub subClear()
    TotalPeople.Caption = 0
    TotalPeopleBatch.Caption = 0
    
    '��ǰ���治�ɲ���
    cgrdInput.rows = 1
            
    cdtpInputDate.Value = Now
    '��յ�ǰ������Ϣ
    ctxtSingleNo.Text = ""
    ctxtSingleNo.Enabled = True
    ctxtName.Text = ""
    ctxtSex.Text = ""
    ctxtAge.Text = ""
    ctxtCompanyName.Text = ""
'    cgrdSingleList.rows = 1   '��Ҫ����б�  2015-12-25 by Ĳ��
    
    '��ѯ������Ϣ
    ctxtcchkNo.Text = ""
    ctxtcchkWork.Text = ""
    ctxtCheckName.Text = ""
    
    '������Ϣ���
    DTP¼������.Value = Now
    ctxt�������.Text = ""
    ctxt�������.Enabled = True
    ctxt����.Text = ""
    ctxt�Ա�.Text = ""
    ctxt����.Text = ""
    ctxt��λ����.Text = ""
    cgrdInfoBatch.rows = 1
    '������Ϣ��־���
    cchk���������.Value = 0
    cchkDateBatch.Value = 0
    cchkCompanyBatch.Value = 0
    ctxtConclun.Text = ""
    TotalPeopleBatch.Caption = "0"
    cdtpDateBatch.Value = Date
    ctxtQueyCompanyBatch.Text = ""
    
    '�޸��ˣ������ 2012-12-28  ��
    '˵�������cgrdInfoBatch�����Ϣ
    'bug�ţ�0000125
    cgrdInfoBatch.Clear
    '�޸��ˣ������ 2012-12-28  ��

'    '��ղ�ѯ�������һ��Ҫ�е�,Ҳûдȫ��
'    cchkDate.Value = 0
'    cdtpDate.Value = Now
'    cgrdInfo.Clear

    '�����Ƭ
    Set cpicPhoto(0).Picture = Nothing
    Set Picture4.Picture = Nothing
    

    '�ָ�Ϊform_loadʱ��״̬��
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
        ctxt�������.Enabled = True
        ctxt����.Enabled = True
        ctxt�Ա�.Enabled = True
        ctxt����.Enabled = True
        ctxt��λ����.Enabled = True
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
    
    '2012-06-21 �ڵ�� ��
    '��ʼ����ǰ¼��״̬(��ǰ�ж�����Ȩ���޸ģ����ޣ�ֱ�Ӹ�ֵΪ3)
    ResultChanged = IIf(ResultChanged <> 3, -1, 3)
    cchkˢ����_Click
    '2012-06-21 �ڵ�� ��
    
    '2012-08-22 �ڵ�� ��
    '�����ʷ��¼��ؿؼ����ݺͿ������
    ccmbHistory.Clear
    ccmbHistory.Enabled = False
    cgrdInputHistory.rows = 1
    ctxtConclunHistory.Text = ""
    '2012-08-22 �ڵ�� ��
    
    '�޸��ˣ������ ʱ�䣺2012-12-10 ��
    '��ʼ����������
    '0000017
    '�޸��ˣ����� 2013.01.11   ����
    'bug�ţ�0000175
    '�޸�˵������������ʱ����ִ��һ�Ρ�
    If ctabPerson.Tab = 1 Then
        sub��������ģ��
    End If
    '�޸��ˣ����� 2013.01.11  ����
    '�޸��ˣ������ ʱ�䣺2012-12-10 ��
'    If ctabPerson.Tab = 1 Then
'        Dim i As Integer
'        Dim lobjRec As Object
'
'        '�������������Ͽ���
'        Set lobjRec = dafuncGetData("select ��������,�����Ա���� from ְҵ�����_����ģ�������Ϣ��")
'
''        Ccmb��������.Clear
'
'
'        For i = 0 To coptClasses.Count - 1
'            If coptClasses(i) Then
'                lobjRec.Filter = "�����Ա���� = '" & Trim(coptClasses(i).Caption) & "'"
'            End If
'        Next
'
'        While Not lobjRec.EOF
'            Ccmb��������.AddItem lobjRec("��������")
'            lobjRec.MoveNext
'        Wend
'        Exit Sub
'    End If

End Sub

'���ܣ������ύ�����Ա�������
'ʱ�䣺2012-04-26
'���ߣ�����
Private Sub sub��������()
    MousePointer = 11
    Dim lblnNotOver As Boolean
    Dim i As Integer
    Dim barCode As Collection '���������������
        'cstbMain.Panels(1) = "���ڱ��棬���Ժ�..."
        
        
        '��ʱ���治�ܲ�����
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
        '��ȡ���������Ա����������
        For i = 1 To cgrdInfoBatch.rows - 1
            barCode.Add cgrdInfoBatch.TextMatrix(i, 0)
        Next i
        
        '����
        'ʱ�䣺2012-05-23 ������
        For i = 1 To cgrdInput.rows - 1
            lcolItem.Add cgrdInput.TextMatrix(i, 0)
            lcolResult.Add cgrdInput.TextMatrix(i, 2)
        Next
        'ʱ�䣺2012-05-23 ������
        
        '��û��¼�꣬������ʾ��
        If lblnNotOver Then
            If Not sffuncMsg("��û��¼�����������Ŀ����������Ƿ���Ҫ���棿", sfѯ��) Then
                '�û�ѡ�񲻱��档
                GoTo errHandler
            End If
        End If

        subSaveBatch
        MousePointer = 0
        'cstbMain.Panels(1) = "����ɹ���"
        'Cancel = True
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�������¼��", "FrmENT_ResultInput", "ccmdSave_Click", 6666, lstrError, False
    
End Sub

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
    Dim lcolConclusion As String '�ι��ܿƵ�������
    Dim i As Integer
    Set barCode = New Collection
    For i = 1 To cgrdInfoBatch.rows - 1
        barCode.Add cgrdInfoBatch.TextMatrix(i, 0)
    Next i
    
    '��ʾ���������
'    ccrpI = barCode.Count
'    ccrp����.Max = ccrpI * 3
'    ccrp����.Visible = True
'    ccrp����.Caption = "0%"
'    ccrp����.Value = 0
    
    
    Set lobjTmp = CreateObject("ְҵ�������¼��.clsCommon")
    For i = 1 To barCode.Count
        isOk = lobjTmp.func���浥�������(barCode(i), um�û���, DTP¼������.Value, lcolItem, lcolResult, "ְҵ�����_�����Ϣ_" & InputFlag)
'        ccrp����.Caption = Int(i / ccrp����.Max * 100) & "%"
'        ccrp����.Value = ccrp����.Value + 1
'        If i = barCode.Count Then ccrpValue = Int(i / ccrp����.Max * 100)
    Next i
    
    '2012-07-15 �ڵ�� ��
    '��������ͼƬ�����
    Dim ifDraw As Boolean
    Dim lstrNo As String
    Dim savedPic As StdPicture
    Dim lcolְҵ������ As Object
    Dim lstrItemNo As String

    Set lcolְҵ������ = CreateObject("ְҵ�������¼��.clsCommon")
    lstrItemNo = lcolְҵ������.func��ȡ�����Ŀ���(mstr���ͼƬ����)
    lstrNo = ctxt�������.Text
    ifDraw = sub�Ƿ��л�ͼ��Ŀ(lstrNo)
    If ifDraw = True Then Set savedPic = lcolְҵ������.func��ȡ���ͼƬ(lstrNo, lstrItemNo, "")
    For i = 1 To barCode.Count
        If ifDraw = True Then
            isOk = sub�Ƿ��л�ͼ��Ŀ(barCode(i))
            If isOk = True Then Call lcolְҵ������.func������ͼƬ(savedPic, barCode(i), lstrItemNo, Now)
        End If
'        ccrp����.Caption = Int(i / ccrp����.Max * 100) + ccrpValue & "%"
'        ccrp����.Value = ccrp����.Value + 1
'        If i = barCode.Count Then ccrpValue = Int(i / ccrp����.Max * 100)
    Next i
    Set savedPic = Nothing
    Set lcolְҵ������ = Nothing
    '2012-07-15 �ڵ�� ��
    
    If ResultChanged <> 3 Then ResultChanged = 1
    If isOk = True Then
        For i = 1 To barCode.Count
            '���浥����Ŀ��ҽ������
            lcolConclusion = ctxtConclun.Text
            pobjҵ�����.sub������д������ barCode(i), priDeptName, lcolConclusion, um�û����
'            ccrp����.Caption = Int((i + 2 * barCode.Count) / ccrp����.Max * 100) & "%"
'            ccrp����.Value = ccrp����.Value + 1
            
            '2012-07-03 �ڵ�� ��
            '����һ���ֶ�"�޸���ʼʱ��"���޸ġ�ͬʱ�޸ĸÿ��ҵ������¼��״̬��
            pobjҵ�����.sub�޸���ʼʱ�� barCode(i), priDeptName
            pobjҵ�����.sub�޸Ľ��¼��״̬ barCode(i), priDeptNo, "2"
            pobjҵ�����.sub���¼���޸����״̬ barCode(i), "4"
            '2012-07-03 �ڵ�� ��
        Next i
        MsgBox ("��������ɹ���")
    Else
        subClear
    End If
    
    subClear
    
    ccrp����.Visible = False
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�������¼��", "frmENT_ResultInput", "subSave", 6666, lstrError, False
End Sub

'2012-06-21 �ڵ��
Sub sub��ȡϵͳ��Ź̶�����()
    '��ȡ����������
    Dim lobjRec As Object
    Set lobjRec = dafuncGetData("select getdate()")
    ctxtSingleNo.Text = um����վ��� & um���������� & Format(lobjRec(0), "yyyy")
    Set lobjRec = Nothing
End Sub

'2012-07-14 �ڵ��
Sub sub���¿��޸Ľ����Ա�޸�״̬()
    Dim lobjRec As Object
    Dim strSQL As String
    Dim canModify As Boolean
    dasubSetQueryTimeout 6000
    strSQL = "select ϵͳ���,�������״̬ from ְҵ�����_���������ݿ� where substring(�������״̬," & priDeptNo & ",1)='1'  or substring(�������״̬," & priDeptNo & ",1)='2' or substring(�������״̬," & priDeptNo & ",1)='3'"
'    strSQL = "select ϵͳ���,�������״̬ from ְҵ�����_���������ݿ� where substring(�������״̬," & priDeptNo & ",1)='1' or substring(�������״̬," & priDeptNo & ",1)='2'"
    Set lobjRec = dafuncGetData(strSQL)
    If lobjRec.RecordCount = 0 Then Exit Sub
    lobjRec.MoveFirst
    While lobjRec.EOF <> True
        canModify = pobjҵ�����.sub�Ƿ����޸�ʱ�䷶Χ��(lobjRec("ϵͳ���"), priDeptName, 8)
'        canModify = pobjҵ�����.sub�Ƿ����޸�ʱ�䷶Χ��(lobjRec("ϵͳ���"), priDeptName, 120)   '���ԣ���ʱ�䷶Χ��Ϊ5�켴120Сʱ  Ĳ��  2015-11-6
        If canModify = False Then Call pobjҵ�����.sub�޸Ľ��¼��״̬(lobjRec("ϵͳ���"), priDeptNo, "3")
        lobjRec.MoveNext
    Wend
End Sub

'2012-07-14 �ڵ��
Sub sub��ѯ�б���ʾ(ByVal coptIndex As Integer)
    If mobjQueryResult Is Nothing Then Exit Sub
    mobjQueryResult.Filter = ""
    
    If mobjQueryResult.RecordCount > 0 Then
    
        If ctabPerson.Tab = 0 Then
            If cchkSigResult(0).Value = 1 And cchkSigResult(1).Value = 0 Then
                mobjQueryResult.Filter = "��дʱ��<>null"
            ElseIf cchkSigResult(0).Value = 0 And cchkSigResult(1).Value = 1 Then
                mobjQueryResult.Filter = "��дʱ��=null"
            ElseIf cchkSigResult(0).Value = 0 And cchkSigResult(1).Value = 0 Then
                mobjQueryResult.Filter = "ϵͳ���='xxx'"
            Else
'                mobjQueryResult.Filter = ""
                '��������δ����ͬʱѡ�У�ҲҪ��ʾ��˲���YK�����洦�� 2015-12-23 by Ĳ��
                 mobjQueryResult.Filter = "�������<>null"
            End If
        ElseIf ctabPerson.Tab = 1 Then
            If cchkBchResult(0).Value = 1 And cchkBchResult(1).Value = 0 Then
                mobjQueryResult.Filter = "��дʱ��<>null"
            ElseIf cchkBchResult(0).Value = 0 And cchkBchResult(1).Value = 1 Then
                mobjQueryResult.Filter = "��дʱ��=null"
            ElseIf cchkBchResult(0).Value = 0 And cchkBchResult(1).Value = 0 Then
                mobjQueryResult.Filter = "ϵͳ���='xxx'"
            
            Else
'                mobjQueryResult.Filter = ""
                '������ͬ���Ĵ���
                mobjQueryResult.Filter = "�������<>null"
            End If
        End If
        
        If mobjQueryResult.Filter <> "" And mobjQueryResult.Filter <> 0 And mobjQueryResult.Filter <> "ϵͳ���='xxx'" Then
        '���˿������˲�����¼���������������Ϊ��˲���XX���� 2015-11-23 by Ĳ��
            mobjQueryResult.Filter = mobjQueryResult.Filter & " and ������� like '" & coptClasses(coptIndex).Caption & "%'"
'            mobjQueryResult.Filter = mobjQueryResult.Filter & " and ������� like '" & coptClasses(coptIndex).Caption & "'"
        Else
'            mobjQueryResult.Filter = "�������='" & coptClasses(coptIndex).Caption & "'"
            '��������δ����ͬʱѡ�У�ҲҪ��ʾ��˲���YK    2015-12-23 by Ĳ��
            mobjQueryResult.Filter = mobjQueryResult.Filter & " and ������� like '" & coptClasses(coptIndex).Caption & "%'"
            
        End If
        
         '2012-10-26 �����
        If ctabPerson.Tab = 1 Then
'            mobjQueryResult.Filter = " and ������='" & Trim(Ccmb��������.Text) & "'"
            '�޸��ˣ����� 2012.12.06
            'bug�ţ�0000014
            '˵�������������ˡ�����
'            mobjQueryResult.Filter = mobjQueryResult.Filter & " and ������='" & Trim(Ccmb��������.Text) & "'"
            mobjQueryResult.Filter = mobjQueryResult.Filter & " and ������='" & Trim(Ccmb��������.Text) & "'"
            '2012.12.06  ����
        End If
        
    End If 'mobjQueryResult.recordcount = 0
    
    If ctabPerson.Tab = 0 Then
        With cgrdSingleList
            Set .DataSource = mobjQueryResult
            .Col = 0
            .Sort = flexSortGenericDescending
            .AutoSize 0, .cols - 1, 0, 0
            .ExplorerBar = flexExSort
            .DataMode = flexDMFree
            .AllowSelection = True
            .AllowBigSelection = False
            .SelectionMode = flexSelectionByRow
        End With
'        TotalPeople.Caption = IIf(mobjQueryResult.RecordCount = 0, "0", mobjQueryResult.RecordCount)
        TotalPeople.Caption = cgrdSingleList.rows - 1
    Else
        With cgrdInfoBatch
            Set .DataSource = mobjQueryResult
            .Col = 0
            .Sort = flexSortGenericDescending
            .AutoSize 0, .cols - 1, 0, 0
            .ExplorerBar = flexExSort
            .DataMode = flexDMFree
            .AllowSelection = True
            .AllowBigSelection = True
            .SelectionMode = flexSelectionListBox
        End With
'       TotalPeopleBatch.Caption = IIf(mobjQueryResult.RecordCount = 0, "0", mobjQueryResult.RecordCount)
        TotalPeopleBatch.Caption = cgrdInfoBatch.rows - 1
    End If
    cgrdInput.rows = 1
    TotalPeople.Caption = IIf(mobjQueryResult.RecordCount = 0, "0", mobjQueryResult.RecordCount)

End Sub

'2012-07-15 �ڵ��
'�ж�ĳ�������Ա�Ƿ��л�ͼ��Ŀ
Private Function sub�Ƿ��л�ͼ��Ŀ(ByVal paraSysNo As String) As Boolean
    Dim lcolְҵ������ As Object
    Dim lobjRec As Object
    Dim strSQL, lstrItemNo As String

    Set lcolְҵ������ = CreateObject("ְҵ�������¼��.clsCommon")
    lstrItemNo = lcolְҵ������.func��ȡ�����Ŀ���(mstr���ͼƬ����)
    strSQL = "select * from ְҵ�����_�����Ϣ_" & InputFlag & " where ϵͳ���='" & paraSysNo & "' and �����Ŀ='" & lstrItemNo & "'"
    Set lobjRec = dafuncGetData(strSQL)
    sub�Ƿ��л�ͼ��Ŀ = lobjRec.RecordCount > 0
    Set lcolְҵ������ = Nothing
    Set lobjRec = Nothing
End Function
'2012-10-26 �����
Sub sub��������ģ��()
    Dim i As Integer
    Dim lobjRec As Object
    On Error GoTo errHandler

    '�������������Ͽ���
    Set lobjRec = dafuncGetData("select ��������,�����Ա���� from ְҵ�����_����ģ�������Ϣ��")
    
    Ccmb��������.Clear
    
 
    For i = 0 To coptClasses.Count - 1
        If coptClasses(i) Then
            lobjRec.Filter = "�����Ա���� = '" & Trim(coptClasses(i).Caption) & "'"
        End If
    Next
    
    While Not lobjRec.EOF
        Ccmb��������.AddItem lobjRec("��������")
        lobjRec.MoveNext
    Wend
    If Ccmb��������.ListCount >= 1 Then
    Ccmb��������.ListIndex = 0
'    Ccmb��������.Refresh
'    Ccmb��������.Enabled = True
    End If
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ������", "frmFinalConclusion", "sub��������ģ��", 6666, lstrError, False
    Exit Sub
    Resume
End Sub
