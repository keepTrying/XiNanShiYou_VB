VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3A75EE84-8E64-43F4-A904-E4835B9A3DB4}#3.9#0"; "DICOMax.ocx"
Object = "{FC07EBD4-FE92-11D0-A199-A0077383D901}#5.1#0"; "CCRPPRG.OCX"
Begin VB.Form frmBUS_ResultInput 
   Caption         =   "B超影像科结果录入窗口"
   ClientHeight    =   9735
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14595
   LinkTopic       =   "Form1"
   ScaleHeight     =   9735
   ScaleWidth      =   14595
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox Picture1 
      Height          =   9615
      Left            =   0
      ScaleHeight     =   9555
      ScaleWidth      =   14475
      TabIndex        =   0
      Top             =   0
      Width           =   14535
      Begin VB.OptionButton coptClasses 
         BackColor       =   &H00C0FFFF&
         Caption         =   "8023部队"
         Height          =   255
         Index           =   4
         Left            =   4440
         TabIndex        =   74
         Top             =   840
         Width           =   1095
      End
      Begin VB.OptionButton coptClasses 
         BackColor       =   &H00C0FFFF&
         Caption         =   "涉核部队"
         Height          =   255
         Index           =   3
         Left            =   3360
         TabIndex        =   73
         Top             =   840
         Width           =   1095
      End
      Begin VB.OptionButton coptClasses 
         BackColor       =   &H00C0FFFF&
         Caption         =   "普通体检"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   72
         Top             =   840
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton coptClasses 
         BackColor       =   &H00C0FFFF&
         Caption         =   "职业健康"
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   71
         Top             =   840
         Width           =   1095
      End
      Begin VB.OptionButton coptClasses 
         BackColor       =   &H00C0FFFF&
         Caption         =   "放射健康"
         Height          =   255
         Index           =   2
         Left            =   2280
         TabIndex        =   70
         Top             =   840
         Width           =   1095
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   9495
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   14415
         Begin VB.CheckBox cchk刷条码 
            Caption         =   "刷条码"
            Height          =   255
            Left            =   11760
            TabIndex        =   69
            Top             =   120
            Width           =   1215
         End
         Begin VB.Frame fraPicShow 
            Caption         =   "图片显示栏"
            Height          =   6495
            Left            =   6720
            TabIndex        =   8
            Top             =   2880
            Width           =   7575
            Begin DICOMax.DICOMX Dicm 
               Height          =   5895
               Left            =   120
               TabIndex        =   9
               Top             =   480
               Width           =   5415
               Object.Visible         =   -1  'True
               BorderStyle     =   0
               Enabled         =   -1  'True
               ImageSlicesCurrent=   0
               ImageZoomPct    =   100
               ImageSmoothOn   =   -1  'True
               ImageWinCenter  =   0
               ImageWinWidth   =   0
               ImageOverlayOn  =   0   'False
               ImageOverlayColor=   1
               ImageColorScheme=   1
               ImageTool       =   1
               ToolBarVisible  =   0   'False
               ImageZoomBestFit=   -1  'True
               ImageUseRefContrast=   -1  'True
               ImageShowHeaderInfo=   0   'False
               OpenFileName    =   ""
               ToolBarPos      =   1
               DICOMInstitutionName=   "Anonymous"
               DICOMInstitutionAddress=   ""
               DICOMStudyDescrp=   "Study Description"
               DICOMManufacturer=   "http://www.dicom3.cn/"
               DICOMSeriesTag  =   ""
               DICOMPatientName=   "NONAME"
               DICOMPatientID  =   "NOID"
               DICOMModality   =   "OT"
               DICOMSOPClassUID=   ""
               DICOMSOPInstanceUID=   ""
               ImagePOS        =   5
               ImageVScrollPosition=   17
               ImageHScrollPosition=   10
               DICOMPatientBirthDate=   ""
               DICOMPatientSex =   ""
               ImageOverlayLanguage=   1
               ImageMagnifyZoomSize=   2
               DICOMStudyDate  =   ""
               DICOMSeriesDate =   ""
               DICOMImageDate  =   ""
               DICOMStudyTime  =   ""
               DICOMSeriesTime =   ""
               DICOMImageTime  =   ""
               DICOMSeriesNumber=   0
               DICOMImageNumber=   0
               DICOMRefPhyName =   ""
               DICOMStudyInstanceUID=   ""
               DICOMSeriesInstanceUID=   ""
               DICOMStudyID    =   ""
               ImageMeasureMaxItem=   8
               ImageMeasureResultIndex=   1
               ImageXGRGBColor =   0
               DICOMDirStudyPos=   1
               DICOMDirSeriesPos=   1
               DICOMDirImagePos=   1
               DICOMImplementationClassUID=   ""
               DICOMImplementationVersionName=   ""
               DICOMSourceApplicationEntityTitle=   ""
               FrameOfReferenceUID=   ""
               ImageMeasureFontSize=   8
               ImageMeasureTextPreSet=   ""
               ImageMeasureTextFontSize=   8
               ImageMeasureSelectIndex=   0
               OCXLanguage     =   1
               ImageMagnifySize=   60
               ImageOverlayFontSize=   13
               ImageOverlayFontName=   "Lucida Console"
               ImageOverlayShowRuler=   0   'False
               EnableMouseScroll=   -1  'True
               DICOMPixelSpaceWidth=   0
               DICOMPixelSpaceHeight=   0
               ShowRealTimeImage=   -1  'True
               ImageSortByFileName=   0   'False
               ImageXGPaletteColor=   2
               ImageMagnifyProcess=   0
               DICOMSeriesDescrp=   "Series Description"
               DICOMProtocolName=   "ProtocolName"
               DICOMMModelName =   ""
               ImagePreviewDataAddress=   0
               LicenseCode     =   ""
               ImageOverlayManualControl=   0   'False
               ImageToolAfterMeasure=   0
               ImageOverlayShowPixelValue=   0   'False
               EnableMouseRightBtnWL=   -1  'True
               EnableMouseDBClick=   -1  'True
               ImageResetPOSOnSizeChange=   -1  'True
               ImageStretchMeasureOnExport=   0   'False
               ImageForceStretchMeasurement=   0   'False
               ImageMaskColor  =   0
               ImageReScaleOnResize=   0   'False
               BorderSize      =   0
               BorderVisible   =   0   'False
               BorderColor     =   0
               ImageMeasureEnableDeleteKey=   -1  'True
               ImageProcessRotate=   0
               DICOMWriteWCElement=   0   'False
               DICOMPhotometricInterpretation=   0
               DICOMPlanarConfiguration=   0
               ImagePositionLinesColor=   65280
               ImagePositionLinesWidth=   1
               ImagePositionLinesSeledtedColor=   255
               AcceptDragItems =   0   'False
               Object.Index           =   0
               AcceptedDragItems=   0   'False
               DICOMFrameOfReferenceUID=   ""
               ImageOverlayShowPosLines=   -1  'True
               ImagePositionLinesFontSize=   1
               ImagePositionLinesDrawStyle=   1
               ImageMeasureLineColor=   255
               ImageMeasureLineColorSelected=   255
               ImageMeasureFontColor=   255
               ImageMeasureTextFontColor=   255
               ImageSaveCompressType=   0
               ImageOverlayRulerColor=   65280
               ImageOverlayTextSimpleDraw=   0   'False
               DICOMImageSliceLocation=   0
               DICOMAcquisitionNumber=   0
               DICOMAccessionNumber=   0
               DICOMImageType  =   ""
               DICOMOperatorsName=   ""
               DICOMBodyPartExamined=   ""
               DICOMInstitutionDepName=   ""
               DICOMPatientPosition=   ""
               DICOMPatientAge =   ""
               ImageMeasureBorderSelectSize=   10
               ImageRulerFontSize=   10
               ImageWLSpeedRatio=   5
               ImageZoomSpeedRatio=   5
               DICOMViewPosition=   ""
               DICOMPatientOrientation=   ""
               DICOMImageBitAllocated=   0
               DICOMImageBitStored=   0
               DICOMImageHighBit=   0
               DICOMWinCenter  =   127
               DICOMWinWidth   =   255
               ImageAnnotationFontSize=   0
               DICOMDirPatientPos=   1
            End
            Begin VB.CommandButton ccmdDCMOpen 
               Caption         =   "打开图片文件"
               Height          =   495
               Left            =   5640
               TabIndex        =   13
               Top             =   600
               Width           =   1815
            End
            Begin VB.CommandButton ccmdSavePic 
               Caption         =   "保存当前图片"
               Height          =   495
               Left            =   5640
               TabIndex        =   12
               Top             =   5640
               Width           =   1815
            End
            Begin VB.FileListBox DCMList 
               Height          =   4230
               Left            =   5640
               TabIndex        =   11
               Top             =   1080
               Width           =   1815
            End
            Begin VB.CheckBox cchkReplace 
               Caption         =   "保存时覆盖"
               Height          =   255
               Left            =   5760
               TabIndex        =   10
               Top             =   6120
               Width           =   1665
            End
            Begin VB.Timer Timer1 
               Interval        =   700
               Left            =   6000
               Top             =   120
            End
            Begin MSComDlg.CommonDialog Diag 
               Left            =   6720
               Top             =   120
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin VB.Label llabCurr 
               Caption         =   "第？张/共？张"
               Height          =   255
               Left            =   5880
               TabIndex        =   15
               Top             =   5400
               Width           =   1455
            End
            Begin VB.Label Label12 
               BackColor       =   &H00FFC0C0&
               Caption         =   "图片上按住鼠标左键或右键拖动可以改变图像对比度"
               Height          =   255
               Left            =   120
               TabIndex        =   14
               Top             =   240
               Width           =   4215
            End
         End
         Begin VB.Frame fraResult 
            Caption         =   "结论录入"
            Height          =   855
            Left            =   6720
            TabIndex        =   5
            Top             =   1800
            Width           =   7575
            Begin VB.CommandButton Cmd结论模板 
               Caption         =   "结论模板"
               Height          =   495
               Left            =   6360
               TabIndex        =   65
               Top             =   240
               Width           =   975
            End
            Begin VB.TextBox ctxtResult 
               Height          =   615
               Left            =   960
               MultiLine       =   -1  'True
               TabIndex        =   6
               Top             =   120
               Width           =   5175
            End
            Begin VB.Label Label8 
               BackColor       =   &H00C0FFC0&
               Caption         =   "医师结论"
               Height          =   255
               Left            =   120
               TabIndex        =   7
               Top             =   240
               Width           =   855
            End
         End
         Begin VB.Frame fraPicTool 
            Caption         =   "描述录入"
            Height          =   855
            Left            =   6720
            TabIndex        =   2
            Top             =   840
            Width           =   7575
            Begin VB.TextBox ctxtPResult 
               Height          =   615
               Left            =   960
               MultiLine       =   -1  'True
               TabIndex        =   3
               Top             =   120
               Width           =   6135
            End
            Begin VB.Label Label10 
               BackColor       =   &H00C0FFC0&
               Caption         =   "图片描述"
               Height          =   255
               Left            =   120
               TabIndex        =   4
               Top             =   240
               Width           =   855
            End
         End
         Begin TabDlg.SSTab SSTPersonalInfo 
            Height          =   8175
            Left            =   0
            TabIndex        =   16
            Top             =   1200
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   14420
            _Version        =   393216
            Tabs            =   2
            Tab             =   1
            TabHeight       =   520
            ForeColor       =   8388608
            TabCaption(0)   =   "单个录入"
            TabPicture(0)   =   "frmBUS_ResultInput.frx":0000
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "fraQuery"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "fraInfo"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).ControlCount=   2
            TabCaption(1)   =   "批量录入"
            TabPicture(1)   =   "frmBUS_ResultInput.frx":001C
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
               Left            =   3120
               TabIndex        =   79
               Top             =   4560
               Value           =   1  'Checked
               Width           =   1095
            End
            Begin VB.CheckBox cchkBchResult 
               Caption         =   "已填结果"
               Height          =   255
               Index           =   0
               Left            =   1800
               TabIndex        =   78
               Top             =   4560
               Width           =   1095
            End
            Begin VB.CommandButton ccmdRemove 
               Caption         =   "移 除"
               Height          =   375
               Left            =   4800
               TabIndex        =   64
               Top             =   4440
               Width           =   855
            End
            Begin VB.CommandButton ccmdClear 
               Caption         =   "清 空"
               Height          =   375
               Left            =   4800
               TabIndex        =   63
               Top             =   4920
               Width           =   855
            End
            Begin VB.CheckBox cchkDateBatch 
               BackColor       =   &H00C0FFC0&
               Caption         =   "体检日期"
               Height          =   255
               Left            =   480
               TabIndex        =   62
               Top             =   3480
               Width           =   1215
            End
            Begin VB.Frame fraQueryBatch 
               Caption         =   "批量查询体检人员"
               Height          =   2895
               Left            =   240
               TabIndex        =   46
               Top             =   480
               Width           =   5775
               Begin VB.CommandButton ccmdLocateBatch 
                  Caption         =   "单位定位"
                  Height          =   375
                  Left            =   4200
                  TabIndex        =   54
                  Top             =   2400
                  Width           =   975
               End
               Begin VB.PictureBox Picture4 
                  Height          =   1935
                  Left            =   3960
                  ScaleHeight     =   1875
                  ScaleWidth      =   1515
                  TabIndex        =   53
                  Top             =   240
                  Width           =   1575
               End
               Begin VB.TextBox ctxt单位名称 
                  Height          =   300
                  Left            =   1320
                  TabIndex        =   52
                  Top             =   2160
                  Width           =   2415
               End
               Begin VB.TextBox ctxt年龄 
                  Height          =   300
                  Left            =   1320
                  TabIndex        =   51
                  Top             =   1800
                  Width           =   2415
               End
               Begin VB.TextBox ctxt性别 
                  Height          =   300
                  Left            =   1320
                  TabIndex        =   50
                  Top             =   1440
                  Width           =   2415
               End
               Begin VB.TextBox ctxt姓名 
                  Height          =   300
                  Left            =   1320
                  TabIndex        =   49
                  Top             =   1080
                  Width           =   2415
               End
               Begin VB.TextBox ctxt体检条码 
                  Height          =   300
                  Left            =   1320
                  TabIndex        =   48
                  Top             =   720
                  Width           =   2415
               End
               Begin VB.CheckBox cchk套用体检结果 
                  BackColor       =   &H008080FF&
                  Caption         =   "该体检人员结果作为批量体检结果录入"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   47
                  Top             =   2520
                  Value           =   1  'Checked
                  Width           =   3615
               End
               Begin MSComCtl2.DTPicker DTP录入日期 
                  BeginProperty DataFormat 
                     Type            =   1
                     Format          =   "yyyy-MM-dd"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2052
                     SubFormatType   =   3
                  EndProperty
                  Height          =   300
                  Left            =   1320
                  TabIndex        =   55
                  Top             =   360
                  Width           =   2415
                  _ExtentX        =   4260
                  _ExtentY        =   529
                  _Version        =   393216
                  Enabled         =   0   'False
                  Format          =   72548352
                  CurrentDate     =   40969
               End
               Begin VB.Label Label11 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "结论录入日期"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   61
                  Top             =   360
                  Width           =   1095
               End
               Begin VB.Label Label14 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "单位名称"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   60
                  Top             =   2160
                  Width           =   975
               End
               Begin VB.Label Label15 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "年龄"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   59
                  Top             =   1800
                  Width           =   975
               End
               Begin VB.Label Label16 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "性别"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   58
                  Top             =   1440
                  Width           =   975
               End
               Begin VB.Label Label17 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "姓名"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   57
                  Top             =   1080
                  Width           =   975
               End
               Begin VB.Label Label18 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "体检条码号"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   56
                  Top             =   720
                  Width           =   975
               End
            End
            Begin VB.CommandButton ccmd查询单位 
               Caption         =   "单位定位"
               Height          =   375
               Left            =   4800
               TabIndex        =   45
               Top             =   3960
               Width           =   855
            End
            Begin VB.TextBox ctxtQueyCompanyBatch 
               Height          =   300
               Left            =   1800
               TabIndex        =   44
               Top             =   3960
               Width           =   2415
            End
            Begin VB.CheckBox cchkCompanyBatch 
               BackColor       =   &H00C0FFC0&
               Caption         =   "单位名称"
               Height          =   255
               Left            =   480
               TabIndex        =   43
               Top             =   3960
               Width           =   1215
            End
            Begin VB.CommandButton ccmdSelInfo 
               Caption         =   "查 询"
               Height          =   375
               Left            =   4800
               TabIndex        =   42
               Top             =   3480
               Width           =   855
            End
            Begin VB.Timer Timerccrp 
               Left            =   5520
               Top             =   4800
            End
            Begin VB.Frame fraQuery 
               Caption         =   "查询体检人员"
               Height          =   4815
               Left            =   -74880
               TabIndex        =   32
               Top             =   3240
               Width           =   6255
               Begin VB.CommandButton ccmdWork 
                  Caption         =   "单位定位"
                  Height          =   375
                  Left            =   3240
                  TabIndex        =   87
                  Top             =   960
                  Width           =   1185
               End
               Begin VB.CheckBox cchkSingleNo 
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "体检条码"
                  Height          =   255
                  Left            =   3120
                  TabIndex        =   86
                  Top             =   240
                  Width           =   1095
               End
               Begin VB.TextBox ctxtcchkNo 
                  Height          =   270
                  Left            =   4320
                  TabIndex        =   85
                  Top             =   240
                  Width           =   1695
               End
               Begin VB.CheckBox cchkCardNo 
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "身份证号"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   84
                  Top             =   600
                  Width           =   1095
               End
               Begin VB.CheckBox cchkWorkUnit 
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "单位名称"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   83
                  Top             =   960
                  Width           =   1095
               End
               Begin VB.TextBox ctxtcchkCardNo 
                  Height          =   270
                  Left            =   1320
                  TabIndex        =   82
                  Top             =   600
                  Width           =   1695
               End
               Begin VB.TextBox ctxtcchkWork 
                  Height          =   270
                  Left            =   1320
                  TabIndex        =   81
                  Top             =   960
                  Width           =   1695
               End
               Begin VB.CheckBox cchkSigResult 
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "未填结果"
                  Height          =   255
                  Index           =   1
                  Left            =   1440
                  TabIndex        =   77
                  Top             =   1320
                  Value           =   1  'Checked
                  Width           =   1095
               End
               Begin VB.CommandButton ccmdQuery 
                  Caption         =   "查   询"
                  Height          =   375
                  Left            =   4800
                  TabIndex        =   37
                  Top             =   960
                  Width           =   1215
               End
               Begin VB.CheckBox cchkSigResult 
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "已填结果"
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   36
                  Top             =   1320
                  Width           =   1095
               End
               Begin VB.TextBox ctxtCheckName 
                  Height          =   240
                  Left            =   4320
                  TabIndex        =   35
                  Top             =   600
                  Width           =   1695
               End
               Begin VB.CheckBox cchkName 
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "姓名"
                  Height          =   255
                  Left            =   3120
                  TabIndex        =   34
                  Top             =   600
                  Width           =   735
               End
               Begin VB.CheckBox cchkDate 
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "体检日期"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   33
                  Top             =   240
                  Width           =   1095
               End
               Begin MSComCtl2.DTPicker cdtpDate 
                  Height          =   255
                  Left            =   1200
                  TabIndex        =   38
                  Top             =   240
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   450
                  _Version        =   393216
                  Format          =   72548352
                  CurrentDate     =   40969
               End
               Begin VSFlex8Ctl.VSFlexGrid cgrdInfo 
                  Height          =   3015
                  Left            =   120
                  TabIndex        =   39
                  ToolTipText     =   "双击自动填入个人信息和已有体检结果"
                  Top             =   1680
                  Width           =   6015
                  _cx             =   2088774002
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
                  Left            =   6000
                  TabIndex        =   76
                  Top             =   1440
                  Width           =   90
               End
               Begin VB.Label Label9 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "人数："
                  Height          =   180
                  Left            =   5280
                  TabIndex        =   75
                  Top             =   1440
                  Width           =   540
               End
            End
            Begin VB.Frame fraInfo 
               Caption         =   "个人信息"
               Height          =   2775
               Left            =   -74880
               TabIndex        =   17
               Top             =   360
               Width           =   6255
               Begin VB.ComboBox ccmbHistory 
                  Height          =   300
                  Left            =   1440
                  Style           =   2  'Dropdown List
                  TabIndex        =   88
                  Top             =   600
                  Width           =   2415
               End
               Begin VB.CommandButton ccmdLocate 
                  Caption         =   "单位定位"
                  Height          =   255
                  Left            =   4920
                  TabIndex        =   24
                  Top             =   2400
                  Visible         =   0   'False
                  Width           =   975
               End
               Begin VB.TextBox ctxtCompanyName 
                  Height          =   270
                  Left            =   1440
                  TabIndex        =   23
                  Top             =   2400
                  Width           =   3375
               End
               Begin VB.TextBox ctxtAge 
                  Height          =   270
                  Left            =   1440
                  TabIndex        =   22
                  Top             =   2040
                  Width           =   2415
               End
               Begin VB.TextBox ctxtSex 
                  Height          =   270
                  Left            =   1440
                  TabIndex        =   21
                  Top             =   1680
                  Width           =   2415
               End
               Begin VB.TextBox ctxtName 
                  Height          =   270
                  Left            =   1440
                  TabIndex        =   20
                  Top             =   1320
                  Width           =   2415
               End
               Begin VB.TextBox ctxtBarCode 
                  Height          =   270
                  Left            =   1440
                  TabIndex        =   19
                  Top             =   960
                  Width           =   2415
               End
               Begin VB.PictureBox Picture2 
                  Height          =   1935
                  Left            =   4560
                  ScaleHeight     =   1875
                  ScaleWidth      =   1515
                  TabIndex        =   18
                  Top             =   240
                  Width           =   1575
               End
               Begin MSComCtl2.DTPicker cdtpConclusionDate 
                  Height          =   255
                  Left            =   1440
                  TabIndex        =   25
                  Top             =   240
                  Width           =   2415
                  _ExtentX        =   4260
                  _ExtentY        =   450
                  _Version        =   393216
                  Enabled         =   0   'False
                  Format          =   72548352
                  CurrentDate     =   40969
               End
               Begin VB.Label Label13 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "历年病历"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   89
                  Top             =   600
                  Width           =   975
               End
               Begin VB.Label Label7 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "结论录入日期"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   31
                  Top             =   240
                  Width           =   1095
               End
               Begin VB.Label Label5 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "单位名称"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   30
                  Top             =   2400
                  Width           =   975
               End
               Begin VB.Label Label4 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "年龄"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   29
                  Top             =   2040
                  Width           =   975
               End
               Begin VB.Label Label3 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "性别"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   28
                  Top             =   1680
                  Width           =   975
               End
               Begin VB.Label Label2 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "姓名"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   27
                  Top             =   1320
                  Width           =   975
               End
               Begin VB.Label Label1 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "体检条码号"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   26
                  Top             =   960
                  Width           =   975
               End
            End
            Begin VSFlex8Ctl.VSFlexGrid cgrdInfoBatch 
               Height          =   2415
               Left            =   240
               TabIndex        =   41
               Top             =   5520
               Width           =   5775
               _cx             =   2088773578
               _cy             =   2088767652
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
            Begin MSComCtl2.DTPicker cdtpDateBatch 
               Height          =   300
               Left            =   1800
               TabIndex        =   80
               Top             =   3480
               Width           =   2415
               _ExtentX        =   4260
               _ExtentY        =   529
               _Version        =   393216
               Format          =   72548352
               CurrentDate     =   40969
            End
            Begin CCRProgressBar.ccrpProgressBar ccrp进度 
               Height          =   375
               Left            =   600
               Top             =   5040
               Visible         =   0   'False
               Width           =   3615
               _ExtentX        =   6376
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
               Left            =   480
               TabIndex        =   68
               Top             =   4560
               Width           =   540
            End
            Begin VB.Label TotalPeopleBatch 
               AutoSize        =   -1  'True
               Caption         =   "0"
               Height          =   180
               Left            =   1080
               TabIndex        =   67
               Top             =   4560
               Width           =   90
            End
         End
         Begin MSComDlg.CommonDialog ccmdFile 
            Left            =   5280
            Top             =   360
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            Flags           =   6148
         End
         Begin MSComctlLib.Toolbar ctlb工具栏 
            Height          =   540
            Left            =   240
            TabIndex        =   40
            Top             =   0
            Width           =   14100
            _ExtentX        =   24871
            _ExtentY        =   953
            ButtonWidth     =   1455
            ButtonHeight    =   953
            Appearance      =   1
            Style           =   1
            ImageList       =   "cimg按钮图标"
            _Version        =   393216
         End
         Begin VB.Label LabelDoctor 
            BackColor       =   &H00C0FFFF&
            Caption         =   "医生："
            Height          =   255
            Left            =   5520
            TabIndex        =   66
            Top             =   840
            Width           =   1095
         End
      End
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
End
Attribute VB_Name = "frmBUS_ResultInput"
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
Private mobj体检医师  As Object   'clsMedicalExamer    获取当前体检医师可以作的指定属性（常规/化验）的体检项目
Private mlobjRec As Object

'查询结果
Private mstrDoctorName As String
Private mobjQueryResult As Object
Private mcolIndex As New Collection
Private indX, indY As Integer
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

Private mstrState As String     '记录当前体检状态

'2012-04-14 于登淼 ↓
'dicom相关变量
Private DCMPath As String       'dicom完整路径
Private DCMDir As String        'dicom文件夹路径
Private DCMFileName As String   'dicom当前文件名
Private DCMIdx As Integer       'dicom当前文件位置(在DCMList中的)
'2012-04-14 于登淼 ↑

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

'2012-06-21 于登淼
'添加刷条码判断
Private Sub cchk刷条码_Click()
    If Not cchk刷条码.Visible Then Exit Sub
    If ctxtBarCode.Enabled = False Then Exit Sub
    
    If SSTPersonalInfo.Tab = 0 Then
        ctxtBarCode.Text = ""
        If cchk刷条码.Value = 0 Then sub获取系统编号固定部分
        ctxtBarCode.Enabled = True
        ctxtBarCode.SetFocus
        ctxtBarCode.SelStart = Len(ctxtBarCode)
        ctxtBarCode.SelLength = 0
    Else
        ctxt体检条码.Text = ""
        ctxt体检条码.SetFocus
    End If
End Sub

'显示选中日期的病历信息
'翁乔
'2012-07-31
Private Sub ccmbHistory_Click()
    Dim lobjRec As Object

    If ccmbHistory.Text <> "——" Then
        ctlb工具栏.Buttons(2).Enabled = False
        Set lobjRec = mobj体检医师.func获取指定年份的体检描述(Trim(ctxtBarCode.Text), ccmbHistory.Text, "B超影像科")
        
        If Not lobjRec Is Nothing Then
            
            fraPicTool.Caption = "历年体检"
            ctxtPResult.Text = lobjRec("体检结果")
            fraPicTool.Enabled = False
            
            Set lobjRec = mobj体检医师.func获取指定年份的体检病历结论(Trim(ctxtBarCode.Text), "11", Trim(ccmbHistory.Text))
            If Not lobjRec Is Nothing Then
                fraResult.Caption = "历年体检"
                ctxtResult.Text = lobjRec("文字结论")
                fraResult.Enabled = False
            End If
            
        End If
        
    ElseIf ccmbHistory.Text = "——" Or ccmbHistory.Text = "" Then
        
        fraPicTool.Enabled = True
        fraPicTool.Caption = "描述录入"
        ctxtPResult.Text = ""
        
        fraResult.Enabled = True
        fraResult.Caption = "结论录入"
        ctxtResult.Text = ""
        
    End If
    
End Sub


'功能：清空查询人员列表
'作者：翁乔
'时间：2012-06-01
Private Sub ccmdClear_Click()
    cgrdInfoBatch.Clear
    cgrdInfoBatch.rows = 1
    cgrdInfoBatch.FormatString = "体检条码编号"
    TotalPeopleBatch.Caption = 0
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
    'lstrWhere = " and 体检类型='" & coptClasses(coptIndex).Caption & "'"
    
    '组装查询条件
    If cchkDate.Value = 1 Then
        lstrWhere = lstrWhere & " and 体检日期>='" & Format(cdtpDate.Value, "yyyy-mm-dd 00:00:00") & "' and 体检日期<='" & Format(cdtpDate.Value, "yyyy-mm-dd 23:59:59") & "'"
    End If
    
'    If cchkName.Value = 1 Then
'        If ctxtCheckName.Text = "" Then
'            MsgBox ("若要查询姓名，则姓名不能为空。")
'            Exit Sub
'        End If
'        lstrWhere = lstrWhere & " and 姓名='" & Trim(ctxtCheckName.Text) & "'"
'    End If

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
    sfsub错误处理 "职业病体检结果录入", "frmBUS_ResultInput", "ccmdQuery_Click", 6666, lstrError, False
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

'功能：查询信息
'作者：翁乔
'时间：2012-06-01
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

'功能：查询单位定位
'作者：翁乔
'时间：2012-06-01
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
    '应该把界面的相关部分清空(代码暂无)
    indX = cgrdInfo.MouseRow
    indY = cgrdInfo.MouseCol
    If indX < 0 Or indY < 0 Then
        Exit Sub
    ElseIf indX >= 0 And indX <= cgrdInfo.rows And indY >= 0 And indY < cgrdInfo.cols Then
        
        
        ccmbHistory.Enabled = True
        Cmd结论模板.Visible = True
        fraPicTool.Enabled = True
        fraPicTool.Caption = "描述录入"
        ctxtPResult.Text = ""
        
        fraResult.Enabled = True
        fraResult.Caption = "结论录入"
        ctxtResult.Text = ""
        
        ctxtBarCode.Text = cgrdInfo.TextMatrix(indX, 0)
        ctxtBarCode_KeyDown 13, 0
        '2012-07-03 于登淼 ↓
        '每次读入个人信息时，判断是否超过修改时间。
        '以此控制保存按钮是否可用。
        If pobj业务对象.sub是否在修改时间范围内(ctxtBarCode.Text, priDeptName, 8) = False Then
            ctlb工具栏.Buttons(2).Enabled = False
        End If
        '2012-07-03 于登淼 ↑
    End If
End Sub

''''功能：改变选择外观
''''作者：翁乔
''''时间：2012-06-01
'''Private Sub cgrdInfoBatch_Click()
'''    cgrdInfoBatch.SelectionMode = flexSelectionByRow
'''End Sub

'功能：读取选中编号的体检人员信息
'作者：翁乔
'时间：2012-06-01
Private Sub cgrdInfoBatch_DblClick()
    indX = cgrdInfoBatch.MouseRow
    indY = cgrdInfoBatch.MouseCol
    If indX <= 0 Or indY < 0 Then
        Exit Sub
    ElseIf indX > 0 And indX < cgrdInfoBatch.rows And indY >= 0 And indY < cgrdInfoBatch.cols Then
        ctxt体检条码.Text = cgrdInfoBatch.TextMatrix(indX, 0)
        ctxt体检条码_KeyDown 13, 0
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

'2012-07-14 于登淼
Private Sub coptClasses_Click(Index As Integer)
    Dim coptIndex As Integer
    coptIndex = Index
    sub查询列表显示 coptIndex
End Sub

'Private Sub coptResult_Click(Index As Integer)
'    If coptResult(Index).Value = True Then
'        ctxtResult.Text = coptResult(Index).Caption
'    End If
'End Sub

Private Sub ctxtBarCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not (KeyCode = 13 And Trim(ctxtBarCode.Text) <> "") Then Exit Sub
    
    Dim lstrNo As String
    Dim i As Integer
    Dim rs As Object
    Dim lcol职业病对象 As Object
    lstrNo = Trim(ctxtBarCode.Text)
    
    '检查条码号是否存在
    Dim lobjTmp As Object
    Set lobjTmp = CreateObject("职业病体检结果录入.clscommon")
    Set mlobjRec = lobjTmp.func获取体检人员基本信息(lstrNo)
    If mlobjRec.recordcount = 0 Then
        '清空当前个人信息
        ctxtBarCode.Enabled = True
        ctxtName.Text = ""
        ctxtSex.Text = ""
        ctxtAge.Text = ""
        ctxtCompanyName.Text = ""
        Exit Sub
    End If
    
    '载入已有的个人信息和现有的体检结果
    LoadPersonalInfo (lstrNo)
    
    '2012-04-15 于登淼 ↓
    '下面注释的代码，保存和读取数据类型和位置都错了
    '故仿照其它科室重新写的
''    Set lcol职业病对象 = CreateObject("职业病对象.clsManageMedicalExam")
''    Set rs = lcol职业病对象.func返回科室和图片结论(ctxtBarCode.Text, priDeptName)
''    If Not rs Is Nothing Then
''        ctxtResult.Text = rs("文字结论")
''        ctxtPResult.Text = rs("图片结论")
''    End If
    
    '2012-05-22 陶露
    '当前科室结论
    Set lcol职业病对象 = CreateObject("职业病对象.clsManageMedicalExam")
    ctxtResult.Text = lcol职业病对象.func返回科室结论(ctxtBarCode.Text, priDeptName)
    '当前科室结果(图片描述)
    Set lcol职业病对象 = CreateObject("职业病体检结果录入.clscommon")
    Set rs = lcol职业病对象.func获取体检人员单科室体检结果(ctxtBarCode.Text, priDeptName)
    If rs.recordcount > 0 And IsNull(rs("体检结果")) = False Then
        ctxtPResult.Text = rs("体检结果")
    Else
        ctxtPResult.Text = ""
    End If
    Set rs = Nothing
    '2012-05-22
    
    '2012-04-15 于登淼 ↑
    
    '一旦确定当前体检人员编号，就不能更改。除非，清空界面内容
    ctxtBarCode.Enabled = False
    ctxtName.Enabled = False
    ctxtSex.Enabled = False
    ctxtAge.Enabled = False
    ctxtCompanyName.Enabled = False
    
    '功能：菜单按钮的控制
    '作者：翁乔
    '时间：2012-06-01
    ctlb工具栏.Buttons(2).Enabled = True
    ctlb工具栏.Buttons(3).Enabled = False

    ''2012-06-27 于登淼 ↓
    '每次读入个人信息时，判断是否超过修改时间。
    '以此控制保存按钮是否可用。
    If pobj业务对象.sub是否在修改时间范围内(ctxtBarCode.Text, priDeptName, 8) = False Then
        ctlb工具栏.Buttons(2).Enabled = False
    End If
    '2012-06-27 于登淼 ↑
End Sub

'2012-06-21 于登淼
'更改当前录入状态
Private Sub ctxtPResult_Change()
    ResultChanged = 2
End Sub

'2012-06-21 于登淼
'更改当前录入状态
Private Sub ctxtResult_Change()
    ResultChanged = 2
End Sub

'功能：根据体检号查询人员信息
'作者：翁乔
'时间：2012-06-01
Private Sub ctxt体检条码_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lstrNo As String
    Dim i As Integer
    Dim str科室结论 As String
    Dim lcol职业病对象 As Object
    lstrNo = Trim(ctxt体检条码.Text)
    Dim rs As Object
    
    '检查条码号是否存在
    Dim lobjTmp As Object
    Set lobjTmp = CreateObject("职业病体检结果录入.clscommon")
    Set mlobjRec = lobjTmp.func获取体检人员基本信息(lstrNo)
    If mlobjRec.recordcount = 0 Then
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
        
'        If cchk套用体检结果.Value = 0 Then
            '2012-05-22 陶露
            '当前科室结论
            Set lcol职业病对象 = CreateObject("职业病对象.clsManageMedicalExam")
            ctxtResult.Text = lcol职业病对象.func返回科室结论(ctxt体检条码.Text, priDeptName)
            '当前科室结果(图片描述)
            Set lcol职业病对象 = CreateObject("职业病体检结果录入.clscommon")
            Set rs = lcol职业病对象.func获取体检人员单科室体检结果(ctxt体检条码.Text, priDeptName)
            If rs.recordcount > 0 And IsNull(rs("体检结果")) = False Then
                ctxtPResult.Text = rs("体检结果")
            Else
                ctxtPResult.Text = ""
            End If
            Set rs = Nothing
            '2012-05-22
'        End If

        '一旦确定当前体检人员编号，就不能更改。除非，清空界面内容。
        ctxt体检条码.Enabled = False
        ctxt姓名.Enabled = False
        ctxt性别.Enabled = False
        ctxt年龄.Enabled = False
        ctxt单位名称.Enabled = False '其实单位灰掉了之后，如果有“单位定位”按钮，还是可以改的。
'''        For i = 0 To 2
'''            If coptClasses(i).Value = False Then coptClasses(i).Enabled = False
'''        Next i
        ctlb工具栏.Buttons(2).Enabled = False
        ctlb工具栏.Buttons(3).Enabled = True
    Else
        Set lobjTmp = Nothing
        MsgBox ("该体检人员没有该科室的体检项目！")
        cgrdInfoBatch.RemoveItem
        subClear
    End If
    
End Sub

Private Sub Form_Activate()
    '2012-05-24 于登淼 ↓
    'ctxtBarCode激活焦点前先必须可用
    ctxtBarCode.Enabled = True
    '2012-05-24 于登淼 ↑
    ctxtBarCode.SetFocus    '激活焦点首先是体检条码号
    ctxtBarCode.SelStart = Len(ctxtBarCode)
    ctxtBarCode.SelLength = 0
    cgrdInfo.SelectionMode = flexSelectionByRow
    cgrdInfo.AllowSelection = False
    
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
        .Add "批量保存(&D)"
        .Add "删除"
        .Add "网络配置(&S)111"
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
    
    ctlb工具栏.Buttons(2).Enabled = False
    ctlb工具栏.Buttons(3).Enabled = False
    ctlb工具栏.Buttons(4).Visible = False
    
    '创建本窗体的全局对象mobj体检医师。
    Set mobj体检医师 = CreateObject("职业病对象.clsMedicalExaminer")
    mobj体检医师.编号 = um用户编号
    
    '得到医师名字，为当前用户名
    mstrDoctorName = um用户名
    LabelDoctor.Caption = LabelDoctor.Caption & " " & mstrDoctorName
    
    '界面权限设置。仅限于该界面上各个按钮盒其它控件的使用
    '设置的功能暂时有：查看、修改、删除、打印、网络配置。（有点儿多啊）
    Dim lobjTmp As Object
    Set lobjTmp = CreateObject("职业病设置.clspermissionconfigure")
    If lobjTmp.func科室操作权限(um用户编号, "职业病体检_B超影像科结果录入_修改") = False Then
        ctlb工具栏.Buttons(2).Visible = False
    End If
    
    If lobjTmp.func科室操作权限(um用户编号, "职业病体检_B超影像科结果录入_删除") = False Then
        ctlb工具栏.Buttons(4).Visible = False
    End If
     
    If lobjTmp.func科室操作权限(um用户编号, "职业病体检_B超影像科结果录入_网络配置") = False Then
        ctlb工具栏.Buttons(5).Visible = False
    End If
    
    '2012-05-22 翁乔 ↓↓↓
    '界面权限设置
    If lobjTmp.func科室操作权限(um用户编号, "职业病体检_B超影像科结果录入_批量修改") = False Then
        ctlb工具栏(3).Visible = False
    End If
    '2012-05-22 ↑↑↑
    Set lobjTmp = Nothing
    
    'form_load 时，界面按钮操作设定
    cdtpConclusionDate.Value = Now
    cdtpDate.Value = Now
    DTP录入日期.Value = Now
    cdtpDateBatch.Value = Now
    
    '2012-04-15 于登淼 ↓
    'dicom控件初始化
    DCMList.Path = ""
    DCMList.Enabled = False
    'DCMList.ListCount = 0
    '2012-04-15 于登淼 ↑

    '2012-06-21 于登淼 ↓
    '省疾控新要求中改变系统编号规则。
    '获取系统编号固定部分。
    sub获取系统编号固定部分
    '2012-06-21 于登淼 ↑
    
    '2012-06-21 于登淼 ↓
    '初始化当前录入状态(提前判断有无权限修改，若无，直接赋值为3)
    ResultChanged = IIf(ResultChanged <> 3, -1, 3)
    cchk刷条码_Click
    '2012-06-21 于登淼 ↑
    
    '2012-07-14 于登淼 ↓
    '初始化查询界面，调整查询列表格式。初始化科室基本信息。
    priDeptName = "B超影像科"
    priDeptNo = "11"
    priDeptResultName = "B超影像结果"
    ccmdQuery_Click
    SSTPersonalInfo.Tab = 1: ccmdSelInfo_Click: SSTPersonalInfo.Tab = 0
    coptClasses_Click (0)
    '2012-07-14 于登淼 ↑
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病体检结果录入", "frmBUS_ResultInput", "Form_Load", 6666, lstrError, False
    Exit Sub
    Resume
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    mblnInUse = False
    Set mobjGUI = Nothing
    Set mobjQueryResult = Nothing
End Sub

Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    On Error GoTo errHandler
    Dim i As Integer
    Cancel = True
    
    Select Case Operate
    Case "清空界面"
        subClear
    '功能：添加菜单新的功能
    '作者：翁乔
    '时间：2012-06-01
    Case "批量保存"
        '2012-07-13 于登淼 ↓
        '如果没有体检项目，则直接退出，不保存。
        If cgrdInfoBatch.rows <= 1 Then Exit Sub
        '2012-07-13 于登淼 ↑
        
        '2012-07-15 于登淼 ↓
        '没有录入体检结论时，提示且不保存。
        If Len(Trim(ctxtResult.Text)) = 0 Then
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
        
    '时间：2012-06-01
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
        If Len(Trim(ctxtResult.Text)) = 0 Then
            MsgBox "你还没有为病人下结论"
            GoTo errHandler
        End If
        '2012-07-15 于登淼 ↑
        
        Dim lstrCheck As String
        Dim lobjTmp As Object
        Dim isOk As Integer
        
        '录入结果部分暂时不能操作
        fraResult.Enabled = False
        
        Set lcolResult = New Collection
        Set lcolItem = New Collection
        
        '2012-04-15 于登淼 ↓
        '下面注释掉的代码，保存位置写错了。故重写
''        '保存B超影像科体检结果
''        If SSTPersonalInfo.Tab = 0 Then
''            lstrCheck = sub添加单项结果(ctxtResult.Text, priDeptResultName, lstrCheck)
''        End If
''
''        '功能：保存单个项目的医生结论
''        '作者：翁乔
''        '时间：2012-04-14
''        pobj业务对象.sub单个填写体检结论和图片描述 ctxtBarCode.Text, priDeptName, ctxtPResult.Text, ctxtResult.Text, um用户编号
''        '作者：翁乔
''        '时间：2012-04-14
        
        '保存图片描述，也就是体检结果
        Call sub添加单项结果(ctxtPResult.Text, priDeptResultName, "")
        
        '保存科室结论
        Dim lobjTmp2 As Object
        Call pobj业务对象.sub单个填写体检结论(ctxtBarCode.Text, priDeptName, ctxtResult.Text, um用户编号)
        Set lobjTmp2 = Nothing
        '2012-04-15 于登淼 ↑
        
        'lstrCheck字符串检查
        If (Not lstrCheck = "") Then
            isOk = MsgBox("以下项目未填写结果，确定保存吗？" & Chr(10) & "未填写项将不会记录到数据库！" & Chr(10) & Chr(10) & Trim(lstrCheck), vbOKCancel)
            If isOk = 2 Then
                Set lcolResult = Nothing
                Set lcolItem = Nothing
                Exit Sub
            End If
        End If
        
        fraResult.Enabled = True
        '2012-07-03 于登淼 ↓
        '增加一个字段"修改起始时间"的修改。同时修改该科室的体检结果录入状态。
        pobj业务对象.sub修改起始时间 Trim(ctxtBarCode.Text), priDeptName
        pobj业务对象.sub修改结果录入状态 Trim(ctxtBarCode.Text), priDeptNo, "2"  '11为B超影像科
        pobj业务对象.sub结果录入修改体检状态 Trim(ctxtBarCode.Text), "4"
        '2012-07-03 于登淼 ↑
        
        subSave
        
        '2012-07-15 于登淼 ↓
        '保存完之后，重新进行查询。
        ccmdQuery_Click
        i = SSTPersonalInfo.Tab
        SSTPersonalInfo.Tab = 1: ccmdSelInfo_Click: SSTPersonalInfo.Tab = i
        '2012-07-15 于登淼 ↑
        
        Set lcolResult = Nothing
        Set lcolItem = Nothing
    Case "删除"
        '
    Case "打印"
        '
    Case "网络配置"
        '
    Case "退出"
        '2012-06-21 于登淼 ↓
        '退出时增加判断是否保存
        ctxtBarCode.Enabled = True
        ctxtBarCode.SetFocus
        ctxtBarCode.Enabled = False
        Dim isSave As Integer
        If ResultChanged = 2 Or ResultChanged = 0 Then
            '修改：如果处于病历查看、则退出不提醒。（翁乔，2012-08-01）
'            If Trim(Frame6.Caption) <> "体检项目结果填写：" Then
'                Unload Me
'                Exit Sub
'            End If
            isSave = MsgBox("是否保存已修改结果？", vbYesNoCancel)
            If isSave = vbCancel Then Exit Sub
            If isSave = vbYes Then mobjGUI_BeforeOperate "保存", False
        End If
        '2012-06-21 于登淼 ↑
        Unload frmBUS_ResultInput
        Set frmBUS_ResultInput = Nothing
    End Select
    
    Exit Sub
errHandler:
    If Err.Number = 0 Then Exit Sub
    sfsub错误处理 "职业病体检结果录入", "frmBUS_ResultInput", "mobjGUI_BeforeOperate", Err.Number, Err.Description, False
    Cancel = True
    Exit Sub
    Resume
End Sub

Sub LoadPersonalInfo(ByVal paraSysNo As String)
    On Error GoTo errHandler
    Dim i As Integer
    Dim lobjTmp, lobjRec As Object
    
    Set lobjTmp = CreateObject("职业病体检结果录入.clscommon")
    Set mlobjRec = lobjTmp.func获取体检人员基本信息(paraSysNo)
    If mlobjRec.recordcount > 0 Then
        ctxtName = mlobjRec("姓名")
        ctxtSex = mlobjRec("性别")
        ctxtAge = mlobjRec("年龄")
        ctxtCompanyName = mlobjRec("单位名称")
        
'''        '设置体检类型
'''        If mlobjRec("体检类型") = "职业健康" Then coptClasses(0).Value = True
'''        If mlobjRec("体检类型") = "放射工作" Then coptClasses(1).Value = True
'''        If mlobjRec("体检类型") = "涉核部队" Then coptClasses(2).Value = True
        
        '显示照片
        Set lobjRec = CreateObject("职业病对象.clspersonexamed")
        lobjRec.系统编号 = ctxtBarCode.Text
        Picture2.Picture = lobjRec.像片
        
        '显示病人的历年病历。（翁乔；2012-07-31）↓↓↓↓↓↓↓↓↓↓↓↓
            Dim lobjDatecobo As Object
            Set lobjDatecobo = mobj体检医师.func获取体检人员的体检病历(Trim(ctxtBarCode.Text), "B超影像科")
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
                
            End If
'            ccmbHistory.ListIndex = 0
            
            '显示病人的历年病历。（翁乔；2012-07-31） ↑↑↑↑↑↑↑↑↑↑↑↑
        
        Set lobjRec = lobjTmp.func是否已经体检过(ctxtBarCode.Text, priDeptName)
        
        If lobjRec.recordcount > 0 Then     '暂没有写，如果填写结果后修改的标记--------------
            sub填写已有的体检结果 lobjRec
            sub载入该人员DICOM图片 ctxtBarCode.Text
        Else
            sub清空当前结果
        End If
    Else
        MsgBox ("没有该条码对应的体检人员信息！")
        Exit Sub
    End If
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病体检结果录入", "frmBUS_ResultInput", "LoadPersonalInfo", 6666, lstrError, False
End Sub

Private Function sub添加单项结果(ByVal paraResult As String, ByVal paraItem As String, ByVal paraCheck As String) As String
    If paraResult = "" Then
        paraCheck = paraCheck & IIf(paraCheck = "", "", Chr(10) & paraItem)
    Else
        lcolItem.Add paraItem
        lcolResult.Add paraResult
    End If
    sub添加单项结果 = paraCheck
End Function

Sub subSaveBatch(ByVal para系统编号 As String)
    On Error GoTo errHandler
    
    Dim isOk As Boolean
    Dim lstrTmp As String
    Dim lobjTmp As Object
    Set lobjTmp = CreateObject("职业病体检结果录入.clscommon")
    isOk = lobjTmp.func保存单人体检结果(para系统编号, mstrDoctorName, cdtpConclusionDate.Value, lcolItem, lcolResult, "职业病体检_结果信息_B超影像科")
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病体检结果录入", "frmBUS_ResultInput", "subSave", 6666, lstrError, False
End Sub

Sub subSave()
    On Error GoTo errHandler
    
    Dim isOk As Boolean
    Dim lstrTmp As String
    Dim lobjTmp As Object
    Set lobjTmp = CreateObject("职业病体检结果录入.clscommon")
    isOk = lobjTmp.func保存单人体检结果(ctxtBarCode.Text, mstrDoctorName, cdtpConclusionDate.Value, lcolItem, lcolResult, "职业病体检_结果信息_B超影像科")
    subClear
    If isOk = True Then MsgBox ("保存成功！")
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病体检结果录入", "frmBUS_ResultInput", "subSave", 6666, lstrError, False
End Sub

Sub sub填写已有的体检结果(ByVal paraRec As Object)
    paraRec.movefirst
    If IsNull(paraRec("体检结果")) = True Then
        ctxtResult.Text = ""
        '2012-06-21 于登淼 ↓
        '设置当前录入状态(已经录入过，正在修改当前结果)
        ResultChanged = 0
        '2012-06-21 于登淼 ↑
    Else
        ctxtResult.Text = paraRec("体检结果")
        '2012-06-21 于登淼 ↓
        '设置当前录入状态(已经录入过，正在修改当前结果)
        ResultChanged = 1
        '2012-06-21 于登淼 ↑
    End If
End Sub

Sub sub清空当前结果()
    ctxtResult.Text = ""
    
    '清空当前图片结果
    '-------------↓↓↓暂无↓↓↓--------------
    
    '-------------↑↑↑暂无↑↑↑--------------
End Sub
Sub sub载入该人员DICOM图片(ByVal paraSysNo As String)
    '-------------↓↓↓暂无↓↓↓--------------
    
    '-------------↑↑↑暂无↑↑↑--------------
End Sub

'2012-04-14 于登淼
'打开某个dicom图片，记录路径和当前文件名，文件夹路径
Private Sub ccmdDCMOpen_Click()
    Diag.ShowOpen
    DCMPath = Diag.FileName
    DCMFileName = Diag.FileTitle
    DCMDir = Replace(DCMPath, "\" & DCMFileName, "")
    
    On Error GoTo errHandler
    Dicm.OpenFile (DCMPath)
    'Dicm.OpenFileNameByMultiple = DCMPath
    subInitDCMFileList
    
    Exit Sub
errHandler:
    On Error GoTo errHandler2
    Dicm.OpenFile (DCMFileName)
    subInitDCMFileList
    Exit Sub
errHandler2:
    MsgBox ("文件读取出错，请稍后重试。")
End Sub

'2012-04-14 于登淼
'保存当前dicom图片。目前保存方法为：在文件名前面加上当前修改日期后，再加上.dcm后缀
'注意：保存文件默认为当前图片文件所在目录下，所以只需填入文件名即可。不能有'/'、'\'等特殊字符。
Private Sub ccmdSavePic_Click()
    If cchkReplace.Value = 1 Then
        Dicm.ImageSaveToDICOM = DCMList.List(DCMIdx)   '替换原有文件(不推荐)
    Else
        Dicm.ImageSaveToDICOM = Replace(DCMList.List(DCMIdx), ".dcm", "") & "_" & Format(Date, "yyyymmdd") & ".dcm"
    End If
End Sub

'2012-04-14 于登淼
'单击文件列表，更新当前显示图片
Private Sub DCMList_Click()
    '出错提示在dicm控件中，无法在代码中控制。故这里省略错误处理。
    '同时，出错原因为原图像数据格式错误
    DCMIdx = DCMList.ListIndex
    DCMPath = DCMDir & "\" & DCMList.List(DCMIdx)
    '如果加上鼠标指针和enable控制，滚动条会失效
    'DCMList.Enabled = False
    'MousePointer = 11
    Dicm.OpenFile (DCMPath)
    llabCurr = "第" & (DCMIdx + 1) & "/共" & DCMList.ListCount
'    Timer1_Timer
    'DCMList.Enabled = True
    'MousePointer = 1
End Sub

'2012-04-14 于登淼
'初始化dicom文件列表
Sub subInitDCMFileList()
    DCMList.Enabled = True
    'DCMList.Pattern = "*.*"
    DCMList.Path = DCMDir
    Dim i As Integer
    For i = 0 To DCMList.ListCount - 1
        If DCMFileName = DCMList.List(i) Then
            DCMIdx = i
            Exit For
        End If
    Next
    DCMList.ListIndex = DCMIdx
End Sub

'功能：批量界面读取个人信息
'作者：翁乔
'时间：2012-06-01
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
        
        '载入已有的个人信息和现有的体检结果
        '显示照片
        Set lobjRec = CreateObject("职业病对象.clspersonexamed")
        lobjRec.系统编号 = ctxt体检条码.Text
        Picture4.Enabled = True
        Picture4.Visible = True
        Picture4.Picture = lobjRec.像片
            
        Set lobjRec = lobjTmp.func是否已经体检过(ctxt体检条码.Text, priDeptName)
        If lobjRec.recordcount = 0 Then
            If ResultChanged <> 3 Then ResultChanged = 0
        ElseIf lobjRec.recordcount > 0 Then
            If ResultChanged <> 3 Then ResultChanged = 1
        End If
    Else
        MsgBox ("没有该条码对应的体检人员信息！")
        Exit Sub
    End If
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病体检结果录入", "frmENT_ResultInput", "LoadPersonalInfo", 6666, lstrError, False
End Sub

'功能：批量保存结果
'作者：翁乔
'时间：2012-06-01
Sub sub批量保存()

    MousePointer = 11
    Dim lblnNotOver As Boolean
    Dim i As Integer
    Dim barCode As Collection '批量保存体检条码
        'cstbMain.Panels(1) = "正在保存，请稍候..."
        
        '暂时界面不能操作。
        Frame1.Enabled = False
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
        
        If cgrdInfoBatch.rows < 1 Then
        MsgBox ("请确认录入人员数目是否正确！")
        Exit Sub
    End If
    Dim ccrpValue As Integer
    Dim ccrpI As Integer
    Dim isOk As Boolean
    Dim lstrTmp As String
    Dim lobjTmp As Object
    
    '显示保存进度条
    ccrpI = barCode.Count
    ccrp进度.Max = ccrpI
    ccrp进度.Visible = True
    ccrp进度.Caption = "0%"
    ccrp进度.Value = 0
    For i = 1 To barCode.Count
        '保存图片描述，也就是体检结果
        Call sub添加单项结果(ctxtPResult.Text, priDeptResultName, "")
            
        '保存科室结论
        Dim lobjTmp2 As Object
        Call pobj业务对象.sub单个填写体检结论(barCode(i), priDeptName, ctxtResult.Text, um用户编号)
        Set lobjTmp2 = Nothing
        '2012-04-15 于登淼 ↑
        
        ccrp进度.Caption = Int(i / ccrp进度.Max * 100) + ccrpValue & "%"
        ccrp进度.Value = ccrp进度.Value + 1
        
        '2012-07-03 于登淼 ↓
        '增加一个字段"修改起始时间"的修改。同时修改该科室的体检结果录入状态。
        pobj业务对象.sub修改起始时间 barCode(i), priDeptName
        pobj业务对象.sub修改结果录入状态 barCode(i), priDeptNo, "2"
        pobj业务对象.sub结果录入修改体检状态 barCode(i), "4"
        '2012-07-03 于登淼 ↑
        
        subSaveBatch barCode(i)
        
    Next i
    MsgBox ("批量保存成功！")
    subClear
    
    ccrp进度.Visible = False

    MousePointer = 0
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func错误处理(Err.Number, Err.Description)
    sfsub错误处理 "职业病体检结果录入", "FrmENT_ResultInput", "ccmdSave_Click", 6666, lstrError, False

End Sub

'功能：清空界面功能
'作者：翁乔
'时间：2012-06-01
Sub subClear()
    TotalPeople.Caption = 0
    TotalPeopleBatch.Caption = 0
    
    '清空当前个人信息
    ctxtBarCode.Text = ""
    ctxtBarCode.Enabled = True
    ctxtName.Text = ""
    ctxtSex.Text = ""
    ctxtAge.Text = ""
    ctxtCompanyName.Text = ""
    cgrdInfo.rows = 1
    cchkDate.Value = 0
    cchkName.Value = 0
    'cchkFilledResult.Value = 0
    'cchkUnfilledResult.Value = 1
    cdtpDate.Value = Now
    ctxtCheckName.Text = ""
    
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
    ctxtResult.Text = ""
    ctxtPResult.Text = ""
    
    cchkDateBatch.Value = 0
    cchkCompanyBatch.Value = 0
    TotalPeopleBatch.Caption = "0"
    
'清空照片
    Set Picture2.Picture = Nothing
    Set Picture4.Picture = Nothing
    
    '恢复为form_load时的状态。
    If SSTPersonalInfo.Tab = 0 Then
        ctxtBarCode.Enabled = True
        ctxtBarCode.SetFocus
        ctxtName.Enabled = True
        ctxtSex.Enabled = True
        ctxtAge.Enabled = True
        ctxtCompanyName.Enabled = True
    Else
        ctxt体检条码.Enabled = True
        ctxt姓名.Enabled = True
        ctxt性别.Enabled = True
        ctxt年龄.Enabled = True
        ctxt单位名称.Enabled = True
    End If
    
    sub清空当前结果
    
    '2012-04-15 于登淼 ↓
    '控制dicom图像文件列表等
    DCMList.Enabled = False
    
'''    coptClasses(0).Enabled = True
'''    coptClasses(1).Enabled = True
'''    coptClasses(2).Enabled = True
    ctlb工具栏.Enabled = True
    SSTPersonalInfo.Enabled = True
    Frame1.Enabled = True
'''    coptClasses(0).Value = 1
    ctlb工具栏.Buttons(1).Enabled = True
    ctlb工具栏.Buttons(2).Enabled = False
    ctlb工具栏.Buttons(3).Enabled = False
    
    '2012-06-21 于登淼 ↓
    '初始化当前录入状态(提前判断有无权限修改，若无，直接赋值为3)
    ResultChanged = IIf(ResultChanged <> 3, -1, 3)
    cchk刷条码_Click
    '2012-06-21 于登淼 ↑
    
End Sub

'2012-06-21 于登淼
Sub sub获取系统编号固定部分()
    '获取服务器日期
    Dim lobjRec As Object
    Set lobjRec = dafuncGetData("select getdate()")
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
        
    End If 'mobjQueryResult.recordcount > 0
    
    If SSTPersonalInfo.Tab = 0 Then
        With cgrdInfo
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
