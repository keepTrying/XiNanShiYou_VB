VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3A75EE84-8E64-43F4-A904-E4835B9A3DB4}#3.9#0"; "DICOMax.ocx"
Object = "{FC07EBD4-FE92-11D0-A199-A0077383D901}#5.1#0"; "CCRPPRG.OCX"
Begin VB.Form frmBUS_ResultInput 
   Caption         =   "B��Ӱ��ƽ��¼�봰��"
   ClientHeight    =   9735
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14595
   LinkTopic       =   "Form1"
   ScaleHeight     =   9735
   ScaleWidth      =   14595
   StartUpPosition =   2  '��Ļ����
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
         Caption         =   "8023����"
         Height          =   255
         Index           =   4
         Left            =   4440
         TabIndex        =   74
         Top             =   840
         Width           =   1095
      End
      Begin VB.OptionButton coptClasses 
         BackColor       =   &H00C0FFFF&
         Caption         =   "��˲���"
         Height          =   255
         Index           =   3
         Left            =   3360
         TabIndex        =   73
         Top             =   840
         Width           =   1095
      End
      Begin VB.OptionButton coptClasses 
         BackColor       =   &H00C0FFFF&
         Caption         =   "��ͨ���"
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
         Caption         =   "ְҵ����"
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   71
         Top             =   840
         Width           =   1095
      End
      Begin VB.OptionButton coptClasses 
         BackColor       =   &H00C0FFFF&
         Caption         =   "���佡��"
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
         Begin VB.CheckBox cchkˢ���� 
            Caption         =   "ˢ����"
            Height          =   255
            Left            =   11760
            TabIndex        =   69
            Top             =   120
            Width           =   1215
         End
         Begin VB.Frame fraPicShow 
            Caption         =   "ͼƬ��ʾ��"
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
               Caption         =   "��ͼƬ�ļ�"
               Height          =   495
               Left            =   5640
               TabIndex        =   13
               Top             =   600
               Width           =   1815
            End
            Begin VB.CommandButton ccmdSavePic 
               Caption         =   "���浱ǰͼƬ"
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
               Caption         =   "����ʱ����"
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
               Caption         =   "�ڣ���/������"
               Height          =   255
               Left            =   5880
               TabIndex        =   15
               Top             =   5400
               Width           =   1455
            End
            Begin VB.Label Label12 
               BackColor       =   &H00FFC0C0&
               Caption         =   "ͼƬ�ϰ�ס���������Ҽ��϶����Ըı�ͼ��Աȶ�"
               Height          =   255
               Left            =   120
               TabIndex        =   14
               Top             =   240
               Width           =   4215
            End
         End
         Begin VB.Frame fraResult 
            Caption         =   "����¼��"
            Height          =   855
            Left            =   6720
            TabIndex        =   5
            Top             =   1800
            Width           =   7575
            Begin VB.CommandButton Cmd����ģ�� 
               Caption         =   "����ģ��"
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
               Caption         =   "ҽʦ����"
               Height          =   255
               Left            =   120
               TabIndex        =   7
               Top             =   240
               Width           =   855
            End
         End
         Begin VB.Frame fraPicTool 
            Caption         =   "����¼��"
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
               Caption         =   "ͼƬ����"
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
            TabCaption(0)   =   "����¼��"
            TabPicture(0)   =   "frmBUS_ResultInput.frx":0000
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "fraQuery"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "fraInfo"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).ControlCount=   2
            TabCaption(1)   =   "����¼��"
            TabPicture(1)   =   "frmBUS_ResultInput.frx":001C
            Tab(1).ControlEnabled=   -1  'True
            Tab(1).Control(0)=   "TotalPeopleBatch"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).Control(1)=   "Label6"
            Tab(1).Control(1).Enabled=   0   'False
            Tab(1).Control(2)=   "ccrp����"
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
            Tab(1).Control(9)=   "ccmd��ѯ��λ"
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
               Caption         =   "δ����"
               Height          =   255
               Index           =   1
               Left            =   3120
               TabIndex        =   79
               Top             =   4560
               Value           =   1  'Checked
               Width           =   1095
            End
            Begin VB.CheckBox cchkBchResult 
               Caption         =   "������"
               Height          =   255
               Index           =   0
               Left            =   1800
               TabIndex        =   78
               Top             =   4560
               Width           =   1095
            End
            Begin VB.CommandButton ccmdRemove 
               Caption         =   "�� ��"
               Height          =   375
               Left            =   4800
               TabIndex        =   64
               Top             =   4440
               Width           =   855
            End
            Begin VB.CommandButton ccmdClear 
               Caption         =   "�� ��"
               Height          =   375
               Left            =   4800
               TabIndex        =   63
               Top             =   4920
               Width           =   855
            End
            Begin VB.CheckBox cchkDateBatch 
               BackColor       =   &H00C0FFC0&
               Caption         =   "�������"
               Height          =   255
               Left            =   480
               TabIndex        =   62
               Top             =   3480
               Width           =   1215
            End
            Begin VB.Frame fraQueryBatch 
               Caption         =   "������ѯ�����Ա"
               Height          =   2895
               Left            =   240
               TabIndex        =   46
               Top             =   480
               Width           =   5775
               Begin VB.CommandButton ccmdLocateBatch 
                  Caption         =   "��λ��λ"
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
               Begin VB.TextBox ctxt��λ���� 
                  Height          =   300
                  Left            =   1320
                  TabIndex        =   52
                  Top             =   2160
                  Width           =   2415
               End
               Begin VB.TextBox ctxt���� 
                  Height          =   300
                  Left            =   1320
                  TabIndex        =   51
                  Top             =   1800
                  Width           =   2415
               End
               Begin VB.TextBox ctxt�Ա� 
                  Height          =   300
                  Left            =   1320
                  TabIndex        =   50
                  Top             =   1440
                  Width           =   2415
               End
               Begin VB.TextBox ctxt���� 
                  Height          =   300
                  Left            =   1320
                  TabIndex        =   49
                  Top             =   1080
                  Width           =   2415
               End
               Begin VB.TextBox ctxt������� 
                  Height          =   300
                  Left            =   1320
                  TabIndex        =   48
                  Top             =   720
                  Width           =   2415
               End
               Begin VB.CheckBox cchk��������� 
                  BackColor       =   &H008080FF&
                  Caption         =   "�������Ա�����Ϊ���������¼��"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   47
                  Top             =   2520
                  Value           =   1  'Checked
                  Width           =   3615
               End
               Begin MSComCtl2.DTPicker DTP¼������ 
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
                  Caption         =   "����¼������"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   61
                  Top             =   360
                  Width           =   1095
               End
               Begin VB.Label Label14 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "��λ����"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   60
                  Top             =   2160
                  Width           =   975
               End
               Begin VB.Label Label15 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "����"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   59
                  Top             =   1800
                  Width           =   975
               End
               Begin VB.Label Label16 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "�Ա�"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   58
                  Top             =   1440
                  Width           =   975
               End
               Begin VB.Label Label17 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "����"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   57
                  Top             =   1080
                  Width           =   975
               End
               Begin VB.Label Label18 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "��������"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   56
                  Top             =   720
                  Width           =   975
               End
            End
            Begin VB.CommandButton ccmd��ѯ��λ 
               Caption         =   "��λ��λ"
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
               Caption         =   "��λ����"
               Height          =   255
               Left            =   480
               TabIndex        =   43
               Top             =   3960
               Width           =   1215
            End
            Begin VB.CommandButton ccmdSelInfo 
               Caption         =   "�� ѯ"
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
               Caption         =   "��ѯ�����Ա"
               Height          =   4815
               Left            =   -74880
               TabIndex        =   32
               Top             =   3240
               Width           =   6255
               Begin VB.CommandButton ccmdWork 
                  Caption         =   "��λ��λ"
                  Height          =   375
                  Left            =   3240
                  TabIndex        =   87
                  Top             =   960
                  Width           =   1185
               End
               Begin VB.CheckBox cchkSingleNo 
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "�������"
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
                  Caption         =   "���֤��"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   84
                  Top             =   600
                  Width           =   1095
               End
               Begin VB.CheckBox cchkWorkUnit 
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "��λ����"
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
                  Caption         =   "δ����"
                  Height          =   255
                  Index           =   1
                  Left            =   1440
                  TabIndex        =   77
                  Top             =   1320
                  Value           =   1  'Checked
                  Width           =   1095
               End
               Begin VB.CommandButton ccmdQuery 
                  Caption         =   "��   ѯ"
                  Height          =   375
                  Left            =   4800
                  TabIndex        =   37
                  Top             =   960
                  Width           =   1215
               End
               Begin VB.CheckBox cchkSigResult 
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "������"
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
                  Caption         =   "����"
                  Height          =   255
                  Left            =   3120
                  TabIndex        =   34
                  Top             =   600
                  Width           =   735
               End
               Begin VB.CheckBox cchkDate 
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "�������"
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
                  ToolTipText     =   "˫���Զ����������Ϣ�����������"
                  Top             =   1680
                  Width           =   6015
                  _cx             =   2088774002
                  _cy             =   2088768710
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
                  Left            =   6000
                  TabIndex        =   76
                  Top             =   1440
                  Width           =   90
               End
               Begin VB.Label Label9 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "������"
                  Height          =   180
                  Left            =   5280
                  TabIndex        =   75
                  Top             =   1440
                  Width           =   540
               End
            End
            Begin VB.Frame fraInfo 
               Caption         =   "������Ϣ"
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
                  Caption         =   "��λ��λ"
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
                  Caption         =   "���겡��"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   89
                  Top             =   600
                  Width           =   975
               End
               Begin VB.Label Label7 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "����¼������"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   31
                  Top             =   240
                  Width           =   1095
               End
               Begin VB.Label Label5 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "��λ����"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   30
                  Top             =   2400
                  Width           =   975
               End
               Begin VB.Label Label4 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "����"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   29
                  Top             =   2040
                  Width           =   975
               End
               Begin VB.Label Label3 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "�Ա�"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   28
                  Top             =   1680
                  Width           =   975
               End
               Begin VB.Label Label2 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "����"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   27
                  Top             =   1320
                  Width           =   975
               End
               Begin VB.Label Label1 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "��������"
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
            Begin CCRProgressBar.ccrpProgressBar ccrp���� 
               Height          =   375
               Left            =   600
               Top             =   5040
               Visible         =   0   'False
               Width           =   3615
               _ExtentX        =   6376
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
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0E0FF&
               Caption         =   "������"
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
         Begin MSComctlLib.Toolbar ctlb������ 
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
            ImageList       =   "cimg��ťͼ��"
            _Version        =   393216
         End
         Begin VB.Label LabelDoctor 
            BackColor       =   &H00C0FFFF&
            Caption         =   "ҽ����"
            Height          =   255
            Left            =   5520
            TabIndex        =   66
            Top             =   840
            Width           =   1095
         End
      End
      Begin MSComctlLib.ImageList cimg��ťͼ�� 
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
'2012-03-01 �ڵ��
'���� ��ٿƽ��¼�봰�壬����Ӧ��������

Option Explicit
Public mblnInUse As Boolean
Private WithEvents mobjGUI As cls����ͨ�ö���
Attribute mobjGUI.VB_VarHelpID = -1
Private mstr��쵥�� As String
Private mstrϵͳ��� As String
Private mobj���ҽʦ  As Object   'clsMedicalExamer    ��ȡ��ǰ���ҽʦ��������ָ�����ԣ�����/���飩�������Ŀ
Private mlobjRec As Object

'��ѯ���
Private mstrDoctorName As String
Private mobjQueryResult As Object
Private mcolIndex As New Collection
Private indX, indY As Integer
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

Private mstrState As String     '��¼��ǰ���״̬

'2012-04-14 �ڵ�� ��
'dicom��ر���
Private DCMPath As String       'dicom����·��
Private DCMDir As String        'dicom�ļ���·��
Private DCMFileName As String   'dicom��ǰ�ļ���
Private DCMIdx As Integer       'dicom��ǰ�ļ�λ��(��DCMList�е�)
'2012-04-14 �ڵ�� ��

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

'2012-06-21 �ڵ��
'���ˢ�����ж�
Private Sub cchkˢ����_Click()
    If Not cchkˢ����.Visible Then Exit Sub
    If ctxtBarCode.Enabled = False Then Exit Sub
    
    If SSTPersonalInfo.Tab = 0 Then
        ctxtBarCode.Text = ""
        If cchkˢ����.Value = 0 Then sub��ȡϵͳ��Ź̶�����
        ctxtBarCode.Enabled = True
        ctxtBarCode.SetFocus
        ctxtBarCode.SelStart = Len(ctxtBarCode)
        ctxtBarCode.SelLength = 0
    Else
        ctxt�������.Text = ""
        ctxt�������.SetFocus
    End If
End Sub

'��ʾѡ�����ڵĲ�����Ϣ
'����
'2012-07-31
Private Sub ccmbHistory_Click()
    Dim lobjRec As Object

    If ccmbHistory.Text <> "����" Then
        ctlb������.Buttons(2).Enabled = False
        Set lobjRec = mobj���ҽʦ.func��ȡָ����ݵ��������(Trim(ctxtBarCode.Text), ccmbHistory.Text, "B��Ӱ���")
        
        If Not lobjRec Is Nothing Then
            
            fraPicTool.Caption = "�������"
            ctxtPResult.Text = lobjRec("�����")
            fraPicTool.Enabled = False
            
            Set lobjRec = mobj���ҽʦ.func��ȡָ����ݵ���첡������(Trim(ctxtBarCode.Text), "11", Trim(ccmbHistory.Text))
            If Not lobjRec Is Nothing Then
                fraResult.Caption = "�������"
                ctxtResult.Text = lobjRec("���ֽ���")
                fraResult.Enabled = False
            End If
            
        End If
        
    ElseIf ccmbHistory.Text = "����" Or ccmbHistory.Text = "" Then
        
        fraPicTool.Enabled = True
        fraPicTool.Caption = "����¼��"
        ctxtPResult.Text = ""
        
        fraResult.Enabled = True
        fraResult.Caption = "����¼��"
        ctxtResult.Text = ""
        
    End If
    
End Sub


'���ܣ���ղ�ѯ��Ա�б�
'���ߣ�����
'ʱ�䣺2012-06-01
Private Sub ccmdClear_Click()
    cgrdInfoBatch.Clear
    cgrdInfoBatch.rows = 1
    cgrdInfoBatch.FormatString = "���������"
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
    'lstrWhere = " and �������='" & coptClasses(coptIndex).Caption & "'"
    
    '��װ��ѯ����
    If cchkDate.Value = 1 Then
        lstrWhere = lstrWhere & " and �������>='" & Format(cdtpDate.Value, "yyyy-mm-dd 00:00:00") & "' and �������<='" & Format(cdtpDate.Value, "yyyy-mm-dd 23:59:59") & "'"
    End If
    
'    If cchkName.Value = 1 Then
'        If ctxtCheckName.Text = "" Then
'            MsgBox ("��Ҫ��ѯ����������������Ϊ�ա�")
'            Exit Sub
'        End If
'        lstrWhere = lstrWhere & " and ����='" & Trim(ctxtCheckName.Text) & "'"
'    End If

    '2012-07-24 ���� �޸ģ�����ɸѡ������
    'ϵͳ���
    If cchkSingleNo.Value = 1 Then
        lstrWhere = lstrWhere & " and a.ϵͳ���='" & Trim(ctxtcchkNo.Text) & "'"
    End If
    '���֤��
    If cchkCardNo.Value = 1 Then
        lstrWhere = lstrWhere & " and ������ݺ���='" & ctxtcchkCardNo.Text & "'"
    End If
    '����
    If cchkName.Value = 1 Then
        lstrWhere = lstrWhere & " and ����='" & ctxtCheckName.Text & "'"
    End If
    '������λ
    If cchkWorkUnit.Value = 1 Then
        lstrWhere = lstrWhere & " and ��λ����='" & ctxtcchkWork.Text & "'"
    End If
    
    '2012-07-24 ���� �޸ģ�����ɸѡ������
    
    '2012-07-14 �ڵ�� ��
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
    sfsub������ "ְҵ�������¼��", "frmBUS_ResultInput", "ccmdQuery_Click", 6666, lstrError, False
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

'���ܣ���ѯ��Ϣ
'���ߣ�����
'ʱ�䣺2012-06-01
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
    sfsub������ "ְҵ�������¼��", "FrmENT_ResultInput", "ccmdQuery_Click", 6666, lstrError, False
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
        If lobjRec.recordcount > 0 Then
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
    SendKeys vbTab
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����沿��", "Ѫ����¼��", "ccmd��λ��λ_Click", 6666, lstrError, False
End Sub

'���ܣ���ѯ��λ��λ
'���ߣ�����
'ʱ�䣺2012-06-01
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
    'Ӧ�ðѽ������ز������(��������)
    indX = cgrdInfo.MouseRow
    indY = cgrdInfo.MouseCol
    If indX < 0 Or indY < 0 Then
        Exit Sub
    ElseIf indX >= 0 And indX <= cgrdInfo.rows And indY >= 0 And indY < cgrdInfo.cols Then
        
        
        ccmbHistory.Enabled = True
        Cmd����ģ��.Visible = True
        fraPicTool.Enabled = True
        fraPicTool.Caption = "����¼��"
        ctxtPResult.Text = ""
        
        fraResult.Enabled = True
        fraResult.Caption = "����¼��"
        ctxtResult.Text = ""
        
        ctxtBarCode.Text = cgrdInfo.TextMatrix(indX, 0)
        ctxtBarCode_KeyDown 13, 0
        '2012-07-03 �ڵ�� ��
        'ÿ�ζ��������Ϣʱ���ж��Ƿ񳬹��޸�ʱ�䡣
        '�Դ˿��Ʊ��水ť�Ƿ���á�
        If pobjҵ�����.sub�Ƿ����޸�ʱ�䷶Χ��(ctxtBarCode.Text, priDeptName, 8) = False Then
            ctlb������.Buttons(2).Enabled = False
        End If
        '2012-07-03 �ڵ�� ��
    End If
End Sub

''''���ܣ��ı�ѡ�����
''''���ߣ�����
''''ʱ�䣺2012-06-01
'''Private Sub cgrdInfoBatch_Click()
'''    cgrdInfoBatch.SelectionMode = flexSelectionByRow
'''End Sub

'���ܣ���ȡѡ�б�ŵ������Ա��Ϣ
'���ߣ�����
'ʱ�䣺2012-06-01
Private Sub cgrdInfoBatch_DblClick()
    indX = cgrdInfoBatch.MouseRow
    indY = cgrdInfoBatch.MouseCol
    If indX <= 0 Or indY < 0 Then
        Exit Sub
    ElseIf indX > 0 And indX < cgrdInfoBatch.rows And indY >= 0 And indY < cgrdInfoBatch.cols Then
        ctxt�������.Text = cgrdInfoBatch.TextMatrix(indX, 0)
        ctxt�������_KeyDown 13, 0
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

'2012-07-14 �ڵ��
Private Sub coptClasses_Click(Index As Integer)
    Dim coptIndex As Integer
    coptIndex = Index
    sub��ѯ�б���ʾ coptIndex
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
    Dim lcolְҵ������ As Object
    lstrNo = Trim(ctxtBarCode.Text)
    
    '���������Ƿ����
    Dim lobjTmp As Object
    Set lobjTmp = CreateObject("ְҵ�������¼��.clscommon")
    Set mlobjRec = lobjTmp.func��ȡ�����Ա������Ϣ(lstrNo)
    If mlobjRec.recordcount = 0 Then
        '��յ�ǰ������Ϣ
        ctxtBarCode.Enabled = True
        ctxtName.Text = ""
        ctxtSex.Text = ""
        ctxtAge.Text = ""
        ctxtCompanyName.Text = ""
        Exit Sub
    End If
    
    '�������еĸ�����Ϣ�����е������
    LoadPersonalInfo (lstrNo)
    
    '2012-04-15 �ڵ�� ��
    '����ע�͵Ĵ��룬����Ͷ�ȡ�������ͺ�λ�ö�����
    '�ʷ���������������д��
''    Set lcolְҵ������ = CreateObject("ְҵ������.clsManageMedicalExam")
''    Set rs = lcolְҵ������.func���ؿ��Һ�ͼƬ����(ctxtBarCode.Text, priDeptName)
''    If Not rs Is Nothing Then
''        ctxtResult.Text = rs("���ֽ���")
''        ctxtPResult.Text = rs("ͼƬ����")
''    End If
    
    '2012-05-22 ��¶
    '��ǰ���ҽ���
    Set lcolְҵ������ = CreateObject("ְҵ������.clsManageMedicalExam")
    ctxtResult.Text = lcolְҵ������.func���ؿ��ҽ���(ctxtBarCode.Text, priDeptName)
    '��ǰ���ҽ��(ͼƬ����)
    Set lcolְҵ������ = CreateObject("ְҵ�������¼��.clscommon")
    Set rs = lcolְҵ������.func��ȡ�����Ա�����������(ctxtBarCode.Text, priDeptName)
    If rs.recordcount > 0 And IsNull(rs("�����")) = False Then
        ctxtPResult.Text = rs("�����")
    Else
        ctxtPResult.Text = ""
    End If
    Set rs = Nothing
    '2012-05-22
    
    '2012-04-15 �ڵ�� ��
    
    'һ��ȷ����ǰ�����Ա��ţ��Ͳ��ܸ��ġ����ǣ���ս�������
    ctxtBarCode.Enabled = False
    ctxtName.Enabled = False
    ctxtSex.Enabled = False
    ctxtAge.Enabled = False
    ctxtCompanyName.Enabled = False
    
    '���ܣ��˵���ť�Ŀ���
    '���ߣ�����
    'ʱ�䣺2012-06-01
    ctlb������.Buttons(2).Enabled = True
    ctlb������.Buttons(3).Enabled = False

    ''2012-06-27 �ڵ�� ��
    'ÿ�ζ��������Ϣʱ���ж��Ƿ񳬹��޸�ʱ�䡣
    '�Դ˿��Ʊ��水ť�Ƿ���á�
    If pobjҵ�����.sub�Ƿ����޸�ʱ�䷶Χ��(ctxtBarCode.Text, priDeptName, 8) = False Then
        ctlb������.Buttons(2).Enabled = False
    End If
    '2012-06-27 �ڵ�� ��
End Sub

'2012-06-21 �ڵ��
'���ĵ�ǰ¼��״̬
Private Sub ctxtPResult_Change()
    ResultChanged = 2
End Sub

'2012-06-21 �ڵ��
'���ĵ�ǰ¼��״̬
Private Sub ctxtResult_Change()
    ResultChanged = 2
End Sub

'���ܣ��������Ų�ѯ��Ա��Ϣ
'���ߣ�����
'ʱ�䣺2012-06-01
Private Sub ctxt�������_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lstrNo As String
    Dim i As Integer
    Dim str���ҽ��� As String
    Dim lcolְҵ������ As Object
    lstrNo = Trim(ctxt�������.Text)
    Dim rs As Object
    
    '���������Ƿ����
    Dim lobjTmp As Object
    Set lobjTmp = CreateObject("ְҵ�������¼��.clscommon")
    Set mlobjRec = lobjTmp.func��ȡ�����Ա������Ϣ(lstrNo)
    If mlobjRec.recordcount = 0 Then
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
        
'        If cchk���������.Value = 0 Then
            '2012-05-22 ��¶
            '��ǰ���ҽ���
            Set lcolְҵ������ = CreateObject("ְҵ������.clsManageMedicalExam")
            ctxtResult.Text = lcolְҵ������.func���ؿ��ҽ���(ctxt�������.Text, priDeptName)
            '��ǰ���ҽ��(ͼƬ����)
            Set lcolְҵ������ = CreateObject("ְҵ�������¼��.clscommon")
            Set rs = lcolְҵ������.func��ȡ�����Ա�����������(ctxt�������.Text, priDeptName)
            If rs.recordcount > 0 And IsNull(rs("�����")) = False Then
                ctxtPResult.Text = rs("�����")
            Else
                ctxtPResult.Text = ""
            End If
            Set rs = Nothing
            '2012-05-22
'        End If

        'һ��ȷ����ǰ�����Ա��ţ��Ͳ��ܸ��ġ����ǣ���ս������ݡ�
        ctxt�������.Enabled = False
        ctxt����.Enabled = False
        ctxt�Ա�.Enabled = False
        ctxt����.Enabled = False
        ctxt��λ����.Enabled = False '��ʵ��λ�ҵ���֮������С���λ��λ����ť�����ǿ��Ըĵġ�
'''        For i = 0 To 2
'''            If coptClasses(i).Value = False Then coptClasses(i).Enabled = False
'''        Next i
        ctlb������.Buttons(2).Enabled = False
        ctlb������.Buttons(3).Enabled = True
    Else
        Set lobjTmp = Nothing
        MsgBox ("�������Աû�иÿ��ҵ������Ŀ��")
        cgrdInfoBatch.RemoveItem
        subClear
    End If
    
End Sub

Private Sub Form_Activate()
    '2012-05-24 �ڵ�� ��
    'ctxtBarCode�����ǰ�ȱ������
    ctxtBarCode.Enabled = True
    '2012-05-24 �ڵ�� ��
    ctxtBarCode.SetFocus    '�������������������
    ctxtBarCode.SelStart = Len(ctxtBarCode)
    ctxtBarCode.SelLength = 0
    cgrdInfo.SelectionMode = flexSelectionByRow
    cgrdInfo.AllowSelection = False
    
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
        .Add "��������(&D)"
        .Add "ɾ��"
        .Add "��������(&S)111"
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
    
    ctlb������.Buttons(2).Enabled = False
    ctlb������.Buttons(3).Enabled = False
    ctlb������.Buttons(4).Visible = False
    
    '�����������ȫ�ֶ���mobj���ҽʦ��
    Set mobj���ҽʦ = CreateObject("ְҵ������.clsMedicalExaminer")
    mobj���ҽʦ.��� = um�û����
    
    '�õ�ҽʦ���֣�Ϊ��ǰ�û���
    mstrDoctorName = um�û���
    LabelDoctor.Caption = LabelDoctor.Caption & " " & mstrDoctorName
    
    '����Ȩ�����á������ڸý����ϸ�����ť�������ؼ���ʹ��
    '���õĹ�����ʱ�У��鿴���޸ġ�ɾ������ӡ���������á����е���డ��
    Dim lobjTmp As Object
    Set lobjTmp = CreateObject("ְҵ������.clspermissionconfigure")
    If lobjTmp.func���Ҳ���Ȩ��(um�û����, "ְҵ�����_B��Ӱ��ƽ��¼��_�޸�") = False Then
        ctlb������.Buttons(2).Visible = False
    End If
    
    If lobjTmp.func���Ҳ���Ȩ��(um�û����, "ְҵ�����_B��Ӱ��ƽ��¼��_ɾ��") = False Then
        ctlb������.Buttons(4).Visible = False
    End If
     
    If lobjTmp.func���Ҳ���Ȩ��(um�û����, "ְҵ�����_B��Ӱ��ƽ��¼��_��������") = False Then
        ctlb������.Buttons(5).Visible = False
    End If
    
    '2012-05-22 ���� ������
    '����Ȩ������
    If lobjTmp.func���Ҳ���Ȩ��(um�û����, "ְҵ�����_B��Ӱ��ƽ��¼��_�����޸�") = False Then
        ctlb������(3).Visible = False
    End If
    '2012-05-22 ������
    Set lobjTmp = Nothing
    
    'form_load ʱ�����水ť�����趨
    cdtpConclusionDate.Value = Now
    cdtpDate.Value = Now
    DTP¼������.Value = Now
    cdtpDateBatch.Value = Now
    
    '2012-04-15 �ڵ�� ��
    'dicom�ؼ���ʼ��
    DCMList.Path = ""
    DCMList.Enabled = False
    'DCMList.ListCount = 0
    '2012-04-15 �ڵ�� ��

    '2012-06-21 �ڵ�� ��
    'ʡ������Ҫ���иı�ϵͳ��Ź���
    '��ȡϵͳ��Ź̶����֡�
    sub��ȡϵͳ��Ź̶�����
    '2012-06-21 �ڵ�� ��
    
    '2012-06-21 �ڵ�� ��
    '��ʼ����ǰ¼��״̬(��ǰ�ж�����Ȩ���޸ģ����ޣ�ֱ�Ӹ�ֵΪ3)
    ResultChanged = IIf(ResultChanged <> 3, -1, 3)
    cchkˢ����_Click
    '2012-06-21 �ڵ�� ��
    
    '2012-07-14 �ڵ�� ��
    '��ʼ����ѯ���棬������ѯ�б��ʽ����ʼ�����һ�����Ϣ��
    priDeptName = "B��Ӱ���"
    priDeptNo = "11"
    priDeptResultName = "B��Ӱ����"
    ccmdQuery_Click
    SSTPersonalInfo.Tab = 1: ccmdSelInfo_Click: SSTPersonalInfo.Tab = 0
    coptClasses_Click (0)
    '2012-07-14 �ڵ�� ��
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�������¼��", "frmBUS_ResultInput", "Form_Load", 6666, lstrError, False
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
    Case "��ս���"
        subClear
    '���ܣ���Ӳ˵��µĹ���
    '���ߣ�����
    'ʱ�䣺2012-06-01
    Case "��������"
        '2012-07-13 �ڵ�� ��
        '���û�������Ŀ����ֱ���˳��������档
        If cgrdInfoBatch.rows <= 1 Then Exit Sub
        '2012-07-13 �ڵ�� ��
        
        '2012-07-15 �ڵ�� ��
        'û��¼��������ʱ����ʾ�Ҳ����档
        If Len(Trim(ctxtResult.Text)) = 0 Then
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
        
    'ʱ�䣺2012-06-01
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
        If Len(Trim(ctxtResult.Text)) = 0 Then
            MsgBox "�㻹û��Ϊ�����½���"
            GoTo errHandler
        End If
        '2012-07-15 �ڵ�� ��
        
        Dim lstrCheck As String
        Dim lobjTmp As Object
        Dim isOk As Integer
        
        '¼����������ʱ���ܲ���
        fraResult.Enabled = False
        
        Set lcolResult = New Collection
        Set lcolItem = New Collection
        
        '2012-04-15 �ڵ�� ��
        '����ע�͵��Ĵ��룬����λ��д���ˡ�����д
''        '����B��Ӱ��������
''        If SSTPersonalInfo.Tab = 0 Then
''            lstrCheck = sub��ӵ�����(ctxtResult.Text, priDeptResultName, lstrCheck)
''        End If
''
''        '���ܣ����浥����Ŀ��ҽ������
''        '���ߣ�����
''        'ʱ�䣺2012-04-14
''        pobjҵ�����.sub������д�����ۺ�ͼƬ���� ctxtBarCode.Text, priDeptName, ctxtPResult.Text, ctxtResult.Text, um�û����
''        '���ߣ�����
''        'ʱ�䣺2012-04-14
        
        '����ͼƬ������Ҳ���������
        Call sub��ӵ�����(ctxtPResult.Text, priDeptResultName, "")
        
        '������ҽ���
        Dim lobjTmp2 As Object
        Call pobjҵ�����.sub������д������(ctxtBarCode.Text, priDeptName, ctxtResult.Text, um�û����)
        Set lobjTmp2 = Nothing
        '2012-04-15 �ڵ�� ��
        
        'lstrCheck�ַ������
        If (Not lstrCheck = "") Then
            isOk = MsgBox("������Ŀδ��д�����ȷ��������" & Chr(10) & "δ��д������¼�����ݿ⣡" & Chr(10) & Chr(10) & Trim(lstrCheck), vbOKCancel)
            If isOk = 2 Then
                Set lcolResult = Nothing
                Set lcolItem = Nothing
                Exit Sub
            End If
        End If
        
        fraResult.Enabled = True
        '2012-07-03 �ڵ�� ��
        '����һ���ֶ�"�޸���ʼʱ��"���޸ġ�ͬʱ�޸ĸÿ��ҵ������¼��״̬��
        pobjҵ�����.sub�޸���ʼʱ�� Trim(ctxtBarCode.Text), priDeptName
        pobjҵ�����.sub�޸Ľ��¼��״̬ Trim(ctxtBarCode.Text), priDeptNo, "2"  '11ΪB��Ӱ���
        pobjҵ�����.sub���¼���޸����״̬ Trim(ctxtBarCode.Text), "4"
        '2012-07-03 �ڵ�� ��
        
        subSave
        
        '2012-07-15 �ڵ�� ��
        '������֮�����½��в�ѯ��
        ccmdQuery_Click
        i = SSTPersonalInfo.Tab
        SSTPersonalInfo.Tab = 1: ccmdSelInfo_Click: SSTPersonalInfo.Tab = i
        '2012-07-15 �ڵ�� ��
        
        Set lcolResult = Nothing
        Set lcolItem = Nothing
    Case "ɾ��"
        '
    Case "��ӡ"
        '
    Case "��������"
        '
    Case "�˳�"
        '2012-06-21 �ڵ�� ��
        '�˳�ʱ�����ж��Ƿ񱣴�
        ctxtBarCode.Enabled = True
        ctxtBarCode.SetFocus
        ctxtBarCode.Enabled = False
        Dim isSave As Integer
        If ResultChanged = 2 Or ResultChanged = 0 Then
            '�޸ģ�������ڲ����鿴�����˳������ѡ������ǣ�2012-08-01��
'            If Trim(Frame6.Caption) <> "�����Ŀ�����д��" Then
'                Unload Me
'                Exit Sub
'            End If
            isSave = MsgBox("�Ƿ񱣴����޸Ľ����", vbYesNoCancel)
            If isSave = vbCancel Then Exit Sub
            If isSave = vbYes Then mobjGUI_BeforeOperate "����", False
        End If
        '2012-06-21 �ڵ�� ��
        Unload frmBUS_ResultInput
        Set frmBUS_ResultInput = Nothing
    End Select
    
    Exit Sub
errHandler:
    If Err.Number = 0 Then Exit Sub
    sfsub������ "ְҵ�������¼��", "frmBUS_ResultInput", "mobjGUI_BeforeOperate", Err.Number, Err.Description, False
    Cancel = True
    Exit Sub
    Resume
End Sub

Sub LoadPersonalInfo(ByVal paraSysNo As String)
    On Error GoTo errHandler
    Dim i As Integer
    Dim lobjTmp, lobjRec As Object
    
    Set lobjTmp = CreateObject("ְҵ�������¼��.clscommon")
    Set mlobjRec = lobjTmp.func��ȡ�����Ա������Ϣ(paraSysNo)
    If mlobjRec.recordcount > 0 Then
        ctxtName = mlobjRec("����")
        ctxtSex = mlobjRec("�Ա�")
        ctxtAge = mlobjRec("����")
        ctxtCompanyName = mlobjRec("��λ����")
        
'''        '�����������
'''        If mlobjRec("�������") = "ְҵ����" Then coptClasses(0).Value = True
'''        If mlobjRec("�������") = "���乤��" Then coptClasses(1).Value = True
'''        If mlobjRec("�������") = "��˲���" Then coptClasses(2).Value = True
        
        '��ʾ��Ƭ
        Set lobjRec = CreateObject("ְҵ������.clspersonexamed")
        lobjRec.ϵͳ��� = ctxtBarCode.Text
        Picture2.Picture = lobjRec.��Ƭ
        
        '��ʾ���˵����겡���������ǣ�2012-07-31��������������������������
            Dim lobjDatecobo As Object
            Set lobjDatecobo = mobj���ҽʦ.func��ȡ�����Ա����첡��(Trim(ctxtBarCode.Text), "B��Ӱ���")
            If Not lobjDatecobo Is Nothing Then
                Label3.Visible = True
                ccmbHistory.Visible = True
                ccmbHistory.Clear
                ccmbHistory.AddItem "����"
                For i = 1 To lobjDatecobo.recordcount
                    ccmbHistory.AddItem Format(lobjDatecobo("��дʱ��"), "yyyy-mm-dd")
'                    ccmbHistory.AddItem
                    lobjDatecobo.MoveNext
                Next i
            Else
                ccmbHistory.Clear
                ccmbHistory.Enabled = False
                
            End If
'            ccmbHistory.ListIndex = 0
            
            '��ʾ���˵����겡���������ǣ�2012-07-31�� ������������������������
        
        Set lobjRec = lobjTmp.func�Ƿ��Ѿ�����(ctxtBarCode.Text, priDeptName)
        
        If lobjRec.recordcount > 0 Then     '��û��д�������д������޸ĵı��--------------
            sub��д���е������ lobjRec
            sub�������ԱDICOMͼƬ ctxtBarCode.Text
        Else
            sub��յ�ǰ���
        End If
    Else
        MsgBox ("û�и������Ӧ�������Ա��Ϣ��")
        Exit Sub
    End If
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�������¼��", "frmBUS_ResultInput", "LoadPersonalInfo", 6666, lstrError, False
End Sub

Private Function sub��ӵ�����(ByVal paraResult As String, ByVal paraItem As String, ByVal paraCheck As String) As String
    If paraResult = "" Then
        paraCheck = paraCheck & IIf(paraCheck = "", "", Chr(10) & paraItem)
    Else
        lcolItem.Add paraItem
        lcolResult.Add paraResult
    End If
    sub��ӵ����� = paraCheck
End Function

Sub subSaveBatch(ByVal paraϵͳ��� As String)
    On Error GoTo errHandler
    
    Dim isOk As Boolean
    Dim lstrTmp As String
    Dim lobjTmp As Object
    Set lobjTmp = CreateObject("ְҵ�������¼��.clscommon")
    isOk = lobjTmp.func���浥�������(paraϵͳ���, mstrDoctorName, cdtpConclusionDate.Value, lcolItem, lcolResult, "ְҵ�����_�����Ϣ_B��Ӱ���")
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�������¼��", "frmBUS_ResultInput", "subSave", 6666, lstrError, False
End Sub

Sub subSave()
    On Error GoTo errHandler
    
    Dim isOk As Boolean
    Dim lstrTmp As String
    Dim lobjTmp As Object
    Set lobjTmp = CreateObject("ְҵ�������¼��.clscommon")
    isOk = lobjTmp.func���浥�������(ctxtBarCode.Text, mstrDoctorName, cdtpConclusionDate.Value, lcolItem, lcolResult, "ְҵ�����_�����Ϣ_B��Ӱ���")
    subClear
    If isOk = True Then MsgBox ("����ɹ���")
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�������¼��", "frmBUS_ResultInput", "subSave", 6666, lstrError, False
End Sub

Sub sub��д���е������(ByVal paraRec As Object)
    paraRec.movefirst
    If IsNull(paraRec("�����")) = True Then
        ctxtResult.Text = ""
        '2012-06-21 �ڵ�� ��
        '���õ�ǰ¼��״̬(�Ѿ�¼����������޸ĵ�ǰ���)
        ResultChanged = 0
        '2012-06-21 �ڵ�� ��
    Else
        ctxtResult.Text = paraRec("�����")
        '2012-06-21 �ڵ�� ��
        '���õ�ǰ¼��״̬(�Ѿ�¼����������޸ĵ�ǰ���)
        ResultChanged = 1
        '2012-06-21 �ڵ�� ��
    End If
End Sub

Sub sub��յ�ǰ���()
    ctxtResult.Text = ""
    
    '��յ�ǰͼƬ���
    '-------------���������ޡ�����--------------
    
    '-------------���������ޡ�����--------------
End Sub
Sub sub�������ԱDICOMͼƬ(ByVal paraSysNo As String)
    '-------------���������ޡ�����--------------
    
    '-------------���������ޡ�����--------------
End Sub

'2012-04-14 �ڵ��
'��ĳ��dicomͼƬ����¼·���͵�ǰ�ļ������ļ���·��
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
    MsgBox ("�ļ���ȡ�������Ժ����ԡ�")
End Sub

'2012-04-14 �ڵ��
'���浱ǰdicomͼƬ��Ŀǰ���淽��Ϊ�����ļ���ǰ����ϵ�ǰ�޸����ں��ټ���.dcm��׺
'ע�⣺�����ļ�Ĭ��Ϊ��ǰͼƬ�ļ�����Ŀ¼�£�����ֻ�������ļ������ɡ�������'/'��'\'�������ַ���
Private Sub ccmdSavePic_Click()
    If cchkReplace.Value = 1 Then
        Dicm.ImageSaveToDICOM = DCMList.List(DCMIdx)   '�滻ԭ���ļ�(���Ƽ�)
    Else
        Dicm.ImageSaveToDICOM = Replace(DCMList.List(DCMIdx), ".dcm", "") & "_" & Format(Date, "yyyymmdd") & ".dcm"
    End If
End Sub

'2012-04-14 �ڵ��
'�����ļ��б����µ�ǰ��ʾͼƬ
Private Sub DCMList_Click()
    '������ʾ��dicm�ؼ��У��޷��ڴ����п��ơ�������ʡ�Դ�����
    'ͬʱ������ԭ��Ϊԭͼ�����ݸ�ʽ����
    DCMIdx = DCMList.ListIndex
    DCMPath = DCMDir & "\" & DCMList.List(DCMIdx)
    '����������ָ���enable���ƣ���������ʧЧ
    'DCMList.Enabled = False
    'MousePointer = 11
    Dicm.OpenFile (DCMPath)
    llabCurr = "��" & (DCMIdx + 1) & "/��" & DCMList.ListCount
'    Timer1_Timer
    'DCMList.Enabled = True
    'MousePointer = 1
End Sub

'2012-04-14 �ڵ��
'��ʼ��dicom�ļ��б�
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

'���ܣ����������ȡ������Ϣ
'���ߣ�����
'ʱ�䣺2012-06-01
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
        
        '�������еĸ�����Ϣ�����е������
        '��ʾ��Ƭ
        Set lobjRec = CreateObject("ְҵ������.clspersonexamed")
        lobjRec.ϵͳ��� = ctxt�������.Text
        Picture4.Enabled = True
        Picture4.Visible = True
        Picture4.Picture = lobjRec.��Ƭ
            
        Set lobjRec = lobjTmp.func�Ƿ��Ѿ�����(ctxt�������.Text, priDeptName)
        If lobjRec.recordcount = 0 Then
            If ResultChanged <> 3 Then ResultChanged = 0
        ElseIf lobjRec.recordcount > 0 Then
            If ResultChanged <> 3 Then ResultChanged = 1
        End If
    Else
        MsgBox ("û�и������Ӧ�������Ա��Ϣ��")
        Exit Sub
    End If
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�������¼��", "frmENT_ResultInput", "LoadPersonalInfo", 6666, lstrError, False
End Sub

'���ܣ�����������
'���ߣ�����
'ʱ�䣺2012-06-01
Sub sub��������()

    MousePointer = 11
    Dim lblnNotOver As Boolean
    Dim i As Integer
    Dim barCode As Collection '���������������
        'cstbMain.Panels(1) = "���ڱ��棬���Ժ�..."
        
        '��ʱ���治�ܲ�����
        Frame1.Enabled = False
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
        
        If cgrdInfoBatch.rows < 1 Then
        MsgBox ("��ȷ��¼����Ա��Ŀ�Ƿ���ȷ��")
        Exit Sub
    End If
    Dim ccrpValue As Integer
    Dim ccrpI As Integer
    Dim isOk As Boolean
    Dim lstrTmp As String
    Dim lobjTmp As Object
    
    '��ʾ���������
    ccrpI = barCode.Count
    ccrp����.Max = ccrpI
    ccrp����.Visible = True
    ccrp����.Caption = "0%"
    ccrp����.Value = 0
    For i = 1 To barCode.Count
        '����ͼƬ������Ҳ���������
        Call sub��ӵ�����(ctxtPResult.Text, priDeptResultName, "")
            
        '������ҽ���
        Dim lobjTmp2 As Object
        Call pobjҵ�����.sub������д������(barCode(i), priDeptName, ctxtResult.Text, um�û����)
        Set lobjTmp2 = Nothing
        '2012-04-15 �ڵ�� ��
        
        ccrp����.Caption = Int(i / ccrp����.Max * 100) + ccrpValue & "%"
        ccrp����.Value = ccrp����.Value + 1
        
        '2012-07-03 �ڵ�� ��
        '����һ���ֶ�"�޸���ʼʱ��"���޸ġ�ͬʱ�޸ĸÿ��ҵ������¼��״̬��
        pobjҵ�����.sub�޸���ʼʱ�� barCode(i), priDeptName
        pobjҵ�����.sub�޸Ľ��¼��״̬ barCode(i), priDeptNo, "2"
        pobjҵ�����.sub���¼���޸����״̬ barCode(i), "4"
        '2012-07-03 �ڵ�� ��
        
        subSaveBatch barCode(i)
        
    Next i
    MsgBox ("��������ɹ���")
    subClear
    
    ccrp����.Visible = False

    MousePointer = 0
    
    Exit Sub
errHandler:
    Dim lstrError As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�������¼��", "FrmENT_ResultInput", "ccmdSave_Click", 6666, lstrError, False

End Sub

'���ܣ���ս��湦��
'���ߣ�����
'ʱ�䣺2012-06-01
Sub subClear()
    TotalPeople.Caption = 0
    TotalPeopleBatch.Caption = 0
    
    '��յ�ǰ������Ϣ
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
    ctxtResult.Text = ""
    ctxtPResult.Text = ""
    
    cchkDateBatch.Value = 0
    cchkCompanyBatch.Value = 0
    TotalPeopleBatch.Caption = "0"
    
'�����Ƭ
    Set Picture2.Picture = Nothing
    Set Picture4.Picture = Nothing
    
    '�ָ�Ϊform_loadʱ��״̬��
    If SSTPersonalInfo.Tab = 0 Then
        ctxtBarCode.Enabled = True
        ctxtBarCode.SetFocus
        ctxtName.Enabled = True
        ctxtSex.Enabled = True
        ctxtAge.Enabled = True
        ctxtCompanyName.Enabled = True
    Else
        ctxt�������.Enabled = True
        ctxt����.Enabled = True
        ctxt�Ա�.Enabled = True
        ctxt����.Enabled = True
        ctxt��λ����.Enabled = True
    End If
    
    sub��յ�ǰ���
    
    '2012-04-15 �ڵ�� ��
    '����dicomͼ���ļ��б��
    DCMList.Enabled = False
    
'''    coptClasses(0).Enabled = True
'''    coptClasses(1).Enabled = True
'''    coptClasses(2).Enabled = True
    ctlb������.Enabled = True
    SSTPersonalInfo.Enabled = True
    Frame1.Enabled = True
'''    coptClasses(0).Value = 1
    ctlb������.Buttons(1).Enabled = True
    ctlb������.Buttons(2).Enabled = False
    ctlb������.Buttons(3).Enabled = False
    
    '2012-06-21 �ڵ�� ��
    '��ʼ����ǰ¼��״̬(��ǰ�ж�����Ȩ���޸ģ����ޣ�ֱ�Ӹ�ֵΪ3)
    ResultChanged = IIf(ResultChanged <> 3, -1, 3)
    cchkˢ����_Click
    '2012-06-21 �ڵ�� ��
    
End Sub

'2012-06-21 �ڵ��
Sub sub��ȡϵͳ��Ź̶�����()
    '��ȡ����������
    Dim lobjRec As Object
    Set lobjRec = dafuncGetData("select getdate()")
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
