VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmCareerHstRegt 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "�ܼ��߸�����Ϣ¼��"
   ClientHeight    =   9315
   ClientLeft      =   4830
   ClientTop       =   2430
   ClientWidth     =   14055
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9315
   ScaleWidth      =   14055
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   5895
      Left            =   120
      TabIndex        =   151
      TabStop         =   0   'False
      Top             =   2880
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   10398
      _Version        =   393216
      Tabs            =   6
      Tab             =   4
      TabsPerRow      =   6
      TabHeight       =   520
      BackColor       =   0
      ForeColor       =   16711680
      TabCaption(0)   =   "��������ʷ"
      TabPicture(0)   =   "frmCareerHstRegister.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "freRadiation"
      Tab(0).Control(1)=   "freOrdinary"
      Tab(0).Control(2)=   "freNuclear"
      Tab(0).Control(3)=   "Frame6"
      Tab(0).Control(4)=   "Frame1"
      Tab(0).Control(5)=   "ctxtOther"
      Tab(0).Control(6)=   "Label5(0)"
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "ְҵʷ"
      TabPicture(1)   =   "frmCareerHstRegister.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cgrdְҵʷ"
      Tab(1).Control(1)=   "Frame10"
      Tab(1).Control(2)=   "Ccmd�޸�"
      Tab(1).Control(3)=   "Ccmdɾ��"
      Tab(1).Control(4)=   "Frame11"
      Tab(1).Control(5)=   "Command4"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "������ʷ(����ְҵ��ʷ)"
      TabPicture(2)   =   "frmCareerHstRegister.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame13"
      Tab(2).Control(1)=   "Cmmdmody��ʷ"
      Tab(2).Control(2)=   "Ccmdcancel��ʷ"
      Tab(2).Control(3)=   "cgrd��ʷ"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "�Ծ�֢״"
      TabPicture(3)   =   "frmCareerHstRegister.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame15"
      Tab(3).Control(1)=   "cgrd֢״"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "���һ�����"
      TabPicture(4)   =   "frmCareerHstRegister.frx":0070
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Frame7"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Frame5"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Text��ʱ"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "��״ѯ��"
      TabPicture(5)   =   "frmCareerHstRegister.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "cgrdzzxw"
      Tab(5).Control(1)=   "Command1"
      Tab(5).Control(2)=   "Command2"
      Tab(5).ControlCount=   3
      Begin VB.TextBox Text��ʱ 
         Height          =   1335
         Left            =   8400
         TabIndex        =   321
         Text            =   "����������سӶ�ȡ���ַ���"
         Top             =   1800
         Width           =   2775
      End
      Begin VB.Frame freRadiation 
         BorderStyle     =   0  'None
         Height          =   3735
         Left            =   -74760
         TabIndex        =   154
         Top             =   600
         Width           =   11895
         Begin VB.Frame Frame19 
            Caption         =   "�̾�ʷ"
            ForeColor       =   &H000080FF&
            Height          =   2175
            Index           =   0
            Left            =   6000
            TabIndex        =   187
            Top             =   1560
            Width           =   5175
            Begin VB.TextBox ctxtMore 
               Height          =   375
               Index           =   0
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   93
               Top             =   1680
               Width           =   4695
            End
            Begin VB.TextBox ctxt���� 
               Height          =   270
               Index           =   0
               Left            =   960
               TabIndex        =   90
               Top             =   1200
               Width           =   975
            End
            Begin VB.TextBox ctxt���� 
               Height          =   270
               Index           =   0
               Left            =   960
               TabIndex        =   88
               Top             =   840
               Width           =   975
            End
            Begin VB.TextBox ctxt���� 
               Height          =   270
               Index           =   0
               Left            =   3360
               TabIndex        =   89
               Top             =   960
               Width           =   975
            End
            Begin VB.TextBox ctxt������ 
               Height          =   270
               Index           =   0
               Left            =   960
               TabIndex        =   92
               Top             =   480
               Width           =   975
            End
            Begin VB.TextBox ctxt������ 
               Height          =   270
               Index           =   0
               Left            =   3360
               TabIndex        =   91
               Top             =   600
               Width           =   975
            End
            Begin VB.ComboBox ccmb���� 
               Height          =   300
               Index           =   0
               Left            =   3360
               TabIndex        =   87
               Top             =   120
               Width           =   1335
            End
            Begin VB.ComboBox ccmb���� 
               Height          =   300
               Index           =   0
               Left            =   960
               TabIndex        =   86
               Top             =   120
               Width           =   1335
            End
            Begin VB.Label Label101 
               AutoSize        =   -1  'True
               Caption         =   "ML/��"
               Height          =   300
               Index           =   0
               Left            =   4440
               TabIndex        =   190
               Top             =   600
               Width           =   810
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               Caption         =   "�����ס��������ʳϰ�ߡ��̾��Ⱥ�������"
               Height          =   180
               Index           =   1
               Left            =   120
               TabIndex        =   241
               Top             =   1440
               Width           =   3420
            End
            Begin VB.Label Label110 
               AutoSize        =   -1  'True
               Caption         =   "��"
               Height          =   180
               Index           =   0
               Left            =   2040
               TabIndex        =   199
               Top             =   1200
               Width           =   180
            End
            Begin VB.Label Label109 
               AutoSize        =   -1  'True
               Caption         =   "��"
               Height          =   180
               Index           =   0
               Left            =   4440
               TabIndex        =   198
               Top             =   960
               Width           =   180
            End
            Begin VB.Label Label108 
               AutoSize        =   -1  'True
               Caption         =   "��"
               Height          =   180
               Index           =   0
               Left            =   2040
               TabIndex        =   197
               Top             =   840
               Width           =   180
            End
            Begin VB.Label Label107 
               AutoSize        =   -1  'True
               Caption         =   "����ʱ����"
               Height          =   180
               Index           =   0
               Left            =   120
               TabIndex        =   196
               Top             =   1200
               Width           =   900
            End
            Begin VB.Label Label106 
               AutoSize        =   -1  'True
               Caption         =   "���䣺"
               Height          =   180
               Index           =   0
               Left            =   240
               TabIndex        =   195
               Top             =   840
               Width           =   540
            End
            Begin VB.Label Label105 
               AutoSize        =   -1  'True
               Caption         =   "���䣺"
               Height          =   180
               Index           =   0
               Left            =   2640
               TabIndex        =   194
               Top             =   960
               Width           =   540
            End
            Begin VB.Label Label104 
               AutoSize        =   -1  'True
               Caption         =   "��������"
               Height          =   180
               Index           =   0
               Left            =   120
               TabIndex        =   193
               Top             =   480
               Width           =   720
            End
            Begin VB.Label Label103 
               AutoSize        =   -1  'True
               Caption         =   "��������"
               Height          =   180
               Index           =   0
               Left            =   2640
               TabIndex        =   192
               Top             =   600
               Width           =   720
            End
            Begin VB.Label Label102 
               AutoSize        =   -1  'True
               Caption         =   "֧/��"
               Height          =   180
               Index           =   0
               Left            =   2040
               TabIndex        =   191
               Top             =   480
               Width           =   450
            End
            Begin VB.Label Label100 
               AutoSize        =   -1  'True
               Caption         =   "���Ƴ̶ȣ�"
               Height          =   180
               Index           =   0
               Left            =   2520
               TabIndex        =   189
               Top             =   195
               Width           =   900
            End
            Begin VB.Label Label99 
               AutoSize        =   -1  'True
               Caption         =   "���̶̳ȣ�"
               Height          =   180
               Index           =   0
               Left            =   120
               TabIndex        =   188
               Top             =   195
               Width           =   900
            End
         End
         Begin VB.ComboBox Combo11 
            Height          =   300
            ItemData        =   "frmCareerHstRegister.frx":00A8
            Left            =   5040
            List            =   "frmCareerHstRegister.frx":00B5
            TabIndex        =   275
            Text            =   "����"
            Top             =   360
            Width           =   855
         End
         Begin VB.Frame Frame2 
            Caption         =   "����ʷ"
            ForeColor       =   &H000080FF&
            Height          =   1095
            Index           =   0
            Left            =   120
            TabIndex        =   167
            Top             =   120
            Width           =   11055
            Begin VB.TextBox ctxtmarrydate 
               Height          =   270
               Index           =   0
               Left            =   960
               TabIndex        =   276
               Text            =   "  �� ��"
               Top             =   720
               Width           =   1575
            End
            Begin VB.TextBox ctxtmatehelh 
               Height          =   270
               Index           =   0
               Left            =   3960
               TabIndex        =   71
               Text            =   "����"
               Top             =   240
               Width           =   855
            End
            Begin VB.TextBox ctxtmateradioac 
               Height          =   495
               Index           =   0
               Left            =   5880
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   74
               Text            =   "frmCareerHstRegister.frx":00CB
               Top             =   480
               Width           =   4815
            End
            Begin VB.TextBox ctxtmatejob 
               Height          =   270
               Index           =   0
               Left            =   3960
               TabIndex        =   73
               Top             =   720
               Width           =   1695
            End
            Begin VB.ComboBox Ccmb��� 
               Height          =   300
               Index           =   0
               Left            =   960
               TabIndex        =   70
               Top             =   240
               Width           =   1695
            End
            Begin MSComCtl2.DTPicker ctxtmarrydate1 
               CausesValidation=   0   'False
               Height          =   300
               Index           =   0
               Left            =   7560
               TabIndex        =   72
               Top             =   120
               Visible         =   0   'False
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   0   'False
               CustomFormat    =   "yyyy/MM"
               Format          =   60030976
               CurrentDate     =   41013
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "��ż����״����"
               Height          =   180
               Index           =   0
               Left            =   2760
               TabIndex        =   172
               Top             =   300
               Width           =   1260
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "��ż�Ӵ������������"
               Height          =   180
               Index           =   0
               Left            =   5880
               TabIndex        =   171
               Top             =   240
               Width           =   1800
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "��żְҵ��"
               Height          =   180
               Index           =   0
               Left            =   2760
               TabIndex        =   170
               Top             =   760
               Width           =   900
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "������ڣ�"
               Height          =   180
               Index           =   0
               Left            =   120
               TabIndex        =   169
               Top             =   760
               Width           =   900
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "�Ƿ��飺"
               Height          =   180
               Index           =   0
               Left            =   120
               TabIndex        =   168
               Top             =   300
               Width           =   900
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "����ʷ(����ż����ʷ)"
            ForeColor       =   &H000080FF&
            Height          =   2295
            Index           =   0
            Left            =   120
            TabIndex        =   155
            Top             =   1440
            Width           =   5775
            Begin VB.TextBox ctxtŮ���������� 
               Height          =   270
               Index           =   0
               Left            =   2880
               TabIndex        =   285
               Top             =   1920
               Width           =   855
            End
            Begin VB.TextBox ctxt����Ů�� 
               Height          =   270
               Index           =   0
               Left            =   1200
               TabIndex        =   283
               Text            =   "0"
               Top             =   1920
               Width           =   495
            End
            Begin VB.TextBox ctxt�к��������� 
               Height          =   270
               Index           =   0
               Left            =   2880
               TabIndex        =   281
               Top             =   1560
               Width           =   855
            End
            Begin VB.ComboBox Combo8 
               Height          =   300
               ItemData        =   "frmCareerHstRegister.frx":00D0
               Left            =   4800
               List            =   "frmCareerHstRegister.frx":00DD
               TabIndex        =   269
               Text            =   "����"
               Top             =   1800
               Width           =   855
            End
            Begin VB.ComboBox Combo7 
               Height          =   300
               ItemData        =   "frmCareerHstRegister.frx":00F3
               Left            =   3960
               List            =   "frmCareerHstRegister.frx":0103
               TabIndex        =   268
               Text            =   "���в���ԭ��ģ��"
               Top             =   1080
               Width           =   1815
            End
            Begin VB.TextBox ctxt���в��� 
               Height          =   375
               Index           =   0
               Left            =   2040
               MultiLine       =   -1  'True
               TabIndex        =   85
               Text            =   "frmCareerHstRegister.frx":0135
               Top             =   1080
               Width           =   1815
            End
            Begin VB.TextBox ctxt���� 
               Height          =   270
               Index           =   0
               Left            =   2880
               TabIndex        =   76
               Text            =   "0"
               Top             =   480
               Width           =   735
            End
            Begin VB.TextBox ctxt��� 
               Height          =   270
               Index           =   0
               Left            =   1920
               TabIndex        =   75
               Text            =   "0"
               Top             =   480
               Width           =   735
            End
            Begin VB.TextBox ctxt��� 
               Height          =   270
               Index           =   0
               Left            =   1080
               TabIndex        =   79
               Text            =   "0"
               Top             =   480
               Width           =   735
            End
            Begin VB.TextBox ctxt��λ���� 
               Height          =   270
               Index           =   0
               Left            =   1080
               TabIndex        =   83
               Text            =   "0"
               Top             =   1080
               Width           =   855
            End
            Begin VB.TextBox ctxt�д� 
               Height          =   270
               Index           =   0
               Left            =   240
               TabIndex        =   77
               Text            =   "0"
               Top             =   480
               Width           =   735
            End
            Begin VB.TextBox ctxt��̥ 
               Height          =   270
               Index           =   0
               Left            =   240
               TabIndex        =   78
               Text            =   "0"
               Top             =   1080
               Width           =   735
            End
            Begin VB.TextBox ctxt��̥ 
               Height          =   270
               Index           =   0
               Left            =   4800
               TabIndex        =   84
               Text            =   "0"
               Top             =   480
               Width           =   735
            End
            Begin VB.TextBox ctxt���� 
               Height          =   270
               Index           =   0
               Left            =   3720
               TabIndex        =   80
               Text            =   "0"
               Top             =   480
               Width           =   855
            End
            Begin VB.TextBox ctxt������Ů 
               Height          =   270
               Index           =   0
               Left            =   1200
               TabIndex        =   81
               Text            =   "0"
               Top             =   1560
               Width           =   495
            End
            Begin VB.TextBox ctxt��Ů���� 
               Height          =   270
               Index           =   0
               Left            =   3960
               TabIndex        =   82
               Text            =   "����"
               Top             =   1800
               Width           =   855
            End
            Begin VB.Label Label73 
               Caption         =   "�������ڣ�"
               Height          =   255
               Left            =   1920
               TabIndex        =   284
               Top             =   1920
               Width           =   975
            End
            Begin VB.Label Label72 
               Caption         =   "����Ů����"
               Height          =   255
               Left            =   240
               TabIndex        =   282
               Top             =   1920
               Width           =   975
            End
            Begin VB.Label Label71 
               Caption         =   "�������ڣ�"
               Height          =   255
               Left            =   1920
               TabIndex        =   280
               Top             =   1560
               Width           =   975
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               Caption         =   "���в���ԭ��"
               Height          =   180
               Index           =   0
               Left            =   2040
               TabIndex        =   166
               Top             =   840
               Width           =   1500
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "������"
               Height          =   180
               Index           =   0
               Left            =   2880
               TabIndex        =   165
               Top             =   240
               Width           =   540
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "�����"
               Height          =   180
               Index           =   0
               Left            =   1920
               TabIndex        =   164
               Top             =   240
               Width           =   540
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "�����"
               Height          =   180
               Index           =   0
               Left            =   1080
               TabIndex        =   163
               Top             =   240
               Width           =   540
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               Caption         =   "��λ���"
               Height          =   180
               Index           =   0
               Left            =   1080
               TabIndex        =   162
               Top             =   840
               Width           =   900
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "�дΣ�"
               Height          =   180
               Index           =   0
               Left            =   240
               TabIndex        =   161
               Top             =   240
               Width           =   540
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               Caption         =   "��̥��"
               Height          =   180
               Index           =   0
               Left            =   240
               TabIndex        =   160
               Top             =   840
               Width           =   540
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               Caption         =   "��̥��"
               Height          =   180
               Index           =   0
               Left            =   4800
               TabIndex        =   159
               Top             =   240
               Width           =   540
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               Caption         =   "������"
               Height          =   180
               Index           =   0
               Left            =   3720
               TabIndex        =   158
               Top             =   240
               Width           =   540
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               Caption         =   "�����к���"
               Height          =   180
               Index           =   0
               Left            =   240
               TabIndex        =   157
               Top             =   1560
               Width           =   900
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               Caption         =   "��Ů����״����"
               Height          =   180
               Index           =   0
               Left            =   3960
               TabIndex        =   156
               Top             =   1560
               Width           =   1260
            End
         End
      End
      Begin VB.Frame freOrdinary 
         BorderStyle     =   0  'None
         Height          =   3735
         Left            =   -74640
         TabIndex        =   207
         Top             =   600
         Width           =   11175
         Begin VB.TextBox ctxt������ 
            Height          =   375
            Index           =   2
            Left            =   6840
            TabIndex        =   307
            Top             =   1800
            Width           =   4215
         End
         Begin VB.Frame Frame3 
            Caption         =   "����ʷ(����ż����ʷ)"
            ForeColor       =   &H000080FF&
            Height          =   2655
            Index           =   2
            Left            =   120
            TabIndex        =   224
            Top             =   960
            Width           =   5775
            Begin VB.TextBox ctxt�쳣̥ 
               Height          =   270
               Left            =   1680
               TabIndex        =   61
               Text            =   "0"
               Top             =   1080
               Width           =   1215
            End
            Begin VB.TextBox ctxt������Ů 
               Height          =   270
               Index           =   2
               Left            =   1680
               TabIndex        =   57
               Text            =   "0"
               Top             =   360
               Width           =   1215
            End
            Begin VB.TextBox ctxt���� 
               Height          =   270
               Index           =   2
               Left            =   1680
               TabIndex        =   59
               Text            =   "0"
               Top             =   720
               Width           =   1215
            End
            Begin VB.TextBox ctxt��� 
               Height          =   270
               Index           =   2
               Left            =   4200
               TabIndex        =   58
               Text            =   "0"
               Top             =   360
               Width           =   975
            End
            Begin VB.TextBox ctxt���� 
               Height          =   270
               Index           =   2
               Left            =   4200
               TabIndex        =   60
               Text            =   "0"
               Top             =   720
               Width           =   975
            End
            Begin VB.Label Label�쳣̥ 
               AutoSize        =   -1  'True
               Caption         =   "�쳣̥��"
               Height          =   180
               Left            =   840
               TabIndex        =   229
               Top             =   1080
               Width           =   720
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               Caption         =   "������Ů��Ŀ��"
               Height          =   180
               Index           =   1
               Left            =   480
               TabIndex        =   228
               Top             =   360
               Width           =   1260
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               Caption         =   "������"
               Height          =   180
               Index           =   1
               Left            =   1080
               TabIndex        =   227
               Top             =   720
               Width           =   540
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "�����"
               Height          =   180
               Index           =   1
               Left            =   3480
               TabIndex        =   226
               Top             =   360
               Width           =   540
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "������"
               Height          =   180
               Index           =   1
               Left            =   3480
               TabIndex        =   225
               Top             =   720
               Width           =   540
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "����ʷ"
            ForeColor       =   &H000080FF&
            Height          =   615
            Index           =   2
            Left            =   120
            TabIndex        =   222
            Top             =   240
            Width           =   5775
            Begin VB.ComboBox Ccmb��� 
               Height          =   300
               Index           =   2
               Left            =   1680
               TabIndex        =   51
               Top             =   240
               Width           =   1695
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "�Ƿ��飺"
               Height          =   180
               Index           =   2
               Left            =   480
               TabIndex        =   223
               Top             =   300
               Width           =   900
            End
         End
         Begin VB.Frame Frame19 
            Caption         =   "�̾�ʷ"
            ForeColor       =   &H000080FF&
            Height          =   1455
            Index           =   1
            Left            =   6000
            TabIndex        =   209
            Top             =   240
            Width           =   5055
            Begin VB.ComboBox ccmb���� 
               Height          =   300
               Index           =   2
               Left            =   960
               TabIndex        =   62
               Top             =   120
               Width           =   1335
            End
            Begin VB.ComboBox ccmb���� 
               Height          =   300
               Index           =   2
               Left            =   3360
               TabIndex        =   63
               Top             =   120
               Width           =   1335
            End
            Begin VB.TextBox ctxt������ 
               Height          =   270
               Index           =   2
               Left            =   3360
               TabIndex        =   68
               Top             =   600
               Width           =   975
            End
            Begin VB.TextBox ctxt������ 
               Height          =   270
               Index           =   2
               Left            =   960
               TabIndex        =   67
               Top             =   480
               Width           =   975
            End
            Begin VB.TextBox ctxt���� 
               Height          =   270
               Index           =   2
               Left            =   3360
               TabIndex        =   65
               Top             =   960
               Width           =   975
            End
            Begin VB.TextBox ctxt���� 
               Height          =   270
               Index           =   2
               Left            =   960
               TabIndex        =   64
               Top             =   840
               Width           =   975
            End
            Begin VB.TextBox ctxt���� 
               Height          =   270
               Index           =   2
               Left            =   960
               TabIndex        =   66
               Top             =   1200
               Width           =   975
            End
            Begin VB.Label Label99 
               AutoSize        =   -1  'True
               Caption         =   "���̶̳ȣ�"
               Height          =   180
               Index           =   1
               Left            =   120
               TabIndex        =   221
               Top             =   195
               Width           =   900
            End
            Begin VB.Label Label100 
               AutoSize        =   -1  'True
               Caption         =   "���Ƴ̶ȣ�"
               Height          =   180
               Index           =   1
               Left            =   2520
               TabIndex        =   220
               Top             =   195
               Width           =   900
            End
            Begin VB.Label Label101 
               AutoSize        =   -1  'True
               Caption         =   "ML/��"
               Height          =   180
               Index           =   1
               Left            =   4440
               TabIndex        =   219
               Top             =   600
               Width           =   450
            End
            Begin VB.Label Label102 
               AutoSize        =   -1  'True
               Caption         =   "֧/��"
               Height          =   180
               Index           =   1
               Left            =   2040
               TabIndex        =   218
               Top             =   480
               Width           =   450
            End
            Begin VB.Label Label103 
               AutoSize        =   -1  'True
               Caption         =   "��������"
               Height          =   180
               Index           =   1
               Left            =   2640
               TabIndex        =   217
               Top             =   600
               Width           =   720
            End
            Begin VB.Label Label104 
               AutoSize        =   -1  'True
               Caption         =   "��������"
               Height          =   180
               Index           =   1
               Left            =   120
               TabIndex        =   216
               Top             =   480
               Width           =   720
            End
            Begin VB.Label Label105 
               AutoSize        =   -1  'True
               Caption         =   "���䣺"
               Height          =   180
               Index           =   1
               Left            =   2640
               TabIndex        =   215
               Top             =   960
               Width           =   540
            End
            Begin VB.Label Label106 
               AutoSize        =   -1  'True
               Caption         =   "���䣺"
               Height          =   180
               Index           =   1
               Left            =   120
               TabIndex        =   214
               Top             =   840
               Width           =   540
            End
            Begin VB.Label Label107 
               AutoSize        =   -1  'True
               Caption         =   "����ʱ����"
               Height          =   180
               Index           =   1
               Left            =   120
               TabIndex        =   213
               Top             =   1200
               Width           =   900
            End
            Begin VB.Label Label108 
               AutoSize        =   -1  'True
               Caption         =   "��"
               Height          =   180
               Index           =   1
               Left            =   2040
               TabIndex        =   212
               Top             =   840
               Width           =   180
            End
            Begin VB.Label Label109 
               AutoSize        =   -1  'True
               Caption         =   "��"
               Height          =   180
               Index           =   1
               Left            =   4440
               TabIndex        =   211
               Top             =   960
               Width           =   180
            End
            Begin VB.Label Label110 
               AutoSize        =   -1  'True
               Caption         =   "��"
               Height          =   180
               Index           =   1
               Left            =   2040
               TabIndex        =   210
               Top             =   1200
               Width           =   180
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "����ʷ"
            ForeColor       =   &H000080FF&
            Height          =   1215
            Index           =   1
            Left            =   6000
            TabIndex        =   208
            Top             =   2280
            Width           =   5055
            Begin VB.ComboBox Combo2 
               Height          =   300
               Left            =   3240
               TabIndex        =   252
               Text            =   "Combo2"
               Top             =   600
               Width           =   1575
            End
            Begin VB.TextBox ctxt����ʷ 
               Height          =   735
               Index           =   2
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   69
               Top             =   240
               Width           =   3135
            End
            Begin VB.Label Label63 
               Caption         =   "��ѡ�����ҩԴ:"
               Height          =   255
               Left            =   3240
               TabIndex        =   253
               Top             =   240
               Width           =   1455
            End
         End
         Begin VB.Label Lab������ 
            Caption         =   "�����أ�"
            Height          =   255
            Left            =   6120
            TabIndex        =   306
            Top             =   1920
            Width           =   735
         End
      End
      Begin VB.Frame freNuclear 
         BorderStyle     =   0  'None
         Height          =   3615
         Left            =   -74760
         TabIndex        =   173
         Top             =   600
         Width           =   10815
         Begin VB.Frame Frame2 
            Caption         =   "����ʷ"
            ForeColor       =   &H000080FF&
            Height          =   1095
            Index           =   1
            Left            =   120
            TabIndex        =   201
            Top             =   120
            Width           =   10935
            Begin VB.TextBox ctxtmarrydate 
               Height          =   270
               Index           =   1
               Left            =   960
               TabIndex        =   286
               Text            =   "����"
               Top             =   720
               Width           =   1695
            End
            Begin VB.ComboBox Ccmb��� 
               Height          =   300
               Index           =   1
               Left            =   960
               TabIndex        =   40
               Top             =   240
               Width           =   1695
            End
            Begin VB.TextBox ctxtmatejob 
               Height          =   270
               Index           =   1
               Left            =   3960
               TabIndex        =   43
               Top             =   720
               Width           =   1695
            End
            Begin VB.TextBox ctxtmateradioac 
               Height          =   495
               Index           =   1
               Left            =   5880
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   44
               Top             =   480
               Width           =   4815
            End
            Begin VB.TextBox ctxtmatehelh 
               Height          =   270
               Index           =   1
               Left            =   3960
               TabIndex        =   41
               Text            =   "����"
               Top             =   240
               Width           =   1695
            End
            Begin MSComCtl2.DTPicker ctxtmarrydate2 
               Height          =   300
               Index           =   1
               Left            =   7680
               TabIndex        =   42
               Top             =   120
               Visible         =   0   'False
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   529
               _Version        =   393216
               CustomFormat    =   "yyyy/MM/dd"
               Format          =   60030976
               CurrentDate     =   41013
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "�Ƿ��飺"
               Height          =   180
               Index           =   1
               Left            =   120
               TabIndex        =   206
               Top             =   300
               Width           =   900
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "������ڣ�"
               Height          =   180
               Index           =   1
               Left            =   120
               TabIndex        =   205
               Top             =   760
               Width           =   900
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "��żְҵ��"
               Height          =   180
               Index           =   1
               Left            =   2760
               TabIndex        =   204
               Top             =   760
               Width           =   900
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "��ż�Ӵ������������"
               Height          =   180
               Index           =   1
               Left            =   5880
               TabIndex        =   203
               Top             =   240
               Width           =   1800
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "��ż����״����"
               Height          =   180
               Index           =   1
               Left            =   2760
               TabIndex        =   202
               Top             =   300
               Width           =   1260
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "����ʷ(����ż����ʷ)"
            ForeColor       =   &H000080FF&
            Height          =   2295
            Index           =   1
            Left            =   120
            TabIndex        =   39
            Top             =   1320
            Width           =   5775
            Begin VB.TextBox ctxt���� 
               Height          =   270
               Index           =   1
               Left            =   120
               TabIndex        =   305
               Top             =   1200
               Width           =   735
            End
            Begin VB.TextBox ctxtŮ���������� 
               Height          =   270
               Index           =   1
               Left            =   2280
               TabIndex        =   297
               Top             =   2040
               Width           =   1215
            End
            Begin VB.TextBox ctxt�к��������� 
               Height          =   270
               Index           =   1
               Left            =   2280
               TabIndex        =   296
               Top             =   1680
               Width           =   1215
            End
            Begin VB.TextBox ctxt����Ů�� 
               Height          =   270
               Index           =   1
               Left            =   960
               TabIndex        =   293
               Top             =   2040
               Width           =   495
            End
            Begin VB.TextBox ctxt������Ů 
               Height          =   270
               Index           =   1
               Left            =   960
               TabIndex        =   292
               Top             =   1680
               Width           =   495
            End
            Begin VB.ComboBox Combo10 
               Height          =   300
               ItemData        =   "frmCareerHstRegister.frx":0139
               Left            =   4680
               List            =   "frmCareerHstRegister.frx":0146
               TabIndex        =   271
               Text            =   "����"
               Top             =   1920
               Width           =   975
            End
            Begin VB.ComboBox Combo9 
               Height          =   300
               ItemData        =   "frmCareerHstRegister.frx":015C
               Left            =   3600
               List            =   "frmCareerHstRegister.frx":016C
               TabIndex        =   270
               Text            =   "���в���ԭ��ģ��"
               Top             =   1200
               Width           =   1935
            End
            Begin VB.TextBox ctxt���в��� 
               Height          =   270
               Index           =   1
               Left            =   960
               MultiLine       =   -1  'True
               TabIndex        =   53
               Text            =   "frmCareerHstRegister.frx":019E
               Top             =   1200
               Width           =   2535
            End
            Begin VB.TextBox ctxt���� 
               Height          =   270
               Index           =   1
               Left            =   2640
               TabIndex        =   48
               Text            =   "0"
               Top             =   600
               Width           =   735
            End
            Begin VB.TextBox ctxt��� 
               Height          =   270
               Index           =   1
               Left            =   1800
               TabIndex        =   47
               Text            =   "0"
               Top             =   600
               Width           =   735
            End
            Begin VB.TextBox ctxt��� 
               Height          =   270
               Index           =   1
               Left            =   960
               TabIndex        =   46
               Text            =   " "
               Top             =   600
               Width           =   735
            End
            Begin VB.TextBox ctxt�д� 
               Height          =   270
               Index           =   1
               Left            =   120
               TabIndex        =   45
               Text            =   "0"
               Top             =   600
               Width           =   735
            End
            Begin VB.TextBox ctxt��̥ 
               Height          =   270
               Index           =   1
               Left            =   3600
               TabIndex        =   49
               Text            =   "0"
               Top             =   600
               Width           =   975
            End
            Begin VB.TextBox ctxt��̥ 
               Height          =   270
               Index           =   1
               Left            =   4560
               TabIndex        =   50
               Text            =   "0"
               Top             =   600
               Width           =   1095
            End
            Begin VB.TextBox ctxt��Ů���� 
               Height          =   270
               Index           =   1
               Left            =   3600
               TabIndex        =   52
               Top             =   1920
               Width           =   975
            End
            Begin VB.Label Label95 
               Caption         =   "������"
               Height          =   255
               Left            =   120
               TabIndex        =   304
               Top             =   960
               Width           =   855
            End
            Begin VB.Label Label78 
               Caption         =   "�������ڣ�"
               Height          =   225
               Left            =   1440
               TabIndex        =   295
               Top             =   2040
               Width           =   975
            End
            Begin VB.Label Label77 
               Caption         =   "�������ڣ�"
               Height          =   225
               Left            =   1440
               TabIndex        =   294
               Top             =   1680
               Width           =   975
            End
            Begin VB.Label Label76 
               Caption         =   "����Ů����"
               Height          =   255
               Left            =   120
               TabIndex        =   291
               Top             =   2040
               Width           =   975
            End
            Begin VB.Label Label75 
               Caption         =   "�����к���"
               Height          =   255
               Left            =   120
               TabIndex        =   290
               Top             =   1680
               Width           =   975
            End
            Begin VB.Label Label79 
               AutoSize        =   -1  'True
               Caption         =   "���в���ԭ��"
               Height          =   180
               Left            =   960
               TabIndex        =   184
               Top             =   960
               Width           =   1260
            End
            Begin VB.Label Label80 
               AutoSize        =   -1  'True
               Caption         =   "������"
               Height          =   180
               Left            =   2640
               TabIndex        =   183
               Top             =   360
               Width           =   540
            End
            Begin VB.Label Label81 
               AutoSize        =   -1  'True
               Caption         =   "�����"
               Height          =   180
               Left            =   1800
               TabIndex        =   182
               Top             =   360
               Width           =   540
            End
            Begin VB.Label Label82 
               AutoSize        =   -1  'True
               Caption         =   "�����"
               Height          =   180
               Left            =   960
               TabIndex        =   181
               Top             =   360
               Width           =   540
            End
            Begin VB.Label Label84 
               AutoSize        =   -1  'True
               Caption         =   "�дΣ�"
               Height          =   180
               Left            =   120
               TabIndex        =   180
               Top             =   360
               Width           =   540
            End
            Begin VB.Label Label85 
               AutoSize        =   -1  'True
               Caption         =   "��̥��"
               Height          =   180
               Left            =   3480
               TabIndex        =   179
               Top             =   360
               Width           =   540
            End
            Begin VB.Label Label86 
               AutoSize        =   -1  'True
               Caption         =   "��̥��"
               Height          =   180
               Left            =   4560
               TabIndex        =   178
               Top             =   360
               Width           =   540
            End
            Begin VB.Label Label89 
               AutoSize        =   -1  'True
               Caption         =   "��Ů����״����"
               Height          =   300
               Left            =   3600
               TabIndex        =   177
               Top             =   1680
               Width           =   1260
            End
         End
         Begin VB.Frame Frame17 
            Caption         =   "�̾�ʷ"
            ForeColor       =   &H000080FF&
            Height          =   2295
            Left            =   5880
            TabIndex        =   174
            Top             =   1320
            Width           =   5055
            Begin VB.TextBox ctxt���� 
               Height          =   270
               Index           =   1
               Left            =   960
               TabIndex        =   313
               Top             =   2040
               Width           =   855
            End
            Begin VB.ComboBox ccmb���� 
               Height          =   300
               Index           =   1
               Left            =   3360
               TabIndex        =   312
               Top             =   960
               Width           =   1455
            End
            Begin VB.ComboBox ccmb���� 
               Height          =   300
               Index           =   1
               Left            =   960
               TabIndex        =   311
               Top             =   960
               Width           =   1335
            End
            Begin VB.TextBox ctxt���� 
               Height          =   270
               Index           =   1
               Left            =   840
               TabIndex        =   302
               Top             =   1680
               Width           =   975
            End
            Begin VB.TextBox ctxt���� 
               Height          =   270
               Index           =   1
               Left            =   3240
               TabIndex        =   299
               Top             =   1680
               Width           =   975
            End
            Begin VB.TextBox ctxtMore 
               Height          =   375
               Index           =   1
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   54
               Top             =   480
               Width           =   4695
            End
            Begin VB.TextBox ctxt������ 
               Height          =   270
               Index           =   1
               Left            =   3360
               TabIndex        =   55
               Top             =   1320
               Width           =   975
            End
            Begin VB.TextBox ctxt������ 
               Height          =   270
               Index           =   1
               Left            =   840
               TabIndex        =   56
               Top             =   1320
               Width           =   975
            End
            Begin VB.Label Label111 
               Caption         =   "��"
               Height          =   255
               Left            =   1920
               TabIndex        =   314
               Top             =   2040
               Width           =   375
            End
            Begin VB.Label Label98 
               Caption         =   "����ʱ����"
               Height          =   255
               Left            =   120
               TabIndex        =   310
               Top             =   2040
               Width           =   1215
            End
            Begin VB.Label Label97 
               Caption         =   "���Ƴ̶ȣ�"
               Height          =   255
               Left            =   2520
               TabIndex        =   309
               Top             =   960
               Width           =   1095
            End
            Begin VB.Label Label96 
               Caption         =   "���̶̳ȣ�"
               Height          =   255
               Left            =   120
               TabIndex        =   308
               Top             =   960
               Width           =   975
            End
            Begin VB.Label Label94 
               Caption         =   "��"
               Height          =   255
               Left            =   4320
               TabIndex        =   303
               Top             =   1680
               Width           =   495
            End
            Begin VB.Label Label93 
               Caption         =   "���䣺"
               Height          =   255
               Left            =   120
               TabIndex        =   301
               Top             =   1680
               Width           =   735
            End
            Begin VB.Label Label92 
               Caption         =   "��"
               Height          =   255
               Left            =   1920
               TabIndex        =   300
               Top             =   1680
               Width           =   495
            End
            Begin VB.Label Label88 
               Caption         =   "���䣺"
               Height          =   255
               Left            =   2520
               TabIndex        =   298
               Top             =   1680
               Width           =   720
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               Caption         =   "�����ס��������ʳϰ�ߡ��̾��Ⱥ�������"
               Height          =   180
               Index           =   0
               Left            =   120
               TabIndex        =   200
               Top             =   240
               Width           =   3420
            End
            Begin VB.Label Label87 
               AutoSize        =   -1  'True
               Caption         =   "ML/��"
               Height          =   180
               Left            =   4440
               TabIndex        =   186
               Top             =   1320
               Width           =   450
            End
            Begin VB.Label Label83 
               AutoSize        =   -1  'True
               Caption         =   "֧/��"
               Height          =   180
               Left            =   1920
               TabIndex        =   185
               Top             =   1320
               Width           =   450
            End
            Begin VB.Label Label90 
               AutoSize        =   -1  'True
               Caption         =   "��������"
               Height          =   180
               Left            =   2520
               TabIndex        =   176
               Top             =   1320
               Width           =   720
            End
            Begin VB.Label Label91 
               AutoSize        =   -1  'True
               Caption         =   "��������"
               Height          =   180
               Left            =   120
               TabIndex        =   175
               Top             =   1320
               Width           =   720
            End
         End
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H0000C000&
         Caption         =   "ȷ  ��"
         Height          =   375
         Left            =   -64800
         Style           =   1  'Graphical
         TabIndex        =   272
         Top             =   4320
         Width           =   1095
      End
      Begin VB.Frame Frame5 
         Caption         =   "����¼��"
         ForeColor       =   &H000080FF&
         Height          =   855
         Left            =   720
         TabIndex        =   255
         Top             =   4080
         Width           =   5295
         Begin VB.CommandButton Command3 
            BackColor       =   &H8000000B&
            Caption         =   "ȷ��"
            Height          =   375
            Left            =   4200
            TabIndex        =   259
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox ctxtxinlv 
            Height          =   270
            Left            =   960
            TabIndex        =   256
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label66 
            Height          =   255
            Left            =   3120
            TabIndex        =   266
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label65 
            Caption         =   "����"
            Height          =   375
            Left            =   240
            TabIndex        =   258
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label64 
            Caption         =   "��/��"
            Height          =   255
            Left            =   2400
            TabIndex        =   257
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00008080&
         Caption         =   "�޸�"
         Height          =   375
         Left            =   -73200
         TabIndex        =   246
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000C000&
         Caption         =   "����"
         Height          =   375
         Left            =   -74760
         TabIndex        =   245
         Top             =   480
         Width           =   1215
      End
      Begin VB.Frame Frame6 
         Caption         =   "����ʷ"
         ForeColor       =   &H000080FF&
         Height          =   975
         Left            =   -68880
         TabIndex        =   232
         Top             =   4320
         Width           =   5175
         Begin VB.TextBox ctxt���� 
            Height          =   615
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   0
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label Label27 
            Caption         =   "��ʾ:�����������Ŵ��Լ�����ѪҺ�������򲡡���Ѫѹ�����񾭾����Լ�������������˲���"
            Height          =   615
            Left            =   2520
            TabIndex        =   250
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "�¾�ʷ"
         ForeColor       =   &H000080FF&
         Height          =   975
         Left            =   -74640
         TabIndex        =   233
         Top             =   4320
         Width           =   5775
         Begin VB.TextBox ctxtͣ�� 
            Height          =   270
            Left            =   3360
            TabIndex        =   98
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox ctxtĩ���¾� 
            Height          =   270
            Left            =   960
            TabIndex        =   97
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox ctxt���� 
            Height          =   270
            Left            =   4080
            TabIndex        =   96
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox ctxt���� 
            Height          =   270
            Left            =   2400
            TabIndex        =   95
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox ctxt���� 
            Height          =   270
            Left            =   600
            TabIndex        =   94
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "ͣ�����䣺"
            Height          =   180
            Index           =   2
            Left            =   2400
            TabIndex        =   239
            Top             =   600
            Width           =   900
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "ĩ���¾���"
            Height          =   180
            Index           =   2
            Left            =   120
            TabIndex        =   238
            Top             =   600
            Width           =   900
         End
         Begin VB.Label Label4 
            Caption         =   "Label4"
            Height          =   15
            Index           =   2
            Left            =   720
            TabIndex        =   237
            Top             =   720
            Width           =   135
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "���ڣ�"
            Height          =   180
            Index           =   2
            Left            =   3600
            TabIndex        =   236
            Top             =   240
            Width           =   540
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "���ڣ�"
            Height          =   180
            Index           =   2
            Left            =   1920
            TabIndex        =   235
            Top             =   240
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "������"
            Height          =   180
            Index           =   2
            Left            =   120
            TabIndex        =   234
            Top             =   240
            Width           =   540
         End
      End
      Begin VB.TextBox ctxtOther 
         Height          =   495
         Left            =   -74640
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   100
         Top             =   5400
         Width           =   10935
      End
      Begin VB.Frame Frame11 
         Caption         =   "�����Թ���ʷ  "
         ForeColor       =   &H000080FF&
         Height          =   1455
         Left            =   -74760
         TabIndex        =   107
         Top             =   2880
         Width           =   11055
         Begin VB.CommandButton ccmdok 
            BackColor       =   &H0000C000&
            Caption         =   "ȷ  ��"
            Height          =   375
            Left            =   9960
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   960
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox ctxtfangshe 
            Height          =   270
            Left            =   5400
            TabIndex        =   263
            Top             =   480
            Width           =   1575
         End
         Begin VB.ComboBox Combo1 
            Height          =   300
            ItemData        =   "frmCareerHstRegister.frx":01A2
            Left            =   8520
            List            =   "frmCareerHstRegister.frx":01AC
            TabIndex        =   251
            Text            =   "��ѡ��"
            Top             =   480
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CheckBox Chkokclear 
            Caption         =   "ȷ�������"
            Height          =   255
            Left            =   8400
            TabIndex        =   35
            Top             =   1080
            Width           =   1335
         End
         Begin VB.ComboBox ctxt�������� 
            Height          =   300
            Left            =   6960
            TabIndex        =   32
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox ctxt������ 
            Height          =   270
            Left            =   5400
            TabIndex        =   34
            Top             =   1080
            Width           =   1815
         End
         Begin VB.TextBox ctxt������ 
            Height          =   270
            Left            =   9480
            TabIndex        =   33
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox ctxt�������� 
            Height          =   855
            Left            =   240
            TabIndex        =   31
            Top             =   480
            Width           =   4935
         End
         Begin VB.Label Label41 
            Caption         =   "���������ࣺ"
            Height          =   255
            Left            =   5400
            TabIndex        =   111
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label40 
            Caption         =   "ÿ�չ���ʱ����������"
            Height          =   255
            Left            =   5400
            TabIndex        =   110
            Top             =   840
            Width           =   2055
         End
         Begin VB.Label Label39 
            Caption         =   "�ۻ�����������"
            Height          =   255
            Left            =   8520
            TabIndex        =   109
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label38 
            Caption         =   "��������ʷ��"
            Height          =   255
            Left            =   240
            TabIndex        =   108
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "���һ�����¼��"
         ForeColor       =   &H000080FF&
         Height          =   3135
         Left            =   720
         TabIndex        =   143
         Top             =   720
         Width           =   6135
         Begin VB.CommandButton Comd��д 
            Caption         =   "��д�������"
            Height          =   495
            Left            =   3720
            TabIndex        =   322
            Top             =   240
            Width           =   1815
         End
         Begin VB.TextBox ctxt����ָ�� 
            Height          =   270
            Left            =   3960
            TabIndex        =   320
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox ctxt���� 
            Height          =   270
            Left            =   3960
            TabIndex        =   319
            Top             =   840
            Width           =   1335
         End
         Begin VB.ComboBox comb���� 
            Height          =   300
            Left            =   960
            TabIndex        =   316
            Top             =   2760
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.ComboBox combӪ�� 
            Height          =   300
            Left            =   960
            TabIndex        =   1
            Text            =   "Combo1"
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox ctxt��� 
            Height          =   270
            Left            =   960
            TabIndex        =   2
            Top             =   840
            Width           =   1335
         End
         Begin VB.TextBox ctxt���� 
            Height          =   270
            Left            =   960
            TabIndex        =   3
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox ctxt����ѹ 
            Height          =   270
            Left            =   960
            TabIndex        =   4
            Top             =   1800
            Width           =   1335
         End
         Begin VB.TextBox ctxt����ѹ 
            Height          =   270
            Left            =   960
            TabIndex        =   5
            Top             =   2280
            Width           =   1335
         End
         Begin VB.Label Lab��� 
            Caption         =   "����ָ��"
            Height          =   255
            Index           =   7
            Left            =   3120
            TabIndex        =   318
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Lab��� 
            Caption         =   "����"
            Height          =   255
            Index           =   6
            Left            =   3120
            TabIndex        =   317
            Top             =   840
            Width           =   495
         End
         Begin VB.Label Lab��� 
            Caption         =   "����"
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   315
            Top             =   2760
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label Lab��� 
            Caption         =   "����ѹ"
            Height          =   375
            Index           =   4
            Left            =   240
            TabIndex        =   153
            Top             =   2280
            Width           =   615
         End
         Begin VB.Label Label56 
            Caption         =   "mmHg"
            Height          =   255
            Left            =   2400
            TabIndex        =   152
            Top             =   2280
            Width           =   375
         End
         Begin VB.Label Lab��� 
            Caption         =   "Ӫ��"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   150
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Lab��� 
            Caption         =   "���"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   149
            Top             =   840
            Width           =   495
         End
         Begin VB.Label Label35 
            Caption         =   "cm"
            Height          =   255
            Index           =   0
            Left            =   2400
            TabIndex        =   148
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Lab��� 
            Caption         =   "����"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   147
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label Label51 
            Caption         =   "kg"
            Height          =   255
            Left            =   2400
            TabIndex        =   146
            Top             =   1320
            Width           =   375
         End
         Begin VB.Label Lab��� 
            Caption         =   "����ѹ"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   145
            Top             =   1800
            Width           =   615
         End
         Begin VB.Label Label55 
            Caption         =   "mmHg"
            Height          =   255
            Left            =   2400
            TabIndex        =   144
            Top             =   1800
            Width           =   375
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid cgrd֢״ 
         Height          =   2535
         Left            =   -74760
         TabIndex        =   142
         Top             =   3300
         Width           =   10935
         _cx             =   2088782680
         _cy             =   2088767863
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
         SelectionMode   =   1
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
         FormatString    =   $"frmCareerHstRegister.frx":01C0
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
      Begin VSFlex8Ctl.VSFlexGrid cgrd��ʷ 
         Height          =   2415
         Left            =   -74760
         TabIndex        =   141
         Top             =   3300
         Width           =   9735
         _cx             =   2088780563
         _cy             =   2088767652
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
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   "ϵͳ���|���|��������|�������|��ϵ�λ|���ƾ���|ת��"
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
      Begin VB.CommandButton Ccmdɾ�� 
         BackColor       =   &H008080FF&
         Caption         =   "ɾ  ��"
         Height          =   375
         Left            =   -64800
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   5280
         Width           =   1095
      End
      Begin VB.CommandButton Ccmd�޸� 
         BackColor       =   &H0000C0C0&
         Caption         =   "��  ��"
         Height          =   375
         Left            =   -64800
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   4800
         Width           =   1095
      End
      Begin VB.CommandButton Ccmdcancel��ʷ 
         BackColor       =   &H008080FF&
         Caption         =   "ɾ  ��"
         Height          =   375
         Left            =   -64920
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   4860
         Width           =   1215
      End
      Begin VB.CommandButton Cmmdmody��ʷ 
         BackColor       =   &H0000C0C0&
         Caption         =   "��  ��"
         Height          =   375
         Left            =   -64920
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   4020
         Width           =   1215
      End
      Begin VB.Frame Frame15 
         Caption         =   "�Ծ�֢״���¼��   "
         ForeColor       =   &H000080FF&
         Height          =   2535
         Left            =   -74760
         TabIndex        =   127
         Top             =   780
         Width           =   10935
         Begin VB.TextBox ctxt��Ŀ 
            Height          =   1935
            Left            =   5040
            MultiLine       =   -1  'True
            TabIndex        =   8
            Top             =   360
            Width           =   2535
         End
         Begin VB.CheckBox Chkokclear֢״ 
            Caption         =   "��������"
            Height          =   255
            Left            =   1680
            TabIndex        =   10
            Top             =   1860
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            ForeColor       =   &H000000FF&
            Height          =   1935
            Left            =   7920
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   244
            Text            =   "frmCareerHstRegister.frx":0257
            Top             =   360
            Width           =   2535
         End
         Begin VB.CommandButton ccmdcancel֢״ 
            BackColor       =   &H008080FF&
            Caption         =   "ɾ  ��"
            Height          =   375
            Left            =   3000
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   1800
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CommandButton ccmbmody֢״ 
            BackColor       =   &H0000C0C0&
            Caption         =   "��  ��"
            Height          =   375
            Left            =   1680
            Style           =   1  'Graphical
            TabIndex        =   243
            Top             =   1800
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.ComboBox ccmb���� 
            Height          =   300
            Left            =   1320
            TabIndex        =   6
            Top             =   360
            Width           =   2535
         End
         Begin VB.ListBox clst��Ŀ 
            Height          =   1950
            Left            =   5040
            Style           =   1  'Checkbox
            TabIndex        =   7
            Top             =   360
            Width           =   2535
         End
         Begin VB.ComboBox ctxt�̶� 
            Height          =   300
            ItemData        =   "frmCareerHstRegister.frx":02FA
            Left            =   1320
            List            =   "frmCareerHstRegister.frx":02FC
            TabIndex        =   9
            Top             =   840
            Width           =   2535
         End
         Begin VB.CommandButton Ccmdok֢״ 
            BackColor       =   &H0000C000&
            Caption         =   "ȷ��"
            Height          =   375
            Left            =   360
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   1800
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label25 
            Caption         =   "֢״��λ��"
            Height          =   255
            Left            =   120
            TabIndex        =   242
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label67 
            Caption         =   "֢״��Ŀ��"
            Height          =   255
            Left            =   4080
            TabIndex        =   129
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label68 
            AutoSize        =   -1  'True
            Caption         =   "�̶ȣ�"
            Height          =   180
            Left            =   360
            TabIndex        =   128
            Top             =   840
            Width           =   540
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "����ʷ  "
         ForeColor       =   &H000080FF&
         Height          =   2055
         Left            =   -74760
         TabIndex        =   112
         Top             =   720
         Width           =   11055
         Begin VB.TextBox ctxt�Ӵ� 
            Height          =   270
            Left            =   1920
            TabIndex        =   288
            Top             =   1680
            Width           =   1935
         End
         Begin VB.TextBox ctxt���� 
            Height          =   300
            Left            =   8880
            TabIndex        =   278
            Text            =   "����"
            Top             =   1200
            Width           =   2055
         End
         Begin VB.TextBox ctxt��ʼ 
            Height          =   300
            Left            =   8880
            TabIndex        =   277
            Text            =   "����"
            Top             =   480
            Width           =   2055
         End
         Begin VB.ComboBox Combo4 
            Height          =   300
            Left            =   240
            TabIndex        =   261
            Text            =   "Combo4"
            Top             =   1200
            Width           =   1575
         End
         Begin VB.TextBox ctxtweihai 
            Height          =   270
            Left            =   6360
            TabIndex        =   260
            Top             =   1200
            Width           =   1815
         End
         Begin MSComCtl2.DTPicker ctxt����1 
            Height          =   300
            Left            =   5160
            TabIndex        =   30
            Top             =   1680
            Visible         =   0   'False
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            Format          =   60030976
            CurrentDate     =   41013
         End
         Begin MSComCtl2.DTPicker ctxt��ʼ1 
            Height          =   300
            Left            =   4320
            TabIndex        =   26
            Top             =   1560
            Visible         =   0   'False
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            Format          =   60030976
            CurrentDate     =   41013
         End
         Begin MSComCtl2.DTPicker ctxt�Ӵ�ʱ�� 
            Height          =   300
            Left            =   9000
            TabIndex        =   29
            Top             =   1680
            Visible         =   0   'False
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   529
            _Version        =   393216
            Format          =   60030976
            CurrentDate     =   41013
         End
         Begin VB.ComboBox ctxtΣ������ 
            Height          =   300
            Left            =   4320
            TabIndex        =   28
            Top             =   1200
            Width           =   1815
         End
         Begin VB.ComboBox ctxt���� 
            Height          =   300
            Left            =   4680
            TabIndex        =   24
            Top             =   480
            Width           =   1575
         End
         Begin VB.ComboBox ctxt���� 
            Height          =   300
            Left            =   2640
            TabIndex        =   23
            Top             =   480
            Width           =   1575
         End
         Begin VB.CheckBox Chk���� 
            Caption         =   "�Ƿ���乤����Ա"
            Height          =   255
            Left            =   7200
            TabIndex        =   130
            Top             =   1680
            Width           =   1815
         End
         Begin VB.TextBox ctxt��ע 
            Height          =   300
            Left            =   6600
            TabIndex        =   25
            Top             =   480
            Width           =   2055
         End
         Begin VB.TextBox ctxt���� 
            Height          =   360
            Left            =   1800
            MultiLine       =   -1  'True
            TabIndex        =   27
            Top             =   1200
            Width           =   2415
         End
         Begin VB.TextBox ctxt��λ 
            Height          =   300
            Left            =   240
            TabIndex        =   22
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label Label74 
            Caption         =   "�Ӵ�ʱ��(Сʱ/��)"
            Height          =   255
            Left            =   240
            TabIndex        =   287
            Top             =   1680
            Width           =   1695
         End
         Begin VB.Label Label28 
            Caption         =   "��ѡ�������ʩ"
            Height          =   255
            Left            =   240
            TabIndex        =   262
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label36 
            Caption         =   "�Ӵ�Σ�����ؽ���ʱ�䣺"
            Height          =   255
            Left            =   8880
            TabIndex        =   121
            Top             =   960
            Width           =   2175
         End
         Begin VB.Label Label37 
            Caption         =   "�Ӵ�Σ�����ؿ�ʼʱ�䣺"
            Height          =   255
            Left            =   8880
            TabIndex        =   120
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label Label43 
            Caption         =   "��ע��"
            Height          =   255
            Left            =   6600
            TabIndex        =   119
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label44 
            Caption         =   "������ʩ��"
            Height          =   255
            Left            =   1800
            TabIndex        =   118
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label45 
            Caption         =   "�Ӵ�Σ�����ؿ�ʼʱ�䣺"
            Height          =   255
            Left            =   9480
            TabIndex        =   117
            Top             =   1560
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Label Label46 
            Caption         =   "�Ӵ�ְҵ��Σ�����ࣺ"
            Height          =   255
            Left            =   4320
            TabIndex        =   116
            Top             =   960
            Width           =   2055
         End
         Begin VB.Label Label47 
            Caption         =   "���֣�"
            Height          =   255
            Left            =   4680
            TabIndex        =   115
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label48 
            Caption         =   "���ţ�"
            Height          =   255
            Left            =   2760
            TabIndex        =   114
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label49 
            Caption         =   "������λ��"
            Height          =   255
            Left            =   240
            TabIndex        =   113
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "������ʷ���¼��   "
         ForeColor       =   &H000080FF&
         Height          =   2535
         Left            =   -74760
         TabIndex        =   101
         Top             =   720
         Width           =   11055
         Begin VB.TextBox ctxt������� 
            Height          =   330
            Left            =   120
            TabIndex        =   279
            Top             =   1440
            Width           =   2535
         End
         Begin VB.ComboBox Combo6 
            Height          =   300
            Left            =   5040
            TabIndex        =   267
            Text            =   "Combo6"
            Top             =   1440
            Width           =   975
         End
         Begin VB.ComboBox Combo5 
            Height          =   300
            Left            =   6240
            TabIndex        =   265
            Text            =   "Combo5"
            Top             =   720
            Width           =   4095
         End
         Begin VB.ComboBox Combo3 
            Height          =   300
            ItemData        =   "frmCareerHstRegister.frx":02FE
            Left            =   120
            List            =   "frmCareerHstRegister.frx":0300
            TabIndex        =   254
            Text            =   "Combo3"
            Top             =   720
            Width           =   1575
         End
         Begin MSComCtl2.DTPicker ctxt�������1 
            Height          =   330
            Left            =   120
            TabIndex        =   15
            Top             =   2040
            Visible         =   0   'False
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   582
            _Version        =   393216
            Format          =   60030976
            CurrentDate     =   41013
         End
         Begin VB.CheckBox Chkokclear��ʷ 
            Caption         =   "ȷ�������"
            Height          =   255
            Left            =   7680
            TabIndex        =   18
            Top             =   2040
            Width           =   1215
         End
         Begin VB.CommandButton ccmdok��ʷ 
            BackColor       =   &H0000C000&
            Caption         =   "ȷ��"
            Height          =   375
            Left            =   9000
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   1920
            Width           =   1335
         End
         Begin VB.TextBox ctxtת�� 
            Height          =   330
            Left            =   3480
            TabIndex        =   16
            Top             =   1440
            Width           =   1335
         End
         Begin VB.TextBox ctxt��ϵ�λ 
            Height          =   330
            Left            =   3480
            TabIndex        =   14
            Top             =   720
            Width           =   2535
         End
         Begin VB.TextBox ctxt���ƾ��� 
            Height          =   615
            Left            =   6240
            TabIndex        =   17
            Top             =   1080
            Width           =   4095
         End
         Begin VB.TextBox ctxt�������� 
            Height          =   330
            Left            =   1680
            TabIndex        =   13
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Label54 
            Caption         =   "��ѡ�񼲲�����:"
            Height          =   255
            Left            =   120
            TabIndex        =   264
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label58 
            Caption         =   "��ϵ�λ��"
            Height          =   255
            Left            =   3480
            TabIndex        =   106
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label59 
            Caption         =   "���ƾ���/������������"
            Height          =   255
            Left            =   6240
            TabIndex        =   105
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label Label60 
            Caption         =   "�������ƣ�"
            Height          =   255
            Left            =   1680
            TabIndex        =   104
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label61 
            Caption         =   "������ڣ�"
            Height          =   255
            Left            =   120
            TabIndex        =   103
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label Label62 
            Caption         =   "ת�飺"
            Height          =   255
            Left            =   3480
            TabIndex        =   102
            Top             =   1200
            Width           =   1095
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid cgrdzzxw 
         Height          =   4335
         Left            =   -74760
         TabIndex        =   247
         Top             =   960
         Width           =   11535
         _cx             =   20346
         _cy             =   7646
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
         SelectionMode   =   1
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
         FormatString    =   $"frmCareerHstRegister.frx":0302
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
      Begin VSFlex8Ctl.VSFlexGrid cgrdְҵʷ 
         Height          =   1455
         Left            =   -74760
         TabIndex        =   289
         Top             =   4320
         Width           =   9855
         _cx             =   2088780775
         _cy             =   2088765958
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
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   16
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmCareerHstRegister.frx":0382
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "������"
         Height          =   180
         Index           =   0
         Left            =   -74640
         TabIndex        =   240
         Top             =   5460
         Width           =   540
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "������Ϣ"
      Height          =   1935
      Left            =   120
      TabIndex        =   122
      Top             =   840
      Width           =   13455
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   4320
         ScaleHeight     =   1785
         ScaleWidth      =   1545
         TabIndex        =   123
         Top             =   120
         Width           =   1575
      End
      Begin VB.TextBox ctxtsysno 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   99
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label������� 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7200
         TabIndex        =   274
         Top             =   1680
         Width           =   90
      End
      Begin VB.Label Label70 
         Caption         =   "���  ���ͣ�"
         Height          =   255
         Left            =   6000
         TabIndex        =   273
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label LabelΣ������ 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7200
         TabIndex        =   249
         Top             =   1320
         Width           =   90
      End
      Begin VB.Label Label26 
         Caption         =   "Σ��  ���أ�"
         Height          =   255
         Left            =   6000
         TabIndex        =   248
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "����:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2760
         TabIndex        =   140
         Top             =   1560
         Width           =   570
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "�Ա�:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   360
         TabIndex        =   139
         Top             =   1560
         Width           =   570
      End
      Begin VB.Label Lab���� 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   138
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Lab�Ա� 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   137
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Lab��ְ�� 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7200
         TabIndex        =   136
         Top             =   960
         Width           =   90
      End
      Begin VB.Label Lab�ֹ��� 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7200
         TabIndex        =   135
         Top             =   600
         Width           =   90
      End
      Begin VB.Label Label30 
         Caption         =   "��  ְ  ��"
         Height          =   255
         Index           =   0
         Left            =   6000
         TabIndex        =   134
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label29 
         Caption         =   "��  ��  �֣�"
         Height          =   255
         Left            =   6000
         TabIndex        =   133
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lab��λ 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7200
         TabIndex        =   132
         Top             =   240
         Width           =   90
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "�ֹ�����λ��"
         Height          =   255
         Left            =   6000
         TabIndex        =   131
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label Label52 
         AutoSize        =   -1  'True
         Caption         =   "���:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   360
         TabIndex        =   126
         Top             =   360
         Width           =   570
      End
      Begin VB.Label Lab���� 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1080
         TabIndex        =   125
         Top             =   840
         Width           =   3015
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         Caption         =   "����:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   360
         TabIndex        =   124
         Top             =   960
         Width           =   570
      End
   End
   Begin MSComDlg.CommonDialog ccmdFile 
      Left            =   0
      Top             =   500
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar ctbMain 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   230
      Top             =   0
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   1111
      ButtonWidth     =   820
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "cimg��ťͼ��"
      _Version        =   393216
      Begin MSCommLib.MSComm MSComm1 
         Left            =   2880
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
      End
      Begin VB.CheckBox ChkClear 
         Caption         =   "��������"
         Height          =   255
         Left            =   9840
         TabIndex        =   231
         Top             =   240
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   3
         Left            =   7440
         Top             =   240
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
Attribute VB_Name = "frmCareerHstRegt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'���ƣ�ְҵ��ʷ(�ܼ��߸�����Ϣ)¼��
'������Private Sub ctxtsysno_LostFocus()  ��ϵͳ����ı���ʧȥ���㣬���
'       �����Ա������Ϣ�������������Ա���Ƭ��
'���ܣ�ְҵ��ʷ(�ܼ��߸�����Ϣ)¼������Ϣ¼�룬�޸ģ�ɾ��
'���ߣ�Yunle Liu
'ʱ�䣺2012.03
'********************************************************************

Option Explicit

Private WithEvents mobjGUI As cls����ͨ�ö���
Attribute mobjGUI.VB_VarHelpID = -1
Private mblninuse As Boolean
Private mblnSys As Boolean
Private lobjPsLifeHst As Object
Private lobjPsWorkHst As Object
Private mcolIndex As Collection    'ְҵʷ
Private mcolIndexwkdis As Collection   '������ʷ
Private mcolindexzz As Collection       '�Ծ�֢״
Private mintrow As Integer     '��ǰ�޸ĵ��кţ�����ְҵʷ   modify by lanchao 2015-9-16
Private jmintrow As Integer     '��ǰ�޸ĵ��к�,���������ʷ modify by lanchao 2015-9-16
Private lobjInDtBase As Object   '���浽���ݿ�
Private mobj��� As Object
Private mcol�����Ŀ As New Collection
Private mIndex As String
Public sysno As String
Public selectsysno As String
Public selectzz As String

Public selectcd As String

Public selectcxrq As String

Private Sub ccmb����_Change()
    Dim lobjRec As Object
    Dim i As Integer
    If ccmb����.Text = "" Then clst��Ŀ.Clear: Exit Sub
    '��ȡ�Ծ�֢״��Ŀ
    dasubSetQueryTimeout 6000
    Set lobjRec = dafuncGetData("select ���� from ϵͳ����_�ֵ�_�ֵ����ݱ� where Parent in (select InnerID from ϵͳ����_�ֵ�_�ֵ����ݱ� where ���� = '" & ccmb����.Text & " ')")
    clst��Ŀ.Clear
'    ccmb����.AddItem ""
    If ccmb����.Text = "����" Or ccmb����.Text = "" Then
        clst��Ŀ.Visible = False
        ctxt��Ŀ.Visible = True
    Else
        clst��Ŀ.Visible = True
        ctxt��Ŀ.Visible = False
    End If
    If (lobjRec.EOF Or lobjRec.BOF) Then
        clst��Ŀ.Clear
        clst��Ŀ.Visible = False
        ctxt��Ŀ.Visible = True
        Exit Sub
    End If
    For i = 1 To lobjRec.RecordCount
        clst��Ŀ.AddItem lobjRec("����")
        clst��Ŀ.ItemData(clst��Ŀ.NewIndex) = i
        lobjRec.MoveNext
    Next
    clst��Ŀ.ListIndex = 0
    
End Sub

Private Sub ccmb����_Click()
    Dim lobjRec As Object
    Dim i As Integer
    If ccmb����.Text = "" Then clst��Ŀ.Clear: Exit Sub
    '��ȡ�Ծ�֢״��Ŀ
    dasubSetQueryTimeout 6000
    Set lobjRec = dafuncGetData("select ���� from ϵͳ����_�ֵ�_�ֵ����ݱ� where Parent in (select InnerID from ϵͳ����_�ֵ�_�ֵ����ݱ� where ���� = '" & ccmb����.Text & " ')")
    clst��Ŀ.Clear
'    ccmb����.AddItem ""
    If ccmb����.Text = "����" Or ccmb����.Text = "" Then
        clst��Ŀ.Visible = False
        ctxt��Ŀ.Visible = True
    Else
        clst��Ŀ.Visible = True
        ctxt��Ŀ.Visible = False
    End If
    If (lobjRec.EOF Or lobjRec.BOF) Then
        clst��Ŀ.Clear
        clst��Ŀ.Visible = False
        ctxt��Ŀ.Visible = True
        Exit Sub
    End If
    For i = 1 To lobjRec.RecordCount
        clst��Ŀ.AddItem lobjRec("����")
        clst��Ŀ.ItemData(clst��Ŀ.NewIndex) = i
        lobjRec.MoveNext
    Next
    clst��Ŀ.ListIndex = 0
    
End Sub

'���δ�������Ա��һЩ��Ϣ��������
Private Sub Ccmb���_Click(Index As Integer)
    On Error GoTo errHandler
    '�޸��ˣ����� 2012.12.06
    'bug�ţ�0000045
    '˵�����޸�if����������Ϊ��ʱ�����ctxtmatehelh��ctxtmatejob��ա� ����
    If Trim(Ccmb���(Index).Text) = "δ��" Or Trim(Ccmb���(Index).Text) = "" Then
        If Index <> 2 Then
            ctxtmatehelh(Index).Enabled = False
            ctxtmarrydate(Index).Enabled = False
            ctxtmatejob(Index).Enabled = False
            ctxtmateradioac(Index).Enabled = False
            ctxtmatehelh(Index).Text = ""
            ctxtmatejob(Index).Text = ""
            ctxtmateradioac(Index).Text = ""
            ctxtmarrydate(Index).Text = ""
            subclear����ʷ
        Else
            ctxt������Ů(Index).Text = ""
            ctxt���(Index).Text = ""
            ctxt����(Index).Text = ""
            ctxt����(Index).Text = ""
            ctxt�쳣̥.Text = ""
        End If
        Frame3(Index).Enabled = False
    '2012.12.06    ����
    Else
        If Index <> 2 Then
            ctxtmatehelh(Index).Enabled = True
            ctxtmarrydate(Index).Enabled = True
            ctxtmatejob(Index).Enabled = True
            ctxtmateradioac(Index).Enabled = True
        End If
        Frame3(Index).Enabled = True
    End If
    Exit Sub
errHandler:
    sfsub������ "ְҵ��ʷ¼��", "frmcareerhstregt", "ccmb���_click", Err.Number, Err.Description, True
End Sub




'ɾ��   ְҵʷ
Private Sub Ccmdɾ��_Click()
    Dim introw As String
    Dim i As Integer
    On Error GoTo errHandler
    If cgrdְҵʷ.Row = 0 Or cgrdְҵʷ.Row > cgrdְҵʷ.Rows - 1 Then
        MsgBox "��ѡ��Ҫɾ������Ϣ��", vbInformation, "ϵͳ��ʾ"
        Exit Sub
    End If
    introw = cgrdְҵʷ.Row
    If introw < cgrdְҵʷ.Rows - 1 Then
        For i = introw + 1 To cgrdְҵʷ.Rows - 1
        cgrdְҵʷ.Cell(flexcpText, i, mcolIndex("���")) = i - 1
        Next
    End If
    cgrdְҵʷ.RemoveItem cgrdְҵʷ.Row
    cgrdְҵʷ.AutoSize 0, cgrdְҵʷ.Cols - 1
    mintrow = 0
    Exit Sub
errHandler:
    sfsub������ "ְҵ��ʷ¼��", "frmcareerhstregt", "ccmbɾ��_click", Err.Number, Err.Description, True
End Sub
'ɾ��  ְҵ��ʷ
Private Sub Ccmdcancel��ʷ_Click()
    Dim zyintrow As String
    Dim i As Integer
    On Error GoTo errHandler
    If cgrd��ʷ.Row = 0 Or cgrd��ʷ.Row > cgrd��ʷ.Rows - 1 Then
        MsgBox "��ѡ��Ҫɾ������Ϣ��", vbInformation, "ϵͳ��ʾ"
        Exit Sub
    End If
    zyintrow = cgrd��ʷ.Row
    If zyintrow < cgrd��ʷ.Rows - 1 Then
        For i = zyintrow + 1 To cgrd��ʷ.Rows - 1
        cgrd��ʷ.Cell(flexcpText, i, mcolIndex("���")) = i - 1
        Next
    End If
    cgrd��ʷ.RemoveItem cgrd��ʷ.Row
    cgrd��ʷ.AutoSize 0, cgrd��ʷ.Cols - 1
    '��mintrow�滻��jmintrow ������ʷ��������  modify by lanchao 2015-9-16
'    mintrow = 0
     jmintrow = 0
    Exit Sub
errHandler:
    sfsub������ "ְҵ��ʷ¼��", "frmcareerhstregt", "ccmdcancle��ʷ_click", Err.Number, Err.Description, True
End Sub

'ɾ��  �Ծ�֢״
Private Sub ccmdcancel֢״_Click()
    Dim zjintrow As String
    Dim i As Integer
    On Error GoTo errHandler
    If cgrd֢״.Row = 0 Or cgrd֢״.Row > cgrd֢״.Rows - 1 Then
        MsgBox "��ѡ��Ҫɾ������Ϣ��", vbInformation, "ϵͳ��ʾ"
        Exit Sub
    End If
    zjintrow = cgrd֢״.Row
    If zjintrow < cgrd֢״.Rows - 1 Then
        For i = zjintrow + 1 To cgrd֢״.Rows - 1
        cgrd֢״.Cell(flexcpText, i, mcolIndex("���")) = i - 1
        Next
    End If
    cgrd֢״.RemoveItem cgrd֢״.Row
    cgrd֢״.AutoSize 0, cgrd֢״.Cols - 1
    zjintrow = 0
    Exit Sub
errHandler:
    sfsub������ "ְҵ��ʷ¼��", "frmcareerhstregt", "ccmdcancel֢״_click", Err.Number, Err.Description, True
End Sub

'ϵͳ���|ְҵʷ���|������λ|����|����|ְҵΣ������|�Ӵ�ʱ��|������ʩ|��ע|����������|ÿ�չ�����|�ۻ�������|��������ʷ|��ʼʱ�������ʱ��|�Ƿ������
'ȷ��  ְҵʷ
Private Sub ccmdOk_Click()
    Dim i As Integer
    On Error GoTo errHandler
    '�ж��Ƿ����޸���Ϣ����Ϊ�޸Ĳ��������¼�¼
    If mintrow = 0 Then
        cgrdְҵʷ.Rows = cgrdְҵʷ.Rows + 1
        i = cgrdְҵʷ.Rows - 1
    Else
        i = mintrow
    End If
    cgrdְҵʷ.Cell(flexcpText, i, mcolIndex("���")) = i
    cgrdְҵʷ.Cell(flexcpText, i, mcolIndex("������λ")) = Trim(ctxt��λ.Text)
    cgrdְҵʷ.Cell(flexcpText, i, mcolIndex("����")) = Trim(ctxt����.Text)
    cgrdְҵʷ.Cell(flexcpText, i, mcolIndex("����")) = Trim(ctxt����.Text)
    '����ע�޸�Ϊ�Ӵ�ʱ����ʾ modify by lanchao 2015-9-6
    cgrdְҵʷ.Cell(flexcpText, i, mcolIndex("��ע")) = Trim(ctxt��ע.Text)
    '���ڸ�ʽ��Ҫ 2015-6-26 by lanchao
    cgrdְҵʷ.Cell(flexcpText, i, mcolIndex("��ʼʱ��")) = Trim(ctxt��ʼ.Text)
    cgrdְҵʷ.Cell(flexcpText, i, mcolIndex("����ʱ��")) = Trim(ctxt����.Text)
    cgrdְҵʷ.Cell(flexcpText, i, mcolIndex("Σ������")) = Trim(ctxtweihai.Text)
    cgrdְҵʷ.Cell(flexcpText, i, mcolIndex("������ʩ")) = Trim(ctxt����.Text)

    cgrdְҵʷ.Cell(flexcpText, i, mcolIndex("��������")) = Trim(ctxtfangshe.Text)
    cgrdְҵʷ.Cell(flexcpText, i, mcolIndex("ÿ�չ�����")) = Trim(ctxt������.Text)
    cgrdְҵʷ.Cell(flexcpText, i, mcolIndex("�ۻ�������")) = Trim(ctxt������.Text)
    cgrdְҵʷ.Cell(flexcpText, i, mcolIndex("��������ʷ")) = Trim(ctxt��������.Text)
    '�Ӵ�ʱ�����´� 2015-9-6 by lanchao
    cgrdְҵʷ.Cell(flexcpText, i, mcolIndex("�Ӵ�ʱ��")) = Trim(ctxt�Ӵ�.Text)
    '�ж��Ƿ������
    If Chk����.Value = 1 Then
        cgrdְҵʷ.Cell(flexcpText, i, mcolIndex("�Ƿ������")) = "��"
    Else
        cgrdְҵʷ.Cell(flexcpText, i, mcolIndex("�Ƿ������")) = "��"
    End If
    '�ж�ȷ�����Ƿ����
    If Chkokclear.Value = 1 Then
        Call subokclear
    End If
    cgrdְҵʷ.AutoSize 0, cgrdְҵʷ.Cols - 1
    mintrow = 0
    Exit Sub
errHandler:
    sfsub������ "ְҵ��ʷ¼��", "frmcareerhstregt", "ccmdok_click", Err.Number, Err.Description, True
End Sub

'ȷ��   ������ʷ
Private Sub ccmdok��ʷ_Click()
    Dim i As Integer
    On Error GoTo errHandler
    '�ж��Ƿ����޸���Ϣ����Ϊ�޸Ĳ��������¼�¼
    '��mintrow�޸�Ϊjmintrow���е������� modify by lanchao 2015-9-16
    If jmintrow = 0 Then
        cgrd��ʷ.Rows = cgrd��ʷ.Rows + 1
        i = cgrd��ʷ.Rows - 1
    Else
        i = jmintrow
    End If
    cgrd��ʷ.Cell(flexcpText, i, mcolIndexwkdis("���")) = i
    cgrd��ʷ.Cell(flexcpText, i, mcolIndexwkdis("��������")) = Trim(ctxt��������.Text)
    cgrd��ʷ.Cell(flexcpText, i, mcolIndexwkdis("��ϵ�λ")) = Trim(ctxt��ϵ�λ.Text)
    cgrd��ʷ.Cell(flexcpText, i, mcolIndexwkdis("�������")) = Trim(ctxt�������.Text)
    cgrd��ʷ.Cell(flexcpText, i, mcolIndexwkdis("ת��")) = Trim(ctxtת��.Text)
    cgrd��ʷ.Cell(flexcpText, i, mcolIndexwkdis("���ƾ���")) = Trim(ctxt���ƾ���.Text)
    If Chkokclear��ʷ.Value = 1 Then
        subokclear��ʷ
    End If
    jmintrow = 0
    cgrd��ʷ.AutoSize 0, cgrd��ʷ.Cols - 1
    Exit Sub
errHandler:
    sfsub������ "ְҵ��ʷ¼��", "frmcareerhstregt", "ccmdok��ʷ_click", Err.Number, Err.Description, True
End Sub

'ȷ��   �Ծ�֢״
Private Sub Ccmdok֢״_Click()
    Dim i As Integer, j As Integer
    On Error GoTo errHandler
    '�ж��Ƿ����޸���Ϣ����Ϊ�޸Ĳ��������¼�¼
    Dim zjmintrow As Integer
    For j = 0 To clst��Ŀ.ListCount - 1
        If clst��Ŀ.Selected(j) = True Then
            If zjmintrow = 0 Then
                cgrd֢״.Rows = cgrd֢״.Rows + 1
                i = cgrd֢״.Rows - 1
            Else
                i = zjmintrow
            End If
            cgrd֢״.Cell(flexcpText, i, mcolindexzz("���")) = i
            cgrd֢״.Cell(flexcpText, i, mcolindexzz("֢״")) = Trim(clst��Ŀ.List(j))
'            clst��Ŀ.RemoveItem (j)
'            cgrd֢״.Cell(flexcpText, i, mcolindexzz("����ʱ��")) = Trim(ctxt����ʱ��.Text)   ' 2015-11-27 by Ĳ��
            cgrd֢״.Cell(flexcpText, i, mcolindexzz("�̶�")) = Trim(ctxt�̶�.Text)
        End If
    Next
    
    If ccmb����.Text = "����" Or clst��Ŀ.ListCount - 1 < 0 Then
        If zjmintrow = 0 Then
            cgrd֢״.Rows = cgrd֢״.Rows + 1
            i = cgrd֢״.Rows - 1
        Else
            i = zjmintrow
        End If
        cgrd֢״.Cell(flexcpText, i, mcolindexzz("���")) = i
        cgrd֢״.Cell(flexcpText, i, mcolindexzz("֢״")) = Trim(ctxt��Ŀ.Text)
'        cgrd֢״.Cell(flexcpText, i, mcolindexzz("����ʱ��")) = Trim(ctxt����ʱ��.Text)        ' 2015-11-27 by Ĳ��
        cgrd֢״.Cell(flexcpText, i, mcolindexzz("�̶�")) = Trim(ctxt�̶�.Text)
    End If
    
    ccmb����_Click
    If Chkokclear֢״.Value = 1 Then
        subokclear֢״
    End If
    cgrd֢״.AutoSize 0, cgrd֢״.Cols - 1
    zjmintrow = 0
    Exit Sub
errHandler:
    sfsub������ "ְҵ��ʷ¼��", "frmcareerhstregt", "ccmdok֢״_click", Err.Number, Err.Description, True
End Sub

'�޸�   ְҵʷ
Private Sub Ccmd�޸�_Click()
    cgrdְҵʷ_DblClick
End Sub

Private Sub cgrdzzxw_Click()
' MsgBox (cgrdzzxw.TextMatrix(cgrdzzxw.RowSel, 0))
selectsysno = cgrdzzxw.TextMatrix(cgrdzzxw.RowSel, 0)
selectzz = cgrdzzxw.TextMatrix(cgrdzzxw.RowSel, 1)
selectcd = cgrdzzxw.TextMatrix(cgrdzzxw.RowSel, 2)
selectcxrq = cgrdzzxw.TextMatrix(cgrdzzxw.RowSel, 3)
End Sub




Private Sub Chk����_Click()
    If Chk����.Value = 1 Then
        Frame11.Enabled = True
    Else
        Frame11.Enabled = False
    End If
End Sub

'�޸�  ְҵ��ʷ
Private Sub Cmmdmody��ʷ_Click()
    cgrd��ʷ_DblClick
End Sub

'�޸�   �Ծ�֢״
'Private Sub ccmbmody֢״_Click()
'    cgrd֢״_DblClick
'End Sub

'˫��grid �޸�  ְҵ��ʷ
Private Sub cgrd��ʷ_DblClick()
    On Error GoTo errHandler
    If cgrd��ʷ.Row = 0 Or cgrd��ʷ.Row > cgrd��ʷ.Rows - 1 Then
        MsgBox "��ѡ��Ҫ�޸ĵ���Ϣ��", vbInformation, "ϵͳ��ʾ"
        Exit Sub
    End If
    '��mintrow�޸�Ϊjmintrow���е������� modify by lanchao 2015-9-16
    jmintrow = cgrd��ʷ.Cell(flexcpText, cgrd��ʷ.Row, mcolIndex("���"))
    ctxt��������.Text = cgrd��ʷ.Cell(flexcpText, cgrd��ʷ.Row, mcolIndexwkdis("��������"))
    ctxt��ϵ�λ.Text = cgrd��ʷ.Cell(flexcpText, cgrd��ʷ.Row, mcolIndexwkdis("��ϵ�λ"))
    ctxt�������.Text = cgrd��ʷ.Cell(flexcpText, cgrd��ʷ.Row, mcolIndexwkdis("�������"))
    ctxtת��.Text = cgrd��ʷ.Cell(flexcpText, cgrd��ʷ.Row, mcolIndexwkdis("ת��"))
    ctxt���ƾ���.Text = cgrd��ʷ.Cell(flexcpText, cgrd��ʷ.Row, mcolIndexwkdis("���ƾ���"))
    Exit Sub
errHandler:
    sfsub������ "ְҵ��ʷ¼��", "frmcareerhstregt", "cgrd��ʷ_dbclick", Err.Number, Err.Description, True
End Sub

'˫���޸�,����ı�������Ϣ   ְҵʷ
Private Sub cgrdְҵʷ_DblClick()
    On Error GoTo errHandler
    If cgrdְҵʷ.Row = 0 Or cgrdְҵʷ.Row > cgrdְҵʷ.Rows - 1 Then
        MsgBox "��ѡ��Ҫ�޸ĵ���Ϣ��", vbInformation, "ϵͳ��ʾ"
        Exit Sub
    End If
    mintrow = cgrdְҵʷ.Cell(flexcpText, cgrdְҵʷ.Row, mcolIndex("���"))
    ctxt��λ.Text = cgrdְҵʷ.Cell(flexcpText, cgrdְҵʷ.Row, mcolIndex("������λ"))
    ctxt����.Text = cgrdְҵʷ.Cell(flexcpText, cgrdְҵʷ.Row, mcolIndex("����"))
    ctxt����.Text = cgrdְҵʷ.Cell(flexcpText, cgrdְҵʷ.Row, mcolIndex("����"))
    '����ע�޸�Ϊ�Ӵ�ʱ����ʾ modify by lanchao 2015-9-6
    ctxt��ע.Text = cgrdְҵʷ.Cell(flexcpText, cgrdְҵʷ.Row, mcolIndex("��ע"))
    ctxt��ʼ.Text = cgrdְҵʷ.Cell(flexcpText, cgrdְҵʷ.Row, mcolIndex("��ʼʱ��"))
    ctxt����.Text = cgrdְҵʷ.Cell(flexcpText, cgrdְҵʷ.Row, mcolIndex("����ʱ��"))
    ctxtΣ������.Text = cgrdְҵʷ.Cell(flexcpText, cgrdְҵʷ.Row, mcolIndex("Σ������"))
    ctxt����.Text = cgrdְҵʷ.Cell(flexcpText, cgrdְҵʷ.Row, mcolIndex("������ʩ"))
    
    ctxtfangshe.Text = cgrdְҵʷ.Cell(flexcpText, cgrdְҵʷ.Row, mcolIndex("��������"))
    ctxt������.Text = cgrdְҵʷ.Cell(flexcpText, cgrdְҵʷ.Row, mcolIndex("ÿ�չ�����"))
    ctxt������.Text = cgrdְҵʷ.Cell(flexcpText, cgrdְҵʷ.Row, mcolIndex("�ۻ�������"))
    ctxt��������.Text = cgrdְҵʷ.Cell(flexcpText, cgrdְҵʷ.Row, mcolIndex("��������ʷ"))
    '�Ӵ�ʱ�����´� modify by lanchao 2015-9-6
    ctxt�Ӵ�.Text = cgrdְҵʷ.Cell(flexcpText, cgrdְҵʷ.Row, mcolIndex("�Ӵ�ʱ��"))
    Dim tjblable As String
    tjblable = Label�������.Caption
    
    Dim tmpstr As String
    tmpstr = cgrdְҵʷ.Cell(flexcpText, cgrdְҵʷ.Row, mcolIndex("�Ƿ������"))
    '��ʼ������ʱ�����乤����Ϣ¼���disable modify by lanchao 2015-8-17
    If tmpstr = "��" Then
        Chk����.Value = 1
        Frame11.Enabled = True
    Else
        Chk����.Value = 0
        Frame11.Enabled = False
    End If
    Exit Sub
errHandler:
    sfsub������ "ְҵ��ʷ¼��", "frmcareerhstregt", "cgrdְҵʷ_dblclick", Err.Number, Err.Description, True
End Sub

Private Sub Combo1_Click()
ctxt������.Text = Combo1.Text
Combo1.Visible = False
Combo1.Text = "��ѡ��"
End Sub

Private Sub Combo10_Click()
ctxt��Ů����(1).Text = Combo10.Text
End Sub

Private Sub Combo2_Click()
If ctxt����ʷ(2).Text = "" Then
ctxt����ʷ(2).Text = Combo2.Text
Else
ctxt����ʷ(2).Text = ctxt����ʷ(2).Text + "��" + Combo2.Text
End If
End Sub


Private Sub Combo3_Click()
Dim i As Integer
ctxt��������.Text = Combo3.Text
'2015-4-8 ��ΰ ���
'If Label������� = "8023����" And Combo3.Text <> "" Then
' Dim lobjRec As Object
'    Set lobjRec = pobjDict.FetchEx("ְҵ��" + Combo3.Text)
'    Combo5.Clear
'    Combo5.AddItem ""
'    For i = 1 To lobjRec.RecordCount
'        Combo5.AddItem lobjRec("����")
'        Combo5.ItemData(Combo5.NewIndex) = lobjRec("���")
'        lobjRec.MoveNext
'    Next
'   Combo5.ListIndex = 0
'   End If
End Sub

Private Sub Combo4_Click()
If ctxt����.Text = "" Then
ctxt����.Text = Combo4.Text
Else
ctxt����.Text = ctxt����.Text + "��" + Combo4.Text
End If
End Sub

Private Sub Combo5_Click()
If Combo5.ListIndex = 0 Then
ctxt���ƾ���.Text = ""
Else
ctxt���ƾ���.Text = Combo5.Text
ctxt���ƾ���.SetFocus
ctxt���ƾ���.SelStart = Len(ctxt���ƾ���.Text)


End If
End Sub

Private Sub Combo6_Click()
ctxtת��.Text = Combo6.Text

End Sub

Private Sub Combo7_Click()
ctxt���в���(0).Text = Combo7.Text
End Sub

Private Sub Combo8_Click()
ctxt��Ů����(0).Text = Combo8.Text
End Sub
'��ż���������ѡ 2015-6-26 by lanchao
Private Sub Combo11_Click()
ctxtmatehelh(0).Text = Combo11.Text
End Sub

Private Sub Combo9_Click()
ctxt���в���(1).Text = Combo9.Text
End Sub


Private Sub Comd��д_Click()
    Dim A, B, C As String
    A = Trim(Text��ʱ.Text)
    '����
    B = Trim(Split(A, "H")(0))
    C = Trim(Mid(B, 3, Len(B) - 2))
    If Left(C, 1) = "0" Then
        C = Right(C, Len(C) - 1)
    End If
    ctxt����.Text = C
    '���
    B = Trim(Split(A, "H")(1))
    C = Trim(Mid(B, 2, Len(B) - 1))
    ctxt���.Text = C
    '����ָ��
    Dim W, H, X As Single
    Dim Z As String
    W = Val(ctxt����.Text)
    H = Val(ctxt���.Text)  '��λcm
    H = H / 100           '��λת��m
    X = W / (H * H)
    X = Format(X, "0.0")
    Z = Str(X)
    ctxt����ָ��.Text = Z
    '����
    If X < 18.5 Then
        ctxt����.Text = "����"
    ElseIf X >= 18.5 And X < 24 Then
        ctxt����.Text = "����"
    ElseIf X >= 24 And X < 27.5 Then
        ctxt����.Text = "����"
    ElseIf X >= 27.5 And X < 30 Then
        ctxt����.Text = "��ȷ���"
    ElseIf X >= 30 And X < 35 Then
        ctxt����.Text = "�жȷ���"
    ElseIf X >= 35 Then
        ctxt����.Text = "�ضȷ���"
    End If
End Sub

Private Sub Command1_Click()

If cgrdzzxw.Row > 67 Then

MsgBox "�Ѿ���ӱ���ɹ���ֻ���޸���Ŀ��"

Else
frm֢״ѯ��.Show vbModal
End If

End Sub

Private Sub Command2_Click()
If cgrdzzxw.Row > 0 Then

frm֢״�޸�.Show vbModal


Else
 MsgBox "��ѡ����Ҫ�޸ĵ�֢״��Ŀ��"
End If
End Sub

Private Sub Command3_Click()
Dim conclusion As String
conclusion = "����"
If ctxtxinlv.Text <> "" And IsNumeric(ctxtxinlv.Text) Then

If CInt(ctxtxinlv.Text) > 100 Then
conclusion = "����"
End If
If CInt(ctxtxinlv.Text) < 60 Then
conclusion = "����"
End If

dafuncGetData ("update ְҵ�����_�����Ϣ_�ڿ� set �����='" & ctxtxinlv.Text & "'  , ���ҽʦ='" & um�û���� & "',�������='" & conclusion & "' where �����Ŀ='02002' and ϵͳ���='" & Trim(ctxtsysno.Text) & "'")
'dafuncGetData ("update ְҵ�����_�����Ϣ_�ڿ� set �����='" & ctxtxinlv.Text & "'  , ���ҽʦ='" & um�û���� & "' ,��дʱ��='" & Now & "',�������='" & conclusion & "' where �����Ŀ='02002' and ϵͳ���='" & Trim(ctxtsysno.Text) & "'")

Label66.Caption = "�ѱ��档"
ctxtxinlv.Text = ""

End If


End Sub

Private Sub Command4_Click()
Dim i As Integer
    On Error GoTo errHandler
    '�ж��Ƿ����޸���Ϣ����Ϊ�޸Ĳ��������¼�¼
    If mintrow = 0 Then
        cgrdְҵʷ.Rows = cgrdְҵʷ.Rows + 1
        i = cgrdְҵʷ.Rows - 1
    Else
        i = mintrow
    End If
    cgrdְҵʷ.Cell(flexcpText, i, mcolIndex("���")) = i
    cgrdְҵʷ.Cell(flexcpText, i, mcolIndex("������λ")) = Trim(ctxt��λ.Text)
    cgrdְҵʷ.Cell(flexcpText, i, mcolIndex("����")) = Trim(ctxt����.Text)
    cgrdְҵʷ.Cell(flexcpText, i, mcolIndex("����")) = Trim(ctxt����.Text)
    cgrdְҵʷ.Cell(flexcpText, i, mcolIndex("��ע")) = Trim(ctxt��ע.Text)
    '���ڸ�ʽ��Ҫ 2015-6-26 by lanchao ��Ҫ�Ӵ�ʱ��
    cgrdְҵʷ.Cell(flexcpText, i, mcolIndex("��ʼʱ��")) = Trim(ctxt��ʼ.Text)
    cgrdְҵʷ.Cell(flexcpText, i, mcolIndex("����ʱ��")) = Trim(ctxt����.Text)
    cgrdְҵʷ.Cell(flexcpText, i, mcolIndex("Σ������")) = Trim(ctxtweihai.Text)
    cgrdְҵʷ.Cell(flexcpText, i, mcolIndex("������ʩ")) = Trim(ctxt����.Text)
    '�Ӵ�ʱ��������ʾ 2015-9-6 by lanchao
    cgrdְҵʷ.Cell(flexcpText, i, mcolIndex("�Ӵ�ʱ��")) = Trim(ctxt�Ӵ�.Text)
    cgrdְҵʷ.Cell(flexcpText, i, mcolIndex("��������")) = Trim(ctxtfangshe.Text)
    cgrdְҵʷ.Cell(flexcpText, i, mcolIndex("ÿ�չ�����")) = Trim(ctxt������.Text)
    cgrdְҵʷ.Cell(flexcpText, i, mcolIndex("�ۻ�������")) = Trim(ctxt������.Text)
    cgrdְҵʷ.Cell(flexcpText, i, mcolIndex("��������ʷ")) = Trim(ctxt��������.Text)
    '�ж��Ƿ������
    If Chk����.Value = 1 Then
        cgrdְҵʷ.Cell(flexcpText, i, mcolIndex("�Ƿ������")) = "��"
    Else
        cgrdְҵʷ.Cell(flexcpText, i, mcolIndex("�Ƿ������")) = "��"
    End If
    '�ж�ȷ�����Ƿ����
    If Chkokclear.Value = 1 Then
        Call subokclear
    End If
    cgrdְҵʷ.AutoSize 0, cgrdְҵʷ.Cols - 1
    mintrow = 0
    Exit Sub
errHandler:
    sfsub������ "ְҵ��ʷ¼��", "frmcareerhstregt", "ccmdok_click", Err.Number, Err.Description, True
   
End Sub

'˫��grid �޸�  �Ծ�֢״
'Private Sub cgrd֢״_DblClick()
'    On Error GoTo errHandler
'    If cgrd֢״.Row = 0 Or cgrd֢״.Row > cgrd֢״.Rows - 1 Then
'        MsgBox "��ѡ��Ҫ�޸ĵ���Ϣ��", vbInformation, "ϵͳ��ʾ"
'        Exit Sub
'    End If
'    mintrow = cgrd֢״.Cell(flexcpText, cgrd֢״.Row, mcolIndex("���"))
''    clst��Ŀ.AddItem cgrd֢״.Cell(flexcpText, cgrd֢״.Row, mcolindexzz("֢״"))
'    ctxt����ʱ��.Text = cgrd֢״.Cell(flexcpText, cgrd֢״.Row, mcolindexzz("����ʱ��"))
'    ctxt�̶�.Text = cgrd֢״.Cell(flexcpText, cgrd֢״.Row, mcolindexzz("�̶�"))
'    Exit Sub
'errHandler:
'    sfsub������ "ְҵ��ʷ¼��", "frmcareerhstregt", "cgrd֢״_dblclick", Err.Number, Err.Description, True
'End Sub
'��������ڲ�����ʽ�ж� 2015-6-26 by lanchao
Private Sub ctxtmarrydate_LostFocus(Index As Integer)
'    If ctxtmarrydate(mIndex).Enabled Then
'        If DateDiff("d", ctxtmarrydate(mIndex).Value, Date) < 0 Then
'            MsgBox "�����������"
'            ctxtmarrydate(mIndex).Value = Date
'        End If
'    End If
End Sub

'����⵽�лس������º��Ƴ�����
Private Sub ctxtsysno_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        Ccmb���(mIndex).SetFocus
    End If
End Sub

'��ϵͳ����ı������߽������������Ա���˻�����Ϣ
Private Sub ctxtsysno_LostFocus()
    Dim lobjRec As Object
    Dim lobjRegt As Object
    Dim lobjlife As Object
    Dim registeragn As String
    On Error GoTo errHandler
    If Len(RTrim(ctxtsysno.Text)) < 5 Then
        MsgBox "ϵͳ��Ŵ������飡"
        ctxtsysno.Text = ""
        gatherclear   '������Ա��������Ҳ��������Ѿ�¼�����ʱ��ս���  2016-3-2 by Ĳ��
        ctxtsysno.SetFocus
        Exit Sub
    End If
' ���������  2016-3-2 by Ĳ��
'    '2012-04-14 �ڵ�� ��
'    '���ܻ������µ������Ա��Ϣ����ʱ�轫4����ʷ�Ĵ�������ȫ�����
'    subclear
'    subokclear
'    subokclear��ʷ
'    subokclear֢״
'    subClear���һ�����
'    '2012-04-14 �ڵ�� ��
    
    '��������ʷ����
    Set lobjPsLifeHst = CreateObject("ְҵ��ʷ¼��.clslifehstregt")
    lobjPsLifeHst.ϵͳ��� = Trim(ctxtsysno.Text)
    If Not lobjPsLifeHst.tmp�ѵǼ� Then
        ctxtsysno.SetFocus
        Exit Sub
    End If
    
    Lab����.Caption = lobjPsLifeHst.����
    Lab�Ա�.Caption = lobjPsLifeHst.�Ա�
    Lab����.Caption = lobjPsLifeHst.����
    lab��λ.Caption = lobjPsLifeHst.��λ����
    Lab�ֹ���.Caption = lobjPsLifeHst.�ֹ���
    Lab��ְ��.Caption = lobjPsLifeHst.ְ��
    LabelΣ������.Caption = lobjPsLifeHst.Σ������
    '2012-07-11 �ڵ�� ��
    '����֪��Ϣȫ���������
    ctxt��λ.Text = lobjPsLifeHst.��λ����
    ctxt����.Text = lobjPsLifeHst.�ֹ���
    '2012-07-11 �ڵ�� ��
    
    '��ȡ��Ƭ
    Set lobjRec = CreateObject("ְҵ������.clspersonexamed")
    lobjRec.ϵͳ��� = Trim(ctxtsysno.Text)
    Picture2.Picture = lobjRec.��Ƭ
    Picture2.Visible = True
    
    '2012-06-14 �ڵ�� ��
    subClear���һ�����
    subLoad���һ����� Trim(ctxtsysno.Text)
    '2012-06-14 �ڵ�� ��
    
    '�ж��Ƿ��ѽ��й� ְҵ��ʷ¼�� ����
    Set lobjRegt = CreateObject("ְҵ��ʷ¼��.clscareerhstmage")
    lobjRegt.ϵͳ��� = Trim(ctxtsysno.Text)
    registeragn = lobjRegt.�Ѽ�¼��־
    If registeragn = "100" Or registeragn = "101" Then
        MsgBox "�������Ա�����ڻ��������ϣ�", vbExclamation, "ϵͳ��ʾ"
        subclearps
        gatherclear  '������Ա��������Ҳ��������Ѿ�¼�����ʱ��ս���  2016-3-2 by Ĳ��
        ctxtsysno.SetFocus
        Exit Sub
    End If
    If registeragn = "2" And ���ʼǺ� = 0 Then
        If MsgBox("�������ѽ��й��ܼ��߸�����Ϣ�Ǽǣ�Ҫ�޸�����", vbYesNo + vbQuestion + vbDefaultButton2, "ϵͳ��ʾ") = vbNo Then
            subclearps
             gatherclear  '������Ա��������Ҳ��������Ѿ�¼�����ʱ��ս���  2016-3-2 by Ĳ��
            ctxtsysno.SetFocus
            Exit Sub
        End If
    End If
    If ���ʼǺ� = 1 Or registeragn = "2" Then
        ���ʼǺ� = 1
        '�޸�����ʷ
        Set lobjlife = lobjRegt.��������ʷ
        sub�޸�����ʷ lobjlife
        
        'sub�޸�ְҵʷ
        Set lobjlife = lobjRegt.ְҵʷ
'        cgrdְҵʷ.Clear
        '�ж����ݿ����Ƿ��ж�Ӧ�����ݣ�����в���ʾ(����Ĳ�ʷ���Ծ�֢״һ���Ĵ���)  2016-3-7 by Ĳ��
        If lobjlife.RecordCount > 0 Then
        Set cgrdְҵʷ.DataSource = lobjlife
'        Set cgrdְҵʷ.DataSource = lobjlife
        'cgrdְҵʷ.ColHidden(mcolIndex("ϵͳ���")) = True
        cgrdְҵʷ.ColHidden(0) = True
        'ְҵ��������ʾ�Ӵ�ʱ��
         If Label�������.Caption <> "ְҵ����" Then
         cgrdְҵʷ.ColHidden(10) = True
         cgrdְҵʷ.AutoSize 0, cgrdְҵʷ.Cols - 1
         End If
        End If
        'sub�޸Ĳ�ʷ
        Set lobjlife = lobjRegt.������ʷ
'        Set cgrd��ʷ.DataSource = lobjlife
        If lobjlife.RecordCount > 0 Then
         Set cgrd��ʷ.DataSource = lobjlife
        'cgrd��ʷ.ColHidden(mcolIndexwkdis("ϵͳ���")) = True
        cgrd��ʷ.ColHidden(0) = True
        cgrd��ʷ.AutoSize 0, cgrd��ʷ.Cols - 1
        End If
        'sub�޸��Ծ�֢״
        Set lobjlife = lobjRegt.�Ծ�֢״
'        Set cgrd֢״.DataSource = lobjlife
        If lobjlife.RecordCount > 0 Then
        Set cgrd֢״.DataSource = lobjlife
        'cgrd֢״.ColHidden(mcolindexzz("ϵͳ���")) = True
        cgrd֢״.ColHidden(0) = True
        cgrd֢״.AutoSize 0, cgrd֢״.Cols - 1
        End If
    Else
        ���ʼǺ� = 0
    End If
'    If Ccmb���(mIndex).Text = "�ѻ�" Or Ccmb���(mIndex).Text = "�ѻ�" Then
'        If mIndex <> 2 Then
'            ctxtmatehelh(mIndex).Enabled = True
'            ctxtmarrydate(mIndex).Enabled = True
'            ctxtmatejob(mIndex).Enabled = True
'            ctxtmateradioac(mIndex).Enabled = True
'        End If
'        Frame3(mIndex).Enabled = True
'    End If
    '�ж��Ա�
    If Trim(Lab�Ա�.Caption) = "��" Then
        Frame1.Enabled = False
    Else
        Frame1.Enabled = True
    End If
    Set lobjRec = Nothing
    
'    MsgBox "���߽������" '2016-3-1 by Ĳ��
    Exit Sub
errHandler:
    sfsub������ "ְҵ��ʷ¼��", "frmcareerhstregt", "ctxtsysno_lostfocus", Err.Number, Err.Description, True
End Sub

'2015-11-27 by Ĳ��

'Private Sub ctxt����ʱ��_LostFocus()
'
'    '2015-03-30 liuwei
'    'If DateDiff("d", ctxt����ʱ��.Text, Date) < 0 Then
'     '   MsgBox "����ʱ�����"
'     '   ctxt����ʱ��.Text = Date
'    'End If
'    'ע��������� ��ΰ2015-4-7
'    'If ctxt����ʱ��.Text = "" Then
'     'ctxt����ʱ��.Text = Format(Now, "yyyy��mm��")
'    'End If
'End Sub



Private Sub ctxt��������_Click()
If ctxtfangshe.Text = "" Then
ctxtfangshe.Text = ctxt��������.Text
Else
ctxtfangshe.Text = ctxtfangshe.Text + "��" + ctxt��������.Text
End If
End Sub





Private Sub ctxt�Ӵ�ʱ��_LostFocus()
    If DateDiff("d", ctxt�Ӵ�ʱ��.Value, Date) < 0 Then
        MsgBox "�Ӵ�ʱ�����"
        ctxt�Ӵ�ʱ��.Value = Date
    End If
End Sub

Private Sub ctxt����_LostFocus()
'    If DateDiff("d", ctxt����.Value, ctxt��ʼ.Value) > 0 Then
'        MsgBox "����ʱ��С����ʼʱ�䣡"
'        ctxt����.Value = Date
'    End If
End Sub

Private Sub ctxt����_LostFocus(Index As Integer)
    If Trim(ctxt����(Index).Text) = "" Then Exit Sub
    If Not IsNumeric(ctxt����(Index).Text) Then
        MsgBox "����ֵ���ԣ�"
        ctxt����(Index).Text = ""
        ctxt����(Index).SetFocus
    Else
        If Int(Val(ctxt����(Index).Text)) > Int(Val(Lab����)) Then
            MsgBox "����ֵ���ԣ�"
            ctxt����(Index).Text = ""
            ctxt����(Index).SetFocus
        End If
    End If
End Sub

Private Sub ctxt����_LostFocus(Index As Integer)
    If Trim(ctxt����(Index).Text) = "" Then Exit Sub
    If Not IsNumeric(ctxt����(Index).Text) Then
        MsgBox "����ֵ���ԣ�"
        ctxt����(Index).Text = ""
        ctxt����(Index).SetFocus
    Else
        If Int(Val(ctxt����(Index).Text)) > Int(Val(Lab����)) Then
            MsgBox "����ֵ���ԣ�"
            ctxt����(Index).Text = ""
            ctxt����(Index).SetFocus
        End If
    End If
End Sub

Private Sub ctxt��ʼ_LostFocus()
'���ڲ����ж� 6-26 by lanchao
'    If DateDiff("d", ctxt��ʼ.Value, Date) < 0 Then
'        MsgBox "��ʼʱ�����"
'        '�޸��ˣ������ 2012-12-4 ��
'        '˵��������ʼʱ������˵�ǰʱ��ʱ��ʾ��ʱ�仹ԭ����ǰʱ��
'        'bug�ţ�0000081
'        ctxt��ʼ.Value = Date
'        '����� 2012-12-4 ��
'    End If
End Sub

Private Sub ctxt���_LostFocus()
    If Trim(ctxt���.Text) = "" Then Exit Sub
    If Not IsNumeric(ctxt���.Text) Then
        MsgBox ("�������������")
        ctxt���.Text = ""
        ctxt���.SetFocus
        Exit Sub
    Else
        If Int(Val(ctxt���.Text)) > 280 Then
            MsgBox ("���ֵ����")
            ctxt���.Text = ""
            ctxt���.SetFocus
            Exit Sub
        ElseIf Int(Val(ctxt���.Text)) < 60 Then
            MsgBox ("���ֵ��С")
            ctxt���.Text = ""
            ctxt���.SetFocus
            Exit Sub
        End If
    End If
    
End Sub

Private Sub ctxt����ѹ_LostFocus()
    If Trim(ctxt����ѹ.Text) = "" Then Exit Sub
    If Not IsNumeric(ctxt����ѹ.Text) Then
        MsgBox ("����ѹ����������")
        ctxt����ѹ.Text = ""
        ctxt����ѹ.SetFocus
        Exit Sub
    Else
        If Int(Val(ctxt����ѹ.Text)) > 230 Then
            MsgBox ("����ѹֵ����")
            ctxt����ѹ.Text = ""
            ctxt����ѹ.SetFocus
            Exit Sub
        ElseIf Int(Val(ctxt����ѹ.Text)) < 60 Then
            MsgBox ("����ѹֵ��С")
            ctxt����ѹ.Text = ""
            ctxt����ѹ.SetFocus
            Exit Sub
        End If
    End If
End Sub

Private Sub ctxt����ѹ_LostFocus()
    If Len(ctxtsysno.Text) = 0 Then Exit Sub    'Ĭ������ѹ�����һ����������ݵġ�
    If Trim(ctxt����ѹ.Text) = "" Then Exit Sub
    If Not IsNumeric(ctxt����ѹ.Text) Then
        MsgBox ("����ѹ����������")
        ctxt����ѹ.Text = ""
        ctxt����ѹ.SetFocus
        Exit Sub
    Else
        If Int(Val(ctxt����ѹ.Text)) > 180 Then
            MsgBox ("����ѹֵ����")
            ctxt����ѹ.Text = ""
            ctxt����ѹ.SetFocus
            Exit Sub
        ElseIf Int(Val(ctxt����ѹ.Text)) < 30 Then
            MsgBox ("����ѹֵ��С")
            ctxt����ѹ.Text = ""
            ctxt����ѹ.SetFocus
            Exit Sub
        End If
    End If
End Sub

Private Sub ctxt����_LostFocus()
    If Trim(ctxt����.Text) = "" Then Exit Sub
    If Not IsNumeric(ctxt����.Text) Then
        MsgBox ("��������������")
        ctxt����.Text = ""
        ctxt����.SetFocus
        Exit Sub
    Else
        If Int(Val(ctxt����.Text)) > 500 Then
            MsgBox ("����ֵ����")
            ctxt����.Text = ""
            ctxt����.SetFocus
            Exit Sub
        ElseIf Int(Val(ctxt����.Text)) < 20 Then
            MsgBox ("����ֵ��С")
            ctxt����.Text = ""
            ctxt����.SetFocus
            Exit Sub
        End If
    End If
End Sub

Private Sub ctxtΣ������_Click()
If ctxtweihai.Text = "" Then
ctxtweihai.Text = ctxtΣ������.Text
Else
ctxtweihai.Text = ctxtweihai.Text + "��" + ctxtΣ������.Text
End If
End Sub

Private Sub ctxt����_LostFocus(Index As Integer)
    If Trim(ctxt����(mIndex).Text) = "" Then Exit Sub
    If Not IsNumeric(ctxt����(mIndex).Text) Then
        MsgBox "����ֵ���ԣ�"
        ctxt����(mIndex).Text = ""
        ctxt����(mIndex).SetFocus
    Else
        If Int(Val(ctxt����(mIndex).Text)) > Int(Val(Lab����)) Then
            MsgBox "����ֵ���ԣ�"
            ctxt����(mIndex).Text = ""
            ctxt����(mIndex).SetFocus
        End If
    End If
End Sub


Private Sub ctxt������_Click()
Combo1.Visible = True

End Sub

Private Sub ctxt�������_LostFocus()
'    If DateDiff("d", ctxt�������.Value, Date) < 0 Then
'        MsgBox "���ʱ�����"
'        ctxt�������.Value = Date
'    End If
End Sub

'�������
Private Sub Form_Load()
    Dim lcol��������ť As New Collection '�������ϵİ�ť��ʼ�����ϡ�
    Dim lcolInfo As Collection
    Dim lobj��� As Object
    Dim i As Integer
    Dim lstrSysno As String
    On Error GoTo errHandler
    
    '��������Ѿ���ʼ�������ٽ��г�ʼ����
    If mblninuse Then Exit Sub
    
    '���ô�������ʹ�õı�־��
    mblninuse = True
    bolenProject = False
    Set mobj��� = CreateObject("ְҵ������.clsMedicalExam")
     
    '��ʼ������ͨ�ö���ÿ�������Ӧһ������ͨ�ö���ʵ�������ɻ��ã��м��мǣ���
    Set mobjGUI = New cls����ͨ�ö���
    
    '���ù�����������Ҫ�ĸ��ְ�ť��
    With lcol��������ť
        '.Add "ȥ����Ա(&D)129"
        '.Add "�����Ա(&R)106"
        .Add "���(&Cl)110"
        .Add "|"
        .Add "�����Ŀ(&T)102"
        .Add "|"
        .Add "����"
        .Add "�޸�"
        .Add "|"
        '.Add "����(&O)111"
        .Add "���沢��ӡ(&P)107"
        .Add "|"
        '.Add "����"
        .Add "�˳�"
    End With
    
    'Ϊ��Ҫͨ������ͨ�ö�����Ƶĸ��ֿؼ�����ʼֵ��
    With mobjGUI
        Set .Form = Me
        Set .c������ = ctbMain
'        Set .c״̬�� = csbMain
    End With
    
    '���ý���ͨ�ö����ṩ�ķ������Խ���ؼ����г�ʼ����
    mobjGUI.subInitialize lcol��������ť, ""
   ' If Right(pubϵͳ���, 1) = "F" Then
    '    lstrSysno = Right(pubϵͳ���, 2)
     '   lstrSysno = Left(lstrSysno, 1)
    'Else
    '    lstrSysno = Right(pubϵͳ���, 1)
    'End If
    '2015-03-30 ��ΰ  6-26 by lanchao "���佡��"<>"8023����"
    Dim resql As Object
    Set resql = dafuncGetData("select �������� From ְҵ�����_�����Ա������Ϣ�� where ϵͳ���='" & pubϵͳ��� & "'")
    If resql("��������") = "��ͨ���" Or resql("��������") = "ְҵ����" Then
        freOrdinary.Visible = True
        freNuclear.Visible = False
        freRadiation.Visible = False
        'ֻ��ְҵ����������ʾ modify by lanchao 2015-9-6
        Label74.Visible = True
        ctxt�Ӵ�.Visible = True
        mIndex = 2
     ElseIf resql("��������") = "8023����" Or resql("��������") = "��˲���" Or resql("��������") = "��˲���YK" Then   '������˲���YK  2015-11-24 by Ĳ��
'    ElseIf resql("��������") = "8023����" Or resql("��������") = "��˲���" Then
        freNuclear.Visible = True
        freOrdinary.Visible = False
        freRadiation.Visible = False
        'ֻ��ְҵ����������ʾ modify by lanchao 2015-9-6
        Label74.Visible = False
        ctxt�Ӵ�.Visible = False
        mIndex = 1
    ElseIf resql("��������") = "���佡��" Then
        freRadiation.Visible = True
        freOrdinary.Visible = False
        freNuclear.Visible = False
        comb����.Visible = True         'ֻ�з��佡������ʾ����  2015-11-30 by Ĳ��
        Lab���(5).Visible = True
        'ֻ��ְҵ����������ʾ modify by lanchao 2015-9-6
        Label74.Visible = False
        ctxt�Ӵ�.Visible = False
        mIndex = 0
    End If
    Label�������.Caption = resql("��������")
    
        
    
    '����ʱ����ʾ��һ������
    SSTab1.Tab = 0
    
    '���ػ����Ͽ�
    'Ccmb���.Clear
    'Ccmb���.AddItem ""
    'Set lobj��� = CreateObject("ְҵ��ʷ¼��.Clslifehstregt")
   '' Set lcolInfo = lobj���.����Ԫ�ؼ�
    'For i = 1 To lcolInfo.Count
        'Ccmb���.AddItem lcolInfo(i), i
        'Ccmb���.ItemData(Ccmb���.NewIndex) = i
    'Next
    'Ccmb���.Text = Ccmb���.List(0)
    
    '2012-04-14 �ڵ�� ��
    '���޸ġ���ťû�й��ܣ�������
    ctbMain.Buttons(6).Visible = False
    '2012-04-14 �ڵ�� ��
    
    '2012-06-13 �ڵ�� ��
    'ʡ������Ҫ��ȥ�������Ŀ�����ʹ�ӡ����
    ctbMain.Buttons(3).Visible = False
    ctbMain.Buttons(4).Visible = False
    ctbMain.Buttons(8).Visible = False
    ctbMain.Buttons(9).Visible = False
    '2012-06-13 �ڵ�� ��
    
    '2012-06-13 �ڵ�� ��
    '��ʼ�����һ��������沿��
    subInit���һ�����
    '2012-06-13 �ڵ�� ��
    
    Set lobj��� = pobjDict.FetchEx("�����ֵ�")
    Ccmb���(mIndex).Clear
    Ccmb���(mIndex).AddItem ""
    For i = 1 To lobj���.RecordCount
        Ccmb���(mIndex).AddItem lobj���("����")
        Ccmb���(mIndex).ItemData(Ccmb���(mIndex).NewIndex) = lobj���("���")
        lobj���.MoveNext
    Next
    Ccmb���(mIndex).ListIndex = 0
    
    'ְҵʷGRID
    Set mcolIndex = New Collection
    For i = 0 To cgrdְҵʷ.Cols - 1
        mcolIndex.Add i, cgrdְҵʷ.TextMatrix(0, i)
    Next
    '��ʼ������ʱ�����乤����Ϣ¼���disable modify by lanchao 2015-8-17
    If Chk����.Value = 1 Then
        Frame11.Enabled = True
    Else
        Frame11.Enabled = False
    End If
    '���û���ʷ���治����
'    If Not mIndex = 2 Then
'        ctxtmatehelh(mIndex).Enabled = False
'        ctxtmarrydate(mIndex).Enabled = False
'        ctxtmatejob(mIndex).Enabled = False
'        ctxtmateradioac(mIndex).Enabled = False
'    End If
'    Frame3(mIndex).Enabled = False
    '������Ϣ�����ݿ�
    Set lobjInDtBase = CreateObject("ְҵ��ʷ¼��.clsCareerhstregt")
    'ְҵ��ʷGRID
    Set mcolIndexwkdis = New Collection
    For i = 0 To cgrd��ʷ.Cols - 1
        mcolIndexwkdis.Add i, cgrd��ʷ.TextMatrix(0, i)
    Next
    '�Ծ�֢״ GRID
    Set mcolindexzz = New Collection
    For i = 0 To cgrd֢״.Cols - 1
        mcolindexzz.Add i, cgrd֢״.TextMatrix(0, i)
    Next
    If ���ʼǺ� = 1 Then
        'ctxtsysno.SetFocus
        ctxtsysno.Text = pubϵͳ���
        'Ccmb���.SetFocus
    End If
    
    
'    ���ʼǺ� = 0
'      MsgBox "�������"   '2016-3-1 by Ĳ��
    Timer1.Enabled = True
    sub�����ն�
    Exit Sub
errHandler:
    sfsub������ "ְҵ��ʷ¼��", "frmcareerhstregt", "form_load", Err.Number, Err.Description, True
    Exit Sub
    Resume
End Sub


Public Sub sub��ѯ�����()
Dim i As Integer
 '���¶���һ���������������䲡״ѯ�ʣ�ԭ�����õ�ְҵʷ�Ķ������mcolIndex�������������mcolIndex����xmcolIndex  2015-12-11 by Ĳ��
Dim xmcolIndex As Object
Dim lobjRec As Object
        dasubSetQueryTimeout 600
        Dim lstrsql As String
        lstrsql = "select ϵͳ���,֢״,�̶�,����ʱ�� from ְҵ�����_�Ծ�֢״�� where ϵͳ���='" & ctxtsysno.Text & "'"
        
        Set lobjRec = dafuncGetData(lstrsql)
        cgrdzzxw.Rows = 1
        
        If Not lobjRec.EOF Then
            With cgrdzzxw
                Set .DataSource = lobjRec
                If cgrdzzxw.Rows > 1 Then
                    Set xmcolIndex = New Collection
                    For i = 0 To cgrdzzxw.Cols - 1
                        xmcolIndex.Add i, cgrdzzxw.TextMatrix(0, i)
                    Next
                End If
              '  clblInfo = .Rows - 1
                .Col = 0
'                .Sort = flexSortGenericDescending
                .AutoSize 0, .Cols - 1, 0, 0
                .ExplorerBar = flexExSort
'                .DataMode = flexDMFree
             '   clblInfo = .Rows - 1
            End With
            
            Exit Sub
        Else
            cgrdzzxw.Rows = 1
          '  clblInfo = cgrdzzxw.Rows - 1
            Exit Sub
        End If

End Sub


'����ȡ��
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    '���ñ�־pblnInUse��
    mblninuse = False
    '�ͷ�ģ�鼶����
    Set mobjGUI = Nothing
    Set lobjInDtBase = Nothing
    Unload frmCareerHstRegt
End Sub




'�����������ť����Ӧ����
Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    Dim i As Long
    Dim lobjRec As Object, lobj������� As Object
    Dim str������� As String
    Dim totalPay As Double
    Dim bol��� As Boolean
    Dim lcolԭ�����Ŀ As Collection
    Dim lcol��� As Collection
    Dim lstr�շ����� As String
    'Set lobjRec = CreateObject("ְҵ��ʷ¼��.clscareerhstregt")
    Dim lstrError As String
    On Error GoTo errHandler
    '2012-05-22 �ڵ�� ��
    '�����á�û�п��Ա�������ݡ����嵯��
    Cancel = True
    '2012-05-22 �ڵ�� ��
    Select Case Operate
    Case "���"
        subclearps
        subclear
        subokclear
        subokclear��ʷ
        subokclear֢״
        subClear���һ�����
        cgrdְҵʷ.Clear
        cgrd��ʷ.Clear
        cgrd֢״.Clear
        subclearps
        '����������
        'cgrdְҵʷ.Clear
        cgrdְҵʷ.Rows = 1
        'cgrd��ʷ.Clear
        cgrd��ʷ.Rows = 1
        'cgrd֢״.Clear
        cgrd֢״.Rows = 1
        ctxtsysno.SetFocus
    Case "ȥ����Ա"
    Case "�����Ա"
    '2012-06-14 �ڵ�� ��
    'ȡ����ӡ���ܣ�case �����沢��ӡ��ȫ��ע��
'''    Case "���沢��ӡ"
'''
'''        If bolenProject = False Then
'''            MsgBox "��ûȷ�������Ŀ�����������Ŀ��ȷ������ܱ��棡"
'''            Exit Sub
'''        End If
'''        '��ʼ����
'''        dasubBeginTran
'''        '��������ʷ����
'''        With lobjPsLifeHst
'''            .mstr��� = Ccmb���.Text
'''            .mstrmatehelh = Trim(ctxtmatehelh.Text)
'''            .mstrmarrydate = Trim(ctxtmarrydate.Value)
'''            .mstrmatejob = Trim(ctxtmatejob.Text)
'''            .mstrmateradioac = Trim(ctxtmateradioac.Text)
'''            .mstr��λ���� = Trim(ctxt��λ����.Text)
'''            .mstr�д� = Trim(ctxt�д�.Text)
'''            .mstr��� = Trim(ctxt���.Text)
'''            .mstr��� = Trim(ctxt���.Text)
'''            .mstr���� = Trim(ctxt����.Text)
'''            .mstr������Ů = Trim(ctxt������Ů.Text)
'''            .mstr���� = Trim(ctxt����.Text)
'''            .mstr��̥ = Trim(ctxt��̥.Text)
'''            .mstr��̥ = Trim(ctxt��̥.Text)
'''            .mstr��Ů���� = Trim(ctxt��Ů����.Text)
'''            .mstr���в��� = Trim(ctxt���в���.Text)
'''            .mstr���� = ccmb����.Text
'''            .mstr���� = ccmb����.Text
'''            .mstr���� = Trim(ctxt����.Text)
'''            .mstr���� = Trim(ctxt����.Text)
'''            .mstr����ʱ�� = Trim(ctxt����.Text)
'''            .mstr����ʷ = Trim(ctxt����ʷ.Text)
'''            .mstr����ʷ = Trim(ctxt����.Text)
'''            .mstr���� = Trim(ctxt����.Text)
'''            .mstr���� = Trim(ctxt����.Text)
'''            .mstr���� = Trim(ctxt����.Text)
'''            .mstrĩ���¾� = Trim(ctxtĩ���¾�.Text)
'''            .mstrͣ�� = Trim(ctxtͣ��.Text)
'''            If ���ʼǺ� = 1 Then
'''                .subDelLifeHst   'ɾ����������ʷ
'''            End If
'''
'''            '��Ԫ�2012.04.10
'''            lobjInDtBase.ϵͳ��� = Trim(ctxtsysno.Text)
'''            '��Ԫ�2012.04.10
'''
'''            .subSaveLifeHst   '�����������ʷ
'''        End With
'''
'''        '���湤��ʷ
'''        saveworkhst
'''
'''        '������ʷ����
'''        SavePastMedcHst
'''
'''        '�Ծ�֢״����
'''        SaveSymptom
'''
'''        '�޸����״̬
'''        lobjInDtBase.sub�޸����״̬
'''
'''        '���������Ŀ  ְҵ��ʷ¼���
'''        Set lobjRec = CreateObject("ְҵ��ʷ¼��.clscareerhstregt")
'''        lobjRec.ϵͳ��� = Trim(ctxtsysno.Text)
'''        Set lobjRec.col�����Ŀ = mcol�����Ŀ
'''        lobjRec.save�Ż��������Ŀ
'''
'''        lstrError = lobjRec.func�շ�(lstr�շ�����)
'''        If lstrError <> "" And lstrError <> "Cancel" Then
'''            MsgBox lstrError, vbOKOnly + vbExclamation, "ϵͳ��ʾ"
'''        End If
'''        '��������
'''        dasubCommitTran
'''
'''        '��ӡ
''''        Set lcol��� = New Collection
''''        lcol���.Add Trim(ctxtsysno.Text)
''''
''''        Set lobj������� = CreateObject("ְҵ������.clsmedicalexam")
''''        lobj�������.ϵͳ��� = Trim(ctxtsysno.Text)
''''        str������� = lobj�������.�������
''''            '��ӡ
''''            pobjҵ�����.Sub��ӡ���� str������� & "���Ǽǵ�", lcol���, False
''''            'Cancel = True
''''        Set lobj������� = Nothing
'''        subPrint Trim(ctxtsysno.Text)
'''
'''        ���ʼǺ� = 0
'''        Set lobjInDtBase.mobjWorkHst = New Collection
'''        Set lobjInDtBase.mobjPastHst = New Collection
'''        Set lobjInDtBase.mobjSymptom = New Collection
'''        If ChkClear Then
'''            subclearps
'''            subclear
'''            subokclear
'''            subokclear��ʷ
'''            subokclear֢״
'''            cgrdְҵʷ.Rows = 1
'''            cgrd��ʷ.Rows = 1
'''            cgrd֢״.Rows = 1
'''        End If
'''        subclearps
'''        '����������
'''        cgrdְҵʷ.Rows = 1
'''        cgrd��ʷ.Rows = 1
'''        cgrd֢״.Rows = 1
'''        ctxtsysno.SetFocus
'''        Cancel = True
'''
'''        MsgBox "��ӡ�ɹ���"
    '2012-06-14 �ڵ�� ��
    'ȡ�������Ŀ�������ܣ�ȫ��ע�� case "�����Ŀ"
'''    Case "�����Ŀ"
'''        Dim lobj���ģ�� As Object
'''
'''        If ctxtsysno.Text = "" Then
'''            MsgBox "ϵͳ��Ų���Ϊ�գ�"
'''            Exit Sub
'''        End If
'''        '��ȡ���������е������Ŀ��
'''        mobj���.ϵͳ��� = Trim(ctxtsysno.Text)
'''        Set lcolԭ�����Ŀ = mobj���.����.�����Ŀ��("")
'''
'''        '����ѡ����Ŀ��������ԡ�
'''        frmSelectItem.pstr�������� = "��ҵ��Ա�����"
'''        Set frmSelectItem.pcol������Ŀ = lcolԭ�����Ŀ
'''        'Set frmSelectItem.pcol�շ���Ŀ = mcol�շ���Ŀ
'''        '����ѡ����Ŀ���档
'''        frmSelectItem.Show 1
'''        If frmSelectItem.pblnOk Then
'''            '��ȡѡ�еĸ�����Ŀ��
'''            Set mcol�����Ŀ = frmSelectItem.pcol������Ŀ
'''            '��ȡ���õ��շ���Ŀ��
'''            'Set mcol�շ���Ŀ = frmSelectItem.pcol�շ���Ŀ
'''
'''            '��ʾ�շѽ�
'''            Dim ldblTotal As Double
'''            'For i = 1 To mcol�շ���Ŀ.Count
'''            '    ldblTotal = Format(ldblTotal + mcol�շ���Ŀ(i)("����"), "0.00")
'''            'Next
'''            'On Error Resume Next
'''            'If sffunc�жϼ��ϼ�ֵ�Ƿ����(mobj���.����.������Ϣ, "�����") Then
'''                'ciptBase.Box1("�����").Text = ldblTotal
'''                'mobj���.����.Sub�����Ϣֵ "�����", ldblTotal
'''                mobj���.����.Sub�����Ϣֵ "�����", 100
'''            'End If
'''            bolenProject = True '��ȷ�������Ŀ
'''        End If
    Case "����"
    
    Case "����"
        '2012-06-14 �ڵ��
        'case "����" ���֣�������Ŀ��첿��ȫ��ע��
        
'''        '�ж��Ƿ���ȷ�������Ŀ
'''
'''        '���ܣ���¼��Ա�������Ŀ
'''        'ʱ�䣺2012-06-05
'''        '���ߣ�����
'''        Dim lobj�����Ŀ As Object
'''        Set lobj�����Ŀ = CreateObject("ְҵ��ʷ¼��.ClsCareerHstRegt")
'''        Set mcol�����Ŀ = lobj�����Ŀ.func��ȡ�����Ա�������Ŀ(ctxtsysno.Text)
'''
'''        '��ʾ�շѽ�
'''        mobj���.����.Sub�����Ϣֵ "�����", 100
'''        bolenProject = True
'''        'ʱ�䣺2012-06-05
'''
'''        If bolenProject = False Then
'''            MsgBox "��ûȷ�������Ŀ�����������Ŀ��ȷ������ܱ��棡"
'''            Exit Sub
'''        End If
        '��ʼ����
        dasubBeginTran
        '��������ʷ����
        With lobjPsLifeHst
            .mstr��� = Ccmb���(mIndex).Text
            .mstr������Ů = Trim(ctxt������Ů(mIndex).Text)

            If mIndex <> 2 Then
                If Ccmb���(mIndex).Text = "�ѻ�" Or Ccmb���(mIndex).Text = "����" Then
                    '.mstrmarrydate = Trim(ctxtmarrydate(mIndex).Value)
                    .mstrmarrydate = Trim(ctxtmarrydate(mIndex).Text)
                    .mstrmatehelh = Trim(ctxtmatehelh(mIndex).Text)
                    .mstrmatejob = Trim(ctxtmatejob(mIndex).Text)
                    .mstrmateradioac = Trim(ctxtmateradioac(mIndex).Text)
                End If
                If mIndex <> 1 Then
                    .mstr��λ���� = Trim(ctxt��λ����(mIndex).Text)
                End If
                .mstr�д� = Trim(ctxt�д�(mIndex).Text)
                .mstr��� = Trim(ctxt���(mIndex).Text)
                .mstr��̥ = Trim(ctxt��̥(mIndex).Text)
                .mstr��̥ = Trim(ctxt��̥(mIndex).Text)
                .mstr��Ů���� = Trim(ctxt��Ů����(mIndex).Text)
                .mstr���в��� = Trim(ctxt���в���(mIndex).Text)
                .mstrMore = Trim(ctxtMore(mIndex).Text)
            End If
            .mstr�쳣̥ = Trim(ctxt�쳣̥.Text)
            .mstr��� = Trim(ctxt���(mIndex).Text)
            .mstr���� = Trim(ctxt����(mIndex).Text)
            If mIndex <> 1 Then
            '8023��������������������䣬���Էŵ�����һ�� 2015-9-29
'                .mstr������Ů = Trim(ctxt������Ů(mIndex).Text)
'                .mstr���� = Trim(ctxt����(mIndex).Text)
'                .mstr���� = ccmb����(mIndex).Text
'                .mstr���� = ccmb����(mIndex).Text
'                .mstr���� = Trim(ctxt����(mIndex).Text)
'                .mstr���� = Trim(ctxt����(mIndex).Text)
'                .mstr����ʱ�� = Trim(ctxt����(mIndex).Text)
                If mIndex <> 0 Then
                    .mstr����ʷ = Trim(ctxt����ʷ(mIndex).Text)
                End If
            End If
            '�����������Ƴ̶ȣ�����ʱ��  2015-11-10 by Ĳ��
            .mstr���� = ccmb����(mIndex).Text
            .mstr���� = ccmb����(mIndex).Text
            .mstr����ʱ�� = Trim(ctxt����(mIndex).Text)
            
            .mstr������ = Trim(ctxt������(mIndex).Text)
            .mstr������ = Trim(ctxt������(mIndex).Text)
            .mstr���� = Trim(ctxt����(mIndex).Text)
            .mstr���� = Trim(ctxt����(mIndex).Text)
            .mstr���� = Trim(ctxt����(mIndex).Text)
            
            .mstr����ʷ = Trim(ctxt����.Text)
            .mstr���� = Trim(ctxt����.Text)
            .mstr���� = Trim(ctxt����.Text)
            .mstr���� = Trim(ctxt����.Text)
            .mstrĩ���¾� = Trim(ctxtĩ���¾�.Text)
            .mstrͣ�� = Trim(ctxtͣ��.Text)
            .mstrOther = Trim(ctxtOther.Text)
            '���佡����촦��Ů���������к���Ů���������� 2015-7-1 by lanchao
            If mIndex = 0 Then
            .mstr����Ů�� = Trim(ctxt����Ů��(mIndex).Text)
            .mstr�к��������� = Trim(ctxt�к���������(mIndex).Text)
            .mstrŮ���������� = Trim(ctxtŮ����������(mIndex).Text)
            End If
              '8023������촦���к�Ů�������ͳ������� 2015 - 9 - 28 by Ĳ��
             If mIndex = 1 Then
            .mstr����Ů�� = Trim(ctxt����Ů��(mIndex).Text)
            .mstr�к��������� = Trim(ctxt�к���������(mIndex).Text)
            .mstrŮ���������� = Trim(ctxtŮ����������(mIndex).Text)
            End If
            
            'ְҵ�������ӳ�����    2015-11-10  by Ĳ��
            If mIndex = 2 Then
            dafuncGetData ("update ְҵ�����_�����Ա������Ϣ�� set ������='" & ctxt������(mIndex).Text & "' where  ϵͳ���='" & Trim(ctxtsysno.Text) & "'")
            End If
            
            If ���ʼǺ� = 1 Then
                .subDelLifeHst   'ɾ����������ʷ
            End If
            
            '��Ԫ�2012.04.10
            lobjInDtBase.ϵͳ��� = Trim(ctxtsysno.Text)
            '��Ԫ�2012.04.10
            
            .subSaveLifeHst   '�����������ʷ
        End With
        
        '2012-06-14 �ڵ�� ��
        subSave���һ����� ctxtsysno.Text
        '2012-06-14 �ڵ�� ��
    
        '���湤��ʷ
        saveworkhst
        
        '������ʷ����
        SavePastMedcHst
        
        '�Ծ�֢״���� 2015-8-20 by lanchao ���屣���ʱ�򲻲���
        'SaveSymptom
        
        '�޸����״̬
        lobjInDtBase.sub�޸����״̬
        
'''        '���������Ŀ  ְҵ��ʷ¼���
'''        Set lobjRec = CreateObject("ְҵ��ʷ¼��.clscareerhstregt")
'''        lobjRec.ϵͳ��� = Trim(ctxtsysno.Text)
'''        Set lobjRec.col�����Ŀ = mcol�����Ŀ
'''        lobjRec.save�Ż��������Ŀ
'''
'''
'''        lstrError = lobjRec.func�շ�(lstr�շ�����)
'''        If lstrError <> "" And lstrError <> "Cancel" Then
'''            MsgBox lstrError, vbOKOnly + vbExclamation, "ϵͳ��ʾ"
'''        End If
        
        '2012-06-15 �ڵ�� ��
        '�������״̬����"δ¼���ܼ��߸�����Ϣ"2����Ϊ"�����"3
        pobjҵ�����.funcд�뵥�˵�ǰ���״̬ ctxtsysno.Text, 3
        '2012-06-15 �ڵ�� ��
        
        '2012-07-04 �ڵ�� ��
        '���¸������״̬�����ܵ����״̬
        pobjҵ�����.sub�޸Ľ��¼��״̬ ctxtsysno.Text, "13", "2"  '13���������Ϣ¼����ң��ұ�ʾ�ڸ������״̬�ַ�����λ�ã��ǿ��ұ�ţ�
        pobjҵ�����.sub���¼���޸����״̬ Trim(ctxtsysno.Text), "4"
        '2012-07-04 �ڵ�� ��
        
        '2012-08-22 �ڵ�� ��
        '����ÿ��������ۣ����꣩
        pobjҵ�����.sub������д������ Trim(ctxtsysno.Text), "�ܼ��߸�����Ϣ¼���", "����", um�û����
        '2012-08-22 �ڵ�� ��
        
'        '����շŵ�������ִ��   2016-2-24 by Ĳ��
'        If ChkClear Then
'            subclearps
'            subclear
'            subokclear
'            subokclear��ʷ
'            subokclear֢״
'            subClear���һ�����
'            cgrdְҵʷ.Rows = 1
'            cgrd��ʷ.Rows = 1
'            cgrd֢״.Rows = 1
'        End If
        
        
        '��������
        dasubCommitTran
        If MSComm1.PortOpen = True Then
            MSComm1.PortOpen = False    '�����ص��ն�����
        End If
        MsgBox "����ɹ���"
        ���ʼǺ� = 0
        Set lobjInDtBase.mobjWorkHst = New Collection
        Set lobjInDtBase.mobjPastHst = New Collection
        Set lobjInDtBase.mobjSymptom = New Collection
        
'        '����ɹ�֮����ս��� 2016-1-7 by Ĳ�� ��  ��ʱ�ֲ�����
'        If ChkClear Then
'            subclearps
'            subclear
'            subokclear
'            subokclear��ʷ
'            subokclear֢״
'            subClear���һ�����
'            cgrdְҵʷ.Rows = 1
'            cgrd��ʷ.Rows = 1
'            cgrd֢״.Rows = 1
'        End If
'    '����ɹ�֮����ս��� 2016-1-7 by Ĳ�� ��   ��ʱ�ֲ�����

'        MsgBox "�������" '2016-3-1 by Ĳ��
        '�رյ�ǰ���� 2015-8-17 modify by lanchao
        Unload frmCareerHstRegt
        '�޸��ˣ������ 2012-12-11 ��
        '˵����ѡ�񱣴�����
        'bug�ţ�0000059
'        subclearps
     
        
        '����������
'        cgrdְҵʷ.Rows = 1
'        cgrd��ʷ.Rows = 1
'        cgrd֢״.Rows = 1
'        ctxtsysno.SetFocus
'        Cancel = True
   '�޸��ˣ������ 2012-12-11 ��
    Case "�˳�"
        If Len(ctxtsysno.Text) > 0 Then
            i = MsgBox("�Ƿ񱣴浱ǰ¼����Ϣ��", vbYesNo, "ϵͳ��ʾ")
            If i = vbYes Then mobjGUI_BeforeOperate "����", True
        End If
        '2012-05-22 �ڵ�� ��
        '����м����ʧ�󣬿��ܵ����˳���ť�����á��������У��ֲ��������
        Unload frmCareerHstRegt
        '2012-05-22 �ڵ�� ��
    End Select
    
    Exit Sub
errHandler:
    sfsub������ "ְҵ��ʷ¼��", "frmcareerhstregt", "mobjGUI_BeforeOperate", Err.Number, Err.Description, True
    MousePointer = 0
'    csbMain.Panels(1) = ""
    Exit Sub
    Resume
End Sub

'����������ݱ��������ݿ�  �Ծ�֢״
Private Sub SaveSymptom()
    Dim i As Integer
    Dim lobjdetail As Object
    On Error GoTo errHandler
    For i = 1 To cgrd֢״.Rows - 1
    '����֢״����
    Set lobjdetail = CreateObject("ְҵ��ʷ¼��.clssymptomdetl")
    '��Ԫ�2012.04.10
    'Ϊϵͳ��Ÿ�ֵ
    lobjInDtBase.ϵͳ��� = Trim(ctxtsysno.Text)
    '��Ԫ�2012.04.10
        lobjdetail.mstr��� = cgrd֢״.Cell(flexcpText, i, mcolindexzz("���"))
        lobjdetail.mstr֢״ = cgrd֢״.Cell(flexcpText, i, mcolindexzz("֢״"))
        lobjdetail.mstr����ʱ�� = cgrd֢״.Cell(flexcpText, i, mcolindexzz("����ʱ��"))
        lobjdetail.mstr�̶� = cgrd֢״.Cell(flexcpText, i, mcolindexzz("�̶�"))
        lobjInDtBase.mobjSymptom.Add lobjdetail
    Next
    If ���ʼǺ� = 1 Then
        lobjInDtBase.subDelSymptom
    End If
    lobjInDtBase.SubSaveSymptom
    Exit Sub
errHandler:
    sfsub������ "ְҵ��ʷ¼��", "frmcareerhstregt", "savesymptom", Err.Number, Err.Description, True
End Sub

'����������ݱ��������ݿ�  ������ʷ
Private Sub SavePastMedcHst()
    Dim i As Integer
    Dim lobjdetail As Object
    On Error GoTo errHandler
    For i = 1 To cgrd��ʷ.Rows - 1
    '������ʷ����
    Set lobjdetail = CreateObject("ְҵ��ʷ¼��.clsPastMedcHstdetl")
    '��Ԫ�2012.04.10
    'Ϊϵͳ��Ÿ�ֵ
    lobjInDtBase.ϵͳ��� = Trim(ctxtsysno.Text)
    '��Ԫ�2012.04.10
        lobjdetail.mstr��� = cgrd��ʷ.Cell(flexcpText, i, mcolIndexwkdis("���"))
        lobjdetail.mstr�������� = cgrd��ʷ.Cell(flexcpText, i, mcolIndexwkdis("��������"))
        lobjdetail.mstr��ϵ�λ = cgrd��ʷ.Cell(flexcpText, i, mcolIndexwkdis("��ϵ�λ"))
        lobjdetail.mstr������� = cgrd��ʷ.Cell(flexcpText, i, mcolIndexwkdis("�������"))
        lobjdetail.mstrת�� = cgrd��ʷ.Cell(flexcpText, i, mcolIndexwkdis("ת��"))
        lobjdetail.mstr���ƾ��� = cgrd��ʷ.Cell(flexcpText, i, mcolIndexwkdis("���ƾ���"))
        lobjInDtBase.mobjPastHst.Add lobjdetail
    Next
    If ���ʼǺ� = 1 Then
        lobjInDtBase.subDelPastMedcHst
    End If
    lobjInDtBase.SubSavePastMedcHst
    Exit Sub
errHandler:
    sfsub������ "ְҵ��ʷ¼��", "frmcareerhstregt", "savepastmedchst", Err.Number, Err.Description, True
End Sub

'����������ݱ��������ݿ�  ְҵʷ
Private Sub saveworkhst()
    Dim i As Integer
    Dim lobjdetail As Object
    On Error GoTo errHandler
    
    Set mcolIndex = New Collection
    For i = 0 To cgrdְҵʷ.Cols - 1
        mcolIndex.Add i, cgrdְҵʷ.TextMatrix(0, i)
    Next
    For i = 1 To cgrdְҵʷ.Rows - 1
        '����ְҵʷ����
        Set lobjdetail = CreateObject("ְҵ��ʷ¼��.clscareerhstDetl")
        lobjInDtBase.ϵͳ��� = Trim(ctxtsysno.Text)
    
        lobjdetail.mstr��� = cgrdְҵʷ.Cell(flexcpText, i, mcolIndex("���"))
        lobjdetail.mstr��λ = cgrdְҵʷ.Cell(flexcpText, i, mcolIndex("������λ"))
        lobjdetail.mstr���� = cgrdְҵʷ.Cell(flexcpText, i, mcolIndex("����"))
        lobjdetail.mstr���� = cgrdְҵʷ.Cell(flexcpText, i, mcolIndex("����"))
        lobjdetail.mstrΣ������ = cgrdְҵʷ.Cell(flexcpText, i, mcolIndex("Σ������"))
        lobjdetail.mstr�Ӵ�ʱ�� = cgrdְҵʷ.Cell(flexcpText, i, mcolIndex("�Ӵ�ʱ��"))
        lobjdetail.mstr��ʩ = cgrdְҵʷ.Cell(flexcpText, i, mcolIndex("������ʩ"))
        lobjdetail.mstr��ע = cgrdְҵʷ.Cell(flexcpText, i, mcolIndex("��ע"))
        lobjdetail.mstr��ʼʱ�� = cgrdְҵʷ.Cell(flexcpText, i, mcolIndex("��ʼʱ��"))
        lobjdetail.mstr����ʱ�� = cgrdְҵʷ.Cell(flexcpText, i, mcolIndex("����ʱ��"))
       
        lobjdetail.mstr�������� = cgrdְҵʷ.Cell(flexcpText, i, mcolIndex("��������"))
        lobjdetail.mstr������ = cgrdְҵʷ.Cell(flexcpText, i, mcolIndex("ÿ�չ�����"))
        lobjdetail.mstr������ = cgrdְҵʷ.Cell(flexcpText, i, mcolIndex("�ۻ�������"))
        lobjdetail.mstr��������ʷ = cgrdְҵʷ.Cell(flexcpText, i, mcolIndex("��������ʷ"))
        lobjdetail.mstr�Ƿ������ = cgrdְҵʷ.Cell(flexcpText, i, mcolIndex("�Ƿ������"))
        lobjInDtBase.mobjWorkHst.Add lobjdetail
    Next
    If ���ʼǺ� = 1 Then
        lobjInDtBase.subDelWorkHst
    End If
    lobjInDtBase.subSaveWorkHst
    Exit Sub
errHandler:
    sfsub������ "ְҵ��ʷ¼��", "frmcareerhstregt", "saveworkhst", Err.Number, Err.Description, True
End Sub

'��ո�����Ϣ
Private Sub subclearps()
    ctxtsysno.Text = ""
    Lab����.Caption = ""
    Lab�Ա�.Caption = ""
    Lab����.Caption = ""
    lab��λ.Caption = ""
    Lab�ֹ���.Caption = ""
    Lab��ְ��.Caption = ""
    LabelΣ������.Caption = ""
    Set Picture2.Picture = Nothing
    
    '2012-04-14 �ڵ�� ��
    '��ա���ʼ��������ȫ�ֱ���
    mintrow = 0
    jmintrow = 0
    Set mobj��� = CreateObject("ְҵ������.clsMedicalExam")
    Set lobjInDtBase = CreateObject("ְҵ��ʷ¼��.clsCareerhstregt")
    Set mcol�����Ŀ = frmSelectItem.pcol������Ŀ
    Set mcolindexzz = New Collection
    Set mcolIndexwkdis = New Collection
    Set mcolIndex = New Collection
    '2012-04-14 �ڵ�� ��
     
End Sub
'��ս�����Ϣ  ����ʷ
Private Sub subclear()
    Ccmb���(mIndex).Text = ""
    '��������ڣ��к�Ů�������������  2016-3-2 by Ĳ��
    ctxtmarrydate(mIndex).Text = ""
    ctxt������Ů(mIndex).Text = "0"
    ctxt����Ů��(mIndex).Text = "0"
    ctxt�к���������(mIndex).Text = ""
    ctxtŮ����������(mIndex).Text = ""
    ctxt����(mIndex).Text = ""
    ctxt����(mIndex).Text = ""
    
    If mIndex <> 2 Then
        'ctxtmatehelh(mIndex).Text = ""
        'modify by lanchao 2015-7.20,��ż����Ĭ��Ϊ�����������
        ctxtmatejob(mIndex).Text = ""
        ctxtmateradioac(mIndex).Text = ""
        If mIndex <> 1 Then
            ctxt��λ����(mIndex).Text = "0"
        End If
        ctxt�д�(mIndex).Text = "0"
         'modify by lanchao 2015-7.20,�����������գ���ʾΪX��XŮ
        If mIndex <> 1 Then
           ctxt���(mIndex).Text = "0"
        End If
        ctxt��̥(mIndex).Text = "0"
        ctxt��̥(mIndex).Text = "0"
        ctxt��Ů����(mIndex).Text = ""
        ctxt���в���(mIndex).Text = ""
        ctxtMore(mIndex).Text = ""
    End If
    ctxt���(mIndex).Text = "0"
    ctxt����(mIndex).Text = "0"
    If mIndex <> 1 Then
        ctxt������Ů(mIndex).Text = "0"
        ctxt����(mIndex).Text = "0"
'        ccmb����(mIndex).Text = ""
'        ccmb����(mIndex).Text = ""
        ctxt����(mIndex).Text = ""
        ctxt����(mIndex).Text = ""
        ctxt����(mIndex).Text = ""
        If mIndex <> 0 Then
            ctxt����ʷ(mIndex).Text = ""
        End If
    End If
    
    ccmb����(mIndex).Text = ""
    ccmb����(mIndex).Text = ""
        
    ctxt������(mIndex).Text = ""
    ctxt������(mIndex).Text = ""
    ctxt�쳣̥.Text = "0"
    ctxt����.Text = ""
    ctxt����.Text = ""
    ctxt����.Text = ""
    ctxt����.Text = ""
    ctxtĩ���¾�.Text = ""
    ctxtͣ��.Text = ""
End Sub

'���  ְҵʷ
Private Sub subokclear()
     ctxt��λ.Text = ""
     ctxt����.Text = ""
     ctxt����.Text = ""
     ctxt��ע.Text = ""

     ctxtΣ������.Text = ""
     ctxt����.Text = ""
     ctxt��������.Text = ""
     ctxtfangshe.Text = ""   '�������������  2015-12-2 by Ĳ��
     ctxt������.Text = ""
     ctxt������.Text = ""
     ctxt��������.Text = ""
     ctxt��ʼ.Text = "����"
     ctxt����.Text = "����"
     ctxt�Ӵ�.Text = ""        '��Ҫ�Ӵ�ʱ��  2015-12-11 by Ĳ��
     ctxtweihai.Text = ""
          
     '�޸��ˣ������ 2012-12-28   ��
     '˵������ԭʱ��
     'bug�ţ�0000122 2015-6-26 ����Ҫ���ڸ�ʽ by lanchao
     'ctxt��ʼ.Value = Date
     'ctxt�Ӵ�ʱ��.Value = Date
     'ctxt����.Value = Date
     '�޸��ˣ������ 2012-12-28   ��
End Sub

'���  ְҵ��ʷ
Private Sub subokclear��ʷ()
    ctxt��������.Text = ""
    ctxt��ϵ�λ.Text = ""
    ctxtת��.Text = ""
    ctxt���ƾ���.Text = ""
    '�޸ģ������ 2012-12-7 ��
    '�ָ�Ĭ�����ʱ��
    'Bug�ţ�0000057
    'ctxt�������.Value = Date
    '�޸ģ������ 2012-12-7 ��
End Sub

'2012.12.11 ����
'˵�����������ʷ  ��
Public Sub subclear����ʷ()
'    ctxt��λ����(Index).Text = ""
    ctxt�д�(mIndex).Text = "0"
    ctxt���(mIndex).Text = "0"
    ctxt���(mIndex).Text = "0"
    ctxt����(mIndex).Text = "0"
    ctxt����(mIndex).Text = "0"
    
    'δ�齫��Ů���������ڶ�����
    ctxt������Ů(mIndex).Text = "0"
    If mIndex <> 2 Then
    ctxt����Ů��(mIndex).Text = "0"
    ctxt�к���������(mIndex).Text = ""
    ctxtŮ����������(mIndex).Text = ""
    End If
    '8023�����Ѿ��������⼸������  2015-9-29
'    If mIndex <> 1 Then
''        ctxt������Ů(mIndex).Text = "0"
''        ctxt����(mIndex).Text = "0"
'    End If
    ctxt��̥(mIndex).Text = "0"
    ctxt��̥(mIndex).Text = "0"
    ctxt���в���(mIndex).Text = ""
    ctxt��Ů����(mIndex).Text = ""
End Sub

'���  �Ծ�֢״
Private Sub subokclear֢״()
'    ctxt֢״.Text = ""
'    ctxt�̶�.Text = ""
    ccmb����.Text = ""
    ctxt�̶�.Text = ""
'    ctxt����ʱ��.Text = "����"
End Sub

'����ܼ��߸�����Ϣ����  2016-3-2 by Ĳ��
Private Sub gatherclear()
    subclear
    subokclear
    subokclear��ʷ
    subokclear֢״
    subClear���һ�����
    cgrdְҵʷ.Rows = 1
    cgrd��ʷ.Rows = 1
    cgrd֢״.Rows = 1
End Sub

'�޸� ����ʷ
Private Sub sub�޸�����ʷ(ByVal lobjlife As Object)
    On Error GoTo errHandler
    '2012-04-14 �ڵ�� ��
    '��û����Ϣʱ��ֱ���˳�
    If lobjlife.RecordCount = 0 Then Exit Sub
    '2012-04-14 �ڵ�� ��
        Ccmb���(mIndex).Text = IIf(IsNull(lobjlife!�Ƿ���), "", lobjlife!�Ƿ���)
        Ccmb���_Click (mIndex)
        ctxt������Ů(mIndex).Text = IIf(IsNull(lobjlife!������Ů��Ŀ), "", lobjlife!������Ů��Ŀ)
        If mIndex <> 2 Then
            ctxtmatehelh(mIndex).Text = IIf(IsNull(lobjlife!��ż����״��), "����", lobjlife!��ż����״��)
            '�����ڸ�ʽ�޸�Ϊ�ı���  2015-6-26 by lanchao
            'ctxtmarrydate(mIndex).Value = Format(IIf(lobjlife!������� = "", Date, lobjlife!�������), "yyyy/mm/dd")
            ctxtmarrydate(mIndex).Text = IIf(IsNull(lobjlife!�������), "", lobjlife!�������)
            ctxtmatejob(mIndex).Text = IIf(IsNull(lobjlife!��żְҵ), "", lobjlife!��żְҵ)
            ctxtmateradioac(mIndex).Text = IIf(IsNull(lobjlife!��ż�Ӵ�����), "", lobjlife!��ż�Ӵ�����)
            If mIndex <> 1 Then
                ctxt��λ����(mIndex).Text = IIf(IsNull(lobjlife!��λ����), "", lobjlife!��λ����)
            End If
            ctxt�д�(mIndex).Text = IIf(IsNull(lobjlife!�д�), "", lobjlife!�д�)
            ctxt���(mIndex).Text = IIf(IsNull(lobjlife!���), "", lobjlife!���)
            ctxt��̥(mIndex).Text = IIf(IsNull(lobjlife!��̥), "", lobjlife!��̥)
            ctxt��̥(mIndex).Text = IIf(IsNull(lobjlife!��̥), "", lobjlife!��̥)
            ctxt��Ů����(mIndex).Text = IIf(IsNull(lobjlife!��Ů����״��), "", lobjlife!��Ů����״��)
            ctxt���в���(mIndex).Text = IIf(IsNull(lobjlife!���в���ԭ��), "", lobjlife!���в���ԭ��)
            ctxtMore(mIndex).Text = IIf(IsNull(lobjlife!�������), "", lobjlife!�������)
        End If
        ctxt�쳣̥.Text = IIf(IsNull(lobjlife!�쳣̥), "", lobjlife!�쳣̥)
        ctxt���(mIndex).Text = IIf(IsNull(lobjlife!���), "", lobjlife!���)
        ctxt����(mIndex).Text = IIf(IsNull(lobjlife!����), "", lobjlife!����)
        If mIndex <> 1 Then
            ctxt������Ů(mIndex).Text = IIf(IsNull(lobjlife!������Ů��Ŀ), "", lobjlife!������Ů��Ŀ)
'            ctxt����(mIndex).Text = IIf(IsNull(lobjlife!��Ȼ����), "", lobjlife!��Ȼ����)

'            ccmb����(mIndex).Text = IIf(IsNull(lobjlife!���Ƴ̶�), "", lobjlife!���Ƴ̶�)
'            ccmb����(mIndex).Text = IIf(IsNull(lobjlife!���̶̳�), "", lobjlife!���̶̳�)

'            ctxt����(mIndex).Text = IIf(IsNull(lobjlife!����), "", lobjlife!����)
'            ctxt����(mIndex).Text = IIf(IsNull(lobjlife!����), "", lobjlife!����)
'            ctxt����(mIndex).Text = IIf(IsNull(lobjlife!����ʱ��), "", lobjlife!����ʱ��)
            If mIndex <> 0 Then
                ctxt����ʷ(mIndex).Text = IIf(IsNull(lobjlife!����ʷ), "", lobjlife!����ʷ)
            End If
        End If
        ccmb����(mIndex).Text = IIf(IsNull(lobjlife!���Ƴ̶�), "", lobjlife!���Ƴ̶�)
        ccmb����(mIndex).Text = IIf(IsNull(lobjlife!���̶̳�), "", lobjlife!���̶̳�)
        ctxt����(mIndex).Text = IIf(IsNull(lobjlife!����ʱ��), "", lobjlife!����ʱ��)
        
        ctxt����(mIndex).Text = IIf(IsNull(lobjlife!��Ȼ����), "", lobjlife!��Ȼ����)
        ctxt������(mIndex).Text = IIf(IsNull(lobjlife!������), "", lobjlife!������)
        ctxt������(mIndex).Text = IIf(IsNull(lobjlife!������), "", lobjlife!������)
        ctxt����(mIndex).Text = IIf(IsNull(lobjlife!����), "", lobjlife!����)
        ctxt����(mIndex).Text = IIf(IsNull(lobjlife!����), "", lobjlife!����)
        ctxt����.Text = IIf(IsNull(lobjlife!����ʷ), "", lobjlife!����ʷ)
        ctxt����.Text = IIf(IsNull(lobjlife!����), "", lobjlife!����)
        ctxt����.Text = IIf(IsNull(lobjlife!����), "", lobjlife!����)
        ctxt����.Text = IIf(IsNull(lobjlife!����), "", lobjlife!����)
        ctxtĩ���¾�.Text = IIf(IsNull(lobjlife!ĩ���¾�), "", lobjlife!ĩ���¾�)
        ctxtͣ��.Text = IIf(IsNull(lobjlife!ͣ������), "", lobjlife!ͣ������)
        ctxtOther.Text = IIf(IsNull(lobjlife!����), "", lobjlife!����)
        '������佡������к���Ů�������������� 2015-7-1 by lanchao
         If mIndex = 0 Then
            ctxt����Ů��(mIndex).Text = IIf(IsNull(lobjlife!����Ů��), "", lobjlife!����Ů��)
            ctxt�к���������(mIndex).Text = IIf(IsNull(lobjlife!�к���������), "", lobjlife!�к���������)
            ctxtŮ����������(mIndex).Text = IIf(IsNull(lobjlife!Ů����������), "", lobjlife!Ů����������)
        End If
        '����8023��������к���Ů�������������� 2015-9-28
         If mIndex = 1 Then
'            ctxt������Ů(mIndex).Text = IIf(IsNull(lobjlife!������Ů), "", lobjlife!������Ů)
            ctxt����Ů��(mIndex).Text = IIf(IsNull(lobjlife!����Ů��), "", lobjlife!����Ů��)
            ctxt�к���������(mIndex).Text = IIf(IsNull(lobjlife!�к���������), "", lobjlife!�к���������)
            ctxtŮ����������(mIndex).Text = IIf(IsNull(lobjlife!Ů����������), "", lobjlife!Ů����������)
        End If
        
        '�޸�ʱ��ʾ�ϴ�д��ĳ�������Ϣ  2015-11-10 by Ĳ��
        If mIndex = 2 Then
        Dim csql As Object
        Set csql = dafuncGetData("select * From ְҵ�����_�����Ա������Ϣ�� where ϵͳ���='" & Trim(ctxtsysno.Text) & "'")
        ctxt������(mIndex).Text = csql("������")
        End If
        Exit Sub
errHandler:
    sfsub������ "ְҵ��ʷ¼��", "frmcareerhstregt", "sub�޸�����ʷ", Err.Number, Err.Description, True
End Sub


Private Sub MSComm1_OnComm()
'Dim S() As Byte
'    Dim SS(1024) As Byte
'    Static N As Long
'    Static T As Variant
'     Dim intInputLen, i As Integer
'Dim instring As String
'    If (MSComm1.CommEvent = comEvReceive) Then
'        S = MSComm1.Input                      'ֻҪ�����ݾ��ս���������ֻ��һ��
'
'        T = Timer
'        For i = 0 To UBound(S)
'        'һ�����ݰ����ܲ������ɸ�oncomm�¼�
'        instring = StrConv(S, vbUnicode)
'                         Text��ʱ.Text = Text��ʱ.Text & instring
'            SS(N + i) = S(i)                 '�������ݰ�������SS()
'            N = N + UBound(S)
'        Next i
'       ' MSComm1.InBufferCount = 0
'    End If

    If MSComm1.InBufferCount Then
        ' ͨѶ���м��������ϵĻ�, ���ȡ����
           Dim InStringB() As Byte
           Dim instring As String
          InStringB = MSComm1.Input
          instring = StrConv(InStringB, vbUnicode)
          Text��ʱ.Text = Text��ʱ.Text & instring
          InStringB = ""
          If Len(Text��ʱ.Text) < 13 Then
                MSComm1_OnComm
          End If
    End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
'ȥ���������ݲ���  2015-12-11 by Ĳ��
'If PreviousTab = 5 Then
'sub��ѯ�����
'End If
'�ĳ�ֻҪ�㲡֢ѯ�ʾ���ʾ  '2015-12-11 by Ĳ��
If SSTab1.Tab = 5 Then
sub��ѯ�����
End If

End Sub

'Ϊ����ϵͳ��Ż�ȡ���㣬���ͷţ��Դ���lostfocus�¼�
Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Dim i As Integer
    Dim lobjRec As Object
    'Dim lobjDetl As Object
    On Error GoTo errHandler
    'Set lobjRec = CreateObject("ְҵ��ʷ¼��.clscareerhstregt")
    
    '��ȡ����
    Set lobjRec = pobjDict.FetchEx("�����ֵ�")
    ctxt����.Clear
    ctxt����.AddItem ""
    For i = 1 To lobjRec.RecordCount
        ctxt����.AddItem lobjRec("����")
        ctxt����.ItemData(ctxt����.NewIndex) = lobjRec("���")
        lobjRec.MoveNext
    Next
    ctxt����.ListIndex = 0
    
    '��ȡ����
    Set lobjRec = pobjDict.FetchEx("�����ֵ�")
    ctxt����.Clear
    ctxt����.AddItem ""
    For i = 1 To lobjRec.RecordCount
        ctxt����.AddItem lobjRec("����")
        ctxt����.ItemData(ctxt����.NewIndex) = lobjRec("���")
        lobjRec.MoveNext
    Next
    ctxt����.ListIndex = 0
    
    '��ȡ����������
    Set lobjRec = pobjDict.FetchEx("�����������ֵ�")
    ctxt��������.Clear
    ctxt��������.AddItem ""
    For i = 1 To lobjRec.RecordCount
        ctxt��������.AddItem lobjRec("����")
        ctxt��������.ItemData(ctxt��������.NewIndex) = lobjRec("���")
        lobjRec.MoveNext
    Next
    ctxt��������.ListIndex = 0
    
    '��ȡְҵΣ������
    Set lobjRec = pobjDict.FetchEx("Σ�������ֵ�")
    ctxtΣ������.Clear
    ctxtΣ������.AddItem ""
    For i = 1 To lobjRec.RecordCount
        ctxtΣ������.AddItem lobjRec("����")
        ctxtΣ������.ItemData(ctxtΣ������.NewIndex) = lobjRec("���")
        lobjRec.MoveNext
    Next
    ctxtΣ������.ListIndex = 0
    
    '��ȡ�Ծ�֢״�̶�
    Set lobjRec = pobjDict.FetchEx("����̶��ֵ�")
    ctxt�̶�.Clear
    ctxt�̶�.AddItem ""
    For i = 1 To lobjRec.RecordCount
        ctxt�̶�.AddItem lobjRec("����")
        ctxt�̶�.ItemData(ctxt�̶�.NewIndex) = lobjRec("���")
        lobjRec.MoveNext
    Next
    ctxt�̶�.ListIndex = 0
    
    
    Set lobjRec = pobjDict.FetchEx("ְҵ������Դ�ֵ�")
    Combo2.Clear
    Combo2.AddItem ""
    For i = 1 To lobjRec.RecordCount
       Combo2.AddItem lobjRec("����")
        Combo2.ItemData(Combo2.NewIndex) = lobjRec("���")
        lobjRec.MoveNext
    Next
    Combo2.ListIndex = 0
    
    
    Set lobjRec = pobjDict.FetchEx("ְҵ��������ʩ�ֵ�")
    Combo4.Clear
    Combo4.AddItem ""
    For i = 1 To lobjRec.RecordCount
    Combo4.AddItem lobjRec("����")
    Combo4.ItemData(Combo4.NewIndex) = lobjRec("���")
    lobjRec.MoveNext
    Next
    Combo4.ListIndex = 0
    
'   8023����������   2015-11-26 by Ĳ��
'      If Label�������.Caption = "8023����" Then
'    Set lobjRec = pobjDict.FetchEx("ְҵ��8023�����ֵ�")
'    Combo3.Clear
'    Combo3.AddItem ""
'    For i = 1 To lobjRec.RecordCount
'    Combo3.AddItem lobjRec("����")
'    Combo3.ItemData(Combo3.NewIndex) = lobjRec("���")
'    lobjRec.MoveNext
'    Next
'    Combo3.ListIndex = 0
'    Else
    Set lobjRec = pobjDict.FetchEx("ְҵ�����������ֵ�")
    Combo3.Clear
    Combo3.AddItem ""
    For i = 1 To lobjRec.RecordCount
    Combo3.AddItem lobjRec("����")
    Combo3.ItemData(Combo3.NewIndex) = lobjRec("���")
    lobjRec.MoveNext
    Next
    Combo3.ListIndex = 0
'    End If
    
    
    
      Set lobjRec = pobjDict.FetchEx("ְҵ���������ƾ����ֵ�")
    Combo5.Clear
    Combo5.AddItem "��ѡ�����ƾ�������ģ�塣"
    For i = 1 To lobjRec.RecordCount
    Combo5.AddItem lobjRec("����")
    Combo5.ItemData(Combo5.NewIndex) = lobjRec("���")
    lobjRec.MoveNext
    Next
    Combo5.ListIndex = 0
    
          Set lobjRec = pobjDict.FetchEx("ְҵ��ת���ֵ�")
    Combo6.Clear
    Combo6.AddItem ""
    For i = 1 To lobjRec.RecordCount
    Combo6.AddItem lobjRec("����")
    Combo6.ItemData(Combo6.NewIndex) = lobjRec("���")
    lobjRec.MoveNext
    Next
    Combo6.ListIndex = 0
    '��ȡ�Ծ�֢״����
'    Set lobjRec = pobjDict.FetchEx("ְҵ�����֢״�ֵ�")
    dasubSetQueryTimeout 6000
    Set lobjRec = dafuncGetData("select ����,��� from ϵͳ����_�ֵ�_�ֵ����ݱ� where ID in " _
                            & "(select ID from ϵͳ����_�ֵ�_�ֵ���б� where ����='ְҵ�����֢״�ֵ�') and Parent='0'")
    ccmb����.Clear
    ccmb����.AddItem ""
    For i = 1 To lobjRec.RecordCount
        ccmb����.AddItem lobjRec("����")
        ccmb����.ItemData(ccmb����.NewIndex) = lobjRec("���")
        lobjRec.MoveNext
    Next
    ccmb����.ListIndex = 0
    
    If ccmb����.Text <> "" Then
        '��ȡ�Ծ�֢״��Ŀ
        dasubSetQueryTimeout 6000
        Set lobjRec = dafuncGetData("select ����,��� from ϵͳ����_�ֵ�_�ֵ����ݱ� where parent in " _
                                & "(select innerid from ϵͳ����_�ֵ�_�ֵ����ݱ� where ���� = '" & ccmb����.Text & " ')")
        clst��Ŀ.Clear
    '    ccmb����.AddItem ""
        For i = 1 To lobjRec.RecordCount
            clst��Ŀ.AddItem lobjRec("����")
            clst��Ŀ.ItemData(clst��Ŀ.NewIndex) = lobjRec("���")
            lobjRec.MoveNext
        Next
        clst��Ŀ.ListIndex = 0
    End If
    
    
    '8023��mIndex=1ʱҲ��Ҫ�̶��ֵ䣬���н��ж����ȥ��   2015-11-11 by Ĳ��
'    If mIndex <> 1 Then
        Set lobjRec = pobjDict.FetchEx("�̶��ֵ�")
        ccmb����(mIndex).Clear
        ccmb����(mIndex).AddItem ""
        For i = 1 To lobjRec.RecordCount
            ccmb����(mIndex).AddItem lobjRec("����")
            ccmb����(mIndex).ItemData(ccmb����(mIndex).NewIndex) = lobjRec("���")
            lobjRec.MoveNext
        Next
        ccmb����(mIndex).ListIndex = 0
        Set lobjRec = pobjDict.FetchEx("�̶��ֵ�")
        ccmb����(mIndex).Clear
        ccmb����(mIndex).AddItem ""
        For i = 1 To lobjRec.RecordCount
            ccmb����(mIndex).AddItem lobjRec("����")
            ccmb����(mIndex).ItemData(ccmb����(mIndex).NewIndex) = lobjRec("���")
            lobjRec.MoveNext
        Next
        ccmb����(mIndex).ListIndex = 0
'    End If
    
    Set lobjRec = Nothing
    'ע�� ��ΰ 2015-4-7
    'ctxt����ʱ��.Text = Date
'    ctxt�������.Value = Date
'    ctxt�Ӵ�ʱ��.Value = Date
'    ctxt��ʼ.Value = Date
'    ctxt����.Value = Date
    cgrdְҵʷ.ColHidden(mcolIndex("ϵͳ���")) = True
    'ְҵ������ʱ����ʾ����������
    If Label�������.Caption <> "ְҵ����" Then
       cgrdְҵʷ.ColHidden(mcolIndex("�Ӵ�ʱ��")) = True
    End If
    cgrd��ʷ.ColHidden(mcolIndexwkdis("ϵͳ���")) = True
    cgrd֢״.ColHidden(mcolindexzz("ϵͳ���")) = True
    
    ctxtsysno.SetFocus
    If ���ʼǺ� = 1 Then
        'Ccmb���(mIndex).SetFocus
    End If
'    MsgBox "timer1����" '2016-3-1 by Ĳ��
    Exit Sub
errHandler:
    sfsub������ "ְҵ��ʷ¼��", "frmcareerhstregt", "sub�޸�����ʷ", Err.Number, Err.Description, True
End Sub

Public Sub subPrint(ByVal paraϵͳ��� As String)

'    Set lcol��� = New Collection
'        lcol���.Add Trim(ctxtsysno.Text)
'
'        Set lobj������� = CreateObject("ְҵ������.clsmedicalexam")
'        lobj�������.ϵͳ��� = Trim(ctxtsysno.Text)
'        str������� = lobj�������.�������
'            '��ӡ
'            pobjҵ�����.Sub��ӡ���� str������� & "���Ǽǵ�", lcol���, False
'            'Cancel = True
'        Set lobj������� = Nothing
    On Error GoTo errHandler
    Dim lobjRec As Object
    Dim lcolInfo As Collection
    Dim lcolItem As Collection
    Dim lstrRptName As String
    Dim mstrPrint As String
    Dim mstrselect As String
    Dim lobj������� As Object
    Dim str������� As String
    Set lobj������� = CreateObject("ְҵ������.clsmedicalexam")
    lobj�������.ϵͳ��� = Trim(paraϵͳ���)
    str������� = lobj�������.�������
    Set lobj������� = Nothing
    lstrRptName = str������� & "���Ǽǵ�"
    mstrPrint = Trim(paraϵͳ���)

    Set lcolInfo = New Collection
    Set lcolItem = New Collection
    lcolItem.Add "ϵͳ���", "����"
    lcolItem.Add mstrPrint, "ֵ"
    lcolInfo.Add lcolItem, lcolItem("����")
'    Set lcolItem = New Collection
'    lcolItem.Add "ѡ������", "����"
'    lcolItem.Add mstrselect, "ֵ"
'    lcolInfo.Add lcolItem, lcolItem("����")
    '��ӡ   falseΪ��Ԥ��
    
     '��ȡǩ��ͼƬ
    Dim lpicPhoto As StdPicture
    Dim lobjSys As Object
    Set lobjSys = CreateObject("Scripting.FileSystemObject")
    '�ȿ����հ�ǩ����ͼƬ��
    lobjSys.CopyFile App.Path & "\�հ���Ƭ.bmp", "c:\�����Ƭ.bmp"
    
    Set lpicPhoto = pmfunc��ȡͼƬ(mstrPrint, "ְҵ�����")
    
    'Set lpicPhoto = pmfunc��ȡͼƬ("0001", "ϵͳ����")
    '���ݿ���û��ͼƬʱ��lpicPhoto����ֵΪ0��������null
    If Not lpicPhoto = 0 Then
        SavePicture lpicPhoto, "c:\�����Ƭ.bmp"
    End If
    
    Set lobjRec = CreateObject("ְҵ������.cls����")
    lobjRec.funcCHPrintReport lstrRptName, lcolInfo, App.Path, False

    Exit Sub
errHandler:
    Dim llngErr As Long
    Dim lstrError As String
    llngErr = Err.Number
    lstrError = Err.Description

    If llngErr = 20526 Then
        lstrError = "���ڷ�����ӡ�����⡣�˴��������ԭ�򼰽���������£� " & Chr(13) & Chr(10) & Chr(13) & Chr(10) _
                    & "  (1) û�д� Windows ��������а�װ��ӡ����" & Chr(13) & Chr(10) _
                    & "      ������򿪿�����壬˫������ӡ����ͼ�꣬ѡ����Ӵ�ӡ������װ���ӡ����" & Chr(13) & Chr(10) & Chr(13) & Chr(10) _
                    & "  (2) ��ӡ��û���ߣ�" & Chr(13) & Chr(10) _
                    & "      ���������ӡ���������������Ƿ�������" & Chr(13) & Chr(10) & Chr(13) & Chr(10) _
                    & "  (3)  ��ӡ��������ȱֽ��" & Chr(13) & Chr(10) _
                    & "      ����������Щ���⡣" & Chr(13) & Chr(10) & Chr(13) & Chr(10) _
                    & "  (4) ��ͼ��ֻ�ܽ����ı��Ĵ�ӡ���ϴ�ӡ���壿" & Chr(13) & Chr(10) _
                    & "      ������л���һ̨�ܴ�ӡͼ�εĴ�ӡ����"
        llngErr = 6666
    End If
    sfsub������ "�����豸�������", "frmBiologMaterialApply", "subPrint", Err.Number, Err.Description, False
End Sub

'2012-06-13 �ڵ�� ��
'��ʼ�����һ��������沿��
Sub subInit���һ�����()
    combӪ��.Clear
    combӪ��.AddItem "����": combӪ��.ItemData(combӪ��.NewIndex) = 0
    combӪ��.AddItem "�е�": combӪ��.ItemData(combӪ��.NewIndex) = 1
    combӪ��.AddItem "����": combӪ��.ItemData(combӪ��.NewIndex) = 2
    combӪ��.ListIndex = 0
 '���ӷ����Ŀ�ѡֵ 2015-11-27 by Ĳ��
    comb����.Clear
    comb����.AddItem "������": comb����.ItemData(comb����.NewIndex) = 0
    comb����.AddItem "������": comb����.ItemData(comb����.NewIndex) = 1
    comb����.AddItem "������": comb����.ItemData(comb����.NewIndex) = 2
    comb����.ListIndex = 0
End Sub

'2012-06-13 �ڵ�� ��
'������һ��������沿��
Sub subClear���һ�����()
    combӪ��.ListIndex = 0
    ctxt���.Text = ""
    ctxt����.Text = ""
    ctxt����.Text = ""
    ctxt����ָ��.Text = ""
    ctxt����ѹ.Text = ""
    ctxt����ѹ.Text = ""
    comb����.ListIndex = 0
End Sub

'2012-06-13 �ڵ�� ��
'�������һ��������沿��(���һ��������������ڿ�)
Sub subLoad���һ�����(ByVal paraSysNo As String)
    Dim lobjRec As Object
    Dim lobjTemp As Object
    Dim lobjResult As Object
    
    Set lobjTemp = CreateObject("ְҵ��ʷ¼��.clsCareerHstRegt")
    
    'Ӫ��
    Set lobjRec = lobjTemp.func��ȡ�����Ŀ���("Ӫ��", "13")
    Set lobjResult = lobjTemp.func��ȡ���˵��������(paraSysNo, lobjRec("����"))
    If Not lobjResult Is Nothing Then
        If lobjResult.RecordCount > 0 Then
            combӪ��.Enabled = True
        Else
            combӪ��.Enabled = False
        End If
        If IsNull(lobjResult("�����")) = False Then combӪ��.Text = lobjResult("�����")
    Else
        combӪ��.Enabled = False
    End If
    
    '���
    Set lobjRec = lobjTemp.func��ȡ�����Ŀ���("���", "13")
    Set lobjResult = lobjTemp.func��ȡ���˵��������(paraSysNo, lobjRec("����"))
    If Not lobjResult Is Nothing Then
        If lobjResult.RecordCount > 0 Then
            ctxt���.Enabled = True
        Else
            ctxt���.Enabled = False
        End If
        If IsNull(lobjResult("�����")) = False Then ctxt���.Text = lobjResult("�����")
    Else
        ctxt���.Enabled = False
    End If
    
    '����
    Set lobjRec = lobjTemp.func��ȡ�����Ŀ���("����", "13")
    Set lobjResult = lobjTemp.func��ȡ���˵��������(paraSysNo, lobjRec("����"))
    If Not lobjResult Is Nothing Then
        If lobjResult.RecordCount > 0 Then
            ctxt����.Enabled = True
        Else
            ctxt����.Enabled = False
        End If
        If IsNull(lobjResult("�����")) = False Then ctxt����.Text = lobjResult("�����")
    Else
        ctxt����.Enabled = False
    End If
    
    '����
    Set lobjRec = lobjTemp.func��ȡ�����Ŀ���("����", "13")
    Set lobjResult = lobjTemp.func��ȡ���˵��������(paraSysNo, lobjRec("����"))
    If Not lobjResult Is Nothing Then
        If lobjResult.RecordCount > 0 Then
            ctxt����.Enabled = True
        Else
            ctxt����.Enabled = False
        End If
        If IsNull(lobjResult("�����")) = False Then ctxt����.Text = lobjResult("�����")
    Else
        ctxt����.Enabled = False
    End If
    
    '����ָ��
    Set lobjRec = lobjTemp.func��ȡ�����Ŀ���("����ָ��", "13")
    Set lobjResult = lobjTemp.func��ȡ���˵��������(paraSysNo, lobjRec("����"))
    If Not lobjResult Is Nothing Then
        If lobjResult.RecordCount > 0 Then
            ctxt����ָ��.Enabled = True
        Else
            ctxt����ָ��.Enabled = False
        End If
        If IsNull(lobjResult("�����")) = False Then ctxt����ָ��.Text = lobjResult("�����")
    Else
        ctxt����ָ��.Enabled = False
    End If
    
    '����ѹ
    Set lobjRec = lobjTemp.func��ȡ�����Ŀ���("����ѹ", "13")
    Set lobjResult = lobjTemp.func��ȡ���˵��������(paraSysNo, lobjRec("����"))
    If Not lobjResult Is Nothing Then
        If lobjResult.RecordCount > 0 Then
            ctxt����ѹ.Enabled = True
        Else
            ctxt����ѹ.Enabled = False
        End If
        If IsNull(lobjResult("�����")) = False Then ctxt����ѹ.Text = lobjResult("�����")
    Else
        ctxt����ѹ.Enabled = False
    End If
    
    '����ѹ
    Set lobjRec = lobjTemp.func��ȡ�����Ŀ���("����ѹ", "13")
    Set lobjResult = lobjTemp.func��ȡ���˵��������(paraSysNo, lobjRec("����"))
    If Not lobjResult Is Nothing Then
        If lobjResult.RecordCount > 0 Then
            ctxt����ѹ.Enabled = True
        Else
            ctxt����ѹ.Enabled = False
        End If
        If IsNull(lobjResult("�����")) = False Then ctxt����ѹ.Text = lobjResult("�����")
    Else
        ctxt����ѹ.Enabled = False
    End If
'    '����������ʾ 2015-7-1 by lanchao
'    Dim xls As Object
'    Set xls = dafuncGetData("select ����� from ְҵ�����_�����Ϣ_�ڿ� where �����Ŀ='02002' and ϵͳ���='" & paraSysNo & "'")
'    If Not IsNull(xls("�����")) Then ctxtxinlv.Text = xls("�����")
   
    '�޸����������������ʾ  2015-12-2 by Ĳ��
    Dim xls As Object
    Set xls = dafuncGetData("select ����� from ְҵ�����_�����Ϣ_�ڿ� where �����Ŀ='02002' and ϵͳ���='" & paraSysNo & "'")
    If xls.RecordCount > 0 And Not IsNull(xls("�����")) Then
     ctxtxinlv.Text = xls("�����")
     Else
     ctxtxinlv.Text = ""
    End If
'    '���ӷ�����ʾ 2015-11-26 by Ĳ��
    Dim ttype As Object
    Dim fy As Object
    Set ttype = dafuncGetData("select �������� from ְҵ�����_�����Ա������Ϣ�� where ϵͳ���='" & paraSysNo & "'")
    If ttype.RecordCount > 0 And ttype("��������") = "���佡��" Then
        Set fy = dafuncGetData("select ����� from ְҵ�����_�����Ϣ_�ڿ� where �����Ŀ='02019' and ϵͳ���='" & paraSysNo & "'")
        If fy.RecordCount > 0 And Not IsNull(fy("�����")) Then
        comb����.Text = fy("�����")
        Else
        comb����.Text = ""
        End If
    End If
End Sub

'2012-06-13 �ڵ�� ��
'�������һ��������沿��(���һ��������������ڿ�)
Sub subSave���һ�����(ByVal paraSysNo As String)
    Dim lobjRec As Object
    Dim lobjTemp As Object
    Dim lstr��� As String
    Dim lstr���� As String
    Dim lstr����ѹ As String
    Dim lstr����ѹ As String
    Dim lstr���� As String
    Dim lstr���� As String
    Dim lstr����ָ�� As String
    '2015-7-1 modify by lanchao  һ��������ݲ�����λ����
    'lstr��� = ctxt���.Text & " cm"
    lstr��� = ctxt���.Text
    lstr���� = ctxt����.Text
    lstr����ѹ = ctxt����ѹ.Text
    lstr����ѹ = ctxt����ѹ.Text
    lstr���� = comb����.Text   '2015-11-26 by Ĳ��
    lstr���� = ctxt����.Text
    lstr����ָ�� = ctxt����ָ��.Text
    Set lobjTemp = CreateObject("ְҵ��ʷ¼��.clsCareerHstRegt")
       
    
    'Ӫ��
    Set lobjRec = lobjTemp.func��ȡ�����Ŀ���("Ӫ��", "13")
    lobjTemp.func���浥�˵�������� paraSysNo, "13", lobjRec("����"), combӪ��.Text
    
    '���
    Set lobjRec = lobjTemp.func��ȡ�����Ŀ���("���", "13")
    lobjTemp.func���浥�˵�������� paraSysNo, "13", lobjRec("����"), lstr���

    '����
    Set lobjRec = lobjTemp.func��ȡ�����Ŀ���("����", "13")
    lobjTemp.func���浥�˵�������� paraSysNo, "13", lobjRec("����"), lstr����
    
    '����
    Set lobjRec = lobjTemp.func��ȡ�����Ŀ���("����", "13")
    lobjTemp.func���浥�˵�������� paraSysNo, "13", lobjRec("����"), lstr����
    
    '����ָ��
    Set lobjRec = lobjTemp.func��ȡ�����Ŀ���("����ָ��", "13")
    lobjTemp.func���浥�˵�������� paraSysNo, "13", lobjRec("����"), lstr����ָ��
    
    '����ѹ
    Set lobjRec = lobjTemp.func��ȡ�����Ŀ���("����ѹ", "13")
    lobjTemp.func���浥�˵�������� paraSysNo, "13", lobjRec("����"), lstr����ѹ
    
    '����ѹ
    Set lobjRec = lobjTemp.func��ȡ�����Ŀ���("����ѹ", "13")
    lobjTemp.func���浥�˵�������� paraSysNo, "13", lobjRec("����"), lstr����ѹ
    
    '����   2015-11-26 by Ĳ��
    Dim teststylelob As Object
    Dim lob As Object
    Dim teststyle As String
    Set teststylelob = dafuncGetData("select ������� from ְҵ�����_��������Ϣ�� where ϵͳ���='" & Trim(ctxtsysno.Text) & "'")
    teststyle = teststylelob("�������")
    If teststyle = "���佡��" Then
        Set lob = dafuncGetData("select �����Ŀ from ְҵ�����_�����Ϣ_�ڿ�  where �����Ŀ='02019' and ϵͳ���='" & Trim(ctxtsysno.Text) & "'")
        If lob.RecordCount < 1 Then
    '    dafuncGetData ("insert into ְҵ�����_�����Ϣ_�ڿ� values('" & Trim(ctxtsysno.Text) & "','02019','" & comb����.Text & "','" & um�û���� & "' ,'" & Now & "','" & conclusion & "')")
        dafuncGetData ("insert into ְҵ�����_�����Ϣ_�ڿ�( ϵͳ���,�����Ŀ,�����,���ҽʦ,��дʱ��) values('" & Trim(ctxtsysno.Text) & "','02019','" & comb����.Text & "','" & um�û���� & "','" & Now & "')")
        Else
        dafuncGetData ("update ְҵ�����_�����Ϣ_�ڿ� set �����='" & comb����.Text & "' ,���ҽʦ='" & um�û���� & "' ,��дʱ��='" & Now & "' where �����Ŀ='02019' and ϵͳ���='" & Trim(ctxtsysno.Text) & "'")
        End If
    End If
End Sub

'2012-06-13 �ڵ�� ��
'ɾ�����һ��������沿��(���һ��������������ڿ�)
Sub subDel���һ�����(ByVal paraSysNo As String)
    Dim lobjRec As Object
    Dim lobjTemp As Object
    Set lobjTemp = CreateObject("ְҵ��ʷ¼��.clsCareerHstRegt")
    
    'Ӫ��
    Set lobjRec = lobjTemp.func��ȡ�����Ŀ���("Ӫ��", "13")
    lobjTemp.funcɾ�����˵�������� paraSysNo, "13", lobjRec("����")
    
    '���
    Set lobjRec = lobjTemp.func��ȡ�����Ŀ���("���", "13")
    lobjTemp.funcɾ�����˵�������� paraSysNo, "13", lobjRec("����")

    '����
    Set lobjRec = lobjTemp.func��ȡ�����Ŀ���("����", "13")
    lobjTemp.funcɾ�����˵�������� paraSysNo, "13", lobjRec("����")
    
    '����ѹ
    Set lobjRec = lobjTemp.func��ȡ�����Ŀ���("����ѹ", "13")
    lobjTemp.funcɾ�����˵�������� paraSysNo, "13", lobjRec("����")
    
    '����ѹ
    Set lobjRec = lobjTemp.func��ȡ�����Ŀ���("����ѹ", "13")
    lobjTemp.funcɾ�����˵�������� paraSysNo, "13", lobjRec("����")
    
    subClear���һ�����
End Sub
Private Sub sub�����ն�()
      With MSComm1
       
        .CommPort = 1
        .Settings = "4800,N,8,1"
        .InBufferSize = 1024 'ԭ��Ϊ19
        .RThreshold = 1      '����1�ֽڴ���oncomm�¼�
        .InputMode = comInputModeBinary
        .InputLen = 1 '���볤��Ϊ19
        .InBufferCount = 0      '������ջ�����
    End With
        '�򿪶˿�
    If MSComm1.PortOpen Then MSComm1.PortOpen = False
                
'                MSComm1.CommPort = 7  '�ٶ�����COM5��
                MSComm1.CommPort = 1
                
                ' �趨�������ʵȣ������������������
                MSComm1.Settings = "4800,N,8,1"
    
                MSComm1.PortOpen = True
                Text��ʱ.Text = ""
End Sub
