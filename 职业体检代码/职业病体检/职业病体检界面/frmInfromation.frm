VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmInfromation 
   Caption         =   "������Ϣ"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13320
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   13320
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame Frame8 
      Caption         =   "������Ϣ"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12375
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   4320
         ScaleHeight     =   1785
         ScaleWidth      =   1545
         TabIndex        =   1
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label�绰 
         Height          =   375
         Left            =   9480
         TabIndex        =   235
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label38 
         Caption         =   "�绰��"
         Height          =   255
         Left            =   9360
         TabIndex        =   234
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Lab��� 
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
         TabIndex        =   141
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label LabelΣ������ 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7200
         TabIndex        =   140
         Top             =   1320
         Width           =   90
      End
      Begin VB.Label Label30 
         Caption         =   "Σ��  ���أ�"
         Height          =   255
         Index           =   1
         Left            =   6000
         TabIndex        =   139
         Top             =   1320
         Width           =   1095
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
         TabIndex        =   16
         Top             =   960
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
         TabIndex        =   15
         Top             =   840
         Width           =   3015
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
         TabIndex        =   14
         Top             =   360
         Width           =   570
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "�ֹ�����λ��"
         Height          =   255
         Left            =   6000
         TabIndex        =   13
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label lab��λ 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   7200
         TabIndex        =   12
         Top             =   240
         Width           =   90
      End
      Begin VB.Label Label29 
         Caption         =   "��  ��  �֣�"
         Height          =   255
         Left            =   6000
         TabIndex        =   11
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label30 
         Caption         =   "��  ְ  ��"
         Height          =   255
         Index           =   0
         Left            =   6000
         TabIndex        =   10
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Lab�ֹ��� 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7200
         TabIndex        =   9
         Top             =   600
         Width           =   90
      End
      Begin VB.Label Lab��ְ�� 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7200
         TabIndex        =   8
         Top             =   960
         Width           =   90
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
         TabIndex        =   7
         Top             =   1440
         Width           =   615
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
         TabIndex        =   6
         Top             =   1440
         Width           =   615
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
         TabIndex        =   5
         Top             =   1560
         Width           =   570
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
         TabIndex        =   4
         Top             =   1560
         Width           =   570
      End
      Begin VB.Label Label70 
         Caption         =   "���  ���ͣ�"
         Height          =   255
         Left            =   6000
         TabIndex        =   3
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label������� 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7200
         TabIndex        =   2
         Top             =   1680
         Width           =   90
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5895
      Left            =   120
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2160
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   10398
      _Version        =   393216
      Tabs            =   6
      Tab             =   1
      TabsPerRow      =   6
      TabHeight       =   520
      BackColor       =   16777215
      ForeColor       =   16711680
      TabCaption(0)   =   "��������ʷ"
      TabPicture(0)   =   "frmInfromation.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "freNuclear"
      Tab(0).Control(1)=   "freRadiation"
      Tab(0).Control(2)=   "freOrdinary"
      Tab(0).Control(3)=   "ctxtOther"
      Tab(0).Control(4)=   "Frame1"
      Tab(0).Control(5)=   "Frame6"
      Tab(0).Control(6)=   "Label5(0)"
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "ְҵʷ"
      TabPicture(1)   =   "frmInfromation.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Labְҵʷ"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cgrdְҵʷ"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "������ʷ(����ְҵ��ʷ)"
      TabPicture(2)   =   "frmInfromation.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Lab��ʷ"
      Tab(2).Control(1)=   "cgrd��ʷ"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "�Ծ�֢״"
      TabPicture(3)   =   "frmInfromation.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cgrd֢״"
      Tab(3).Control(1)=   "Lab֢״"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "���һ�����"
      TabPicture(4)   =   "frmInfromation.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame5"
      Tab(4).Control(1)=   "Frame7"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "��״ѯ��"
      TabPicture(5)   =   "frmInfromation.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "cgrdzzxw"
      Tab(5).ControlCount=   1
      Begin VB.Frame freNuclear 
         BorderStyle     =   0  'None
         Height          =   3615
         Left            =   -74760
         TabIndex        =   18
         Top             =   600
         Width           =   10815
         Begin VB.Frame Frame17 
            Caption         =   "�̾�ʷ"
            ForeColor       =   &H000080FF&
            Height          =   2295
            Left            =   5880
            TabIndex        =   39
            Top             =   1320
            Width           =   5055
            Begin VB.Label Label40 
               Caption         =   "��"
               Height          =   255
               Left            =   1800
               TabIndex        =   229
               Top             =   2040
               Width           =   375
            End
            Begin VB.Label Lab����ʱ�� 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   960
               TabIndex        =   228
               Top             =   2040
               Width           =   855
            End
            Begin VB.Label Lab���̶̳� 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   3360
               TabIndex        =   227
               Top             =   1680
               Width           =   975
            End
            Begin VB.Label Lab���Ƴ̶� 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   960
               TabIndex        =   226
               Top             =   1680
               Width           =   855
            End
            Begin VB.Label Label28 
               Caption         =   "����ʱ����"
               Height          =   255
               Left            =   120
               TabIndex        =   225
               Top             =   2040
               Width           =   975
            End
            Begin VB.Label Label26 
               Caption         =   "���̶̳ȣ�"
               Height          =   255
               Left            =   2520
               TabIndex        =   224
               Top             =   1680
               Width           =   1095
            End
            Begin VB.Label Label25 
               Caption         =   "���Ƴ̶ȣ�"
               Height          =   255
               Left            =   120
               TabIndex        =   223
               Top             =   1680
               Width           =   975
            End
            Begin VB.Label Lab���� 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   3240
               TabIndex        =   212
               Top             =   1320
               Width           =   975
            End
            Begin VB.Label Lab���� 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   840
               TabIndex        =   211
               Top             =   1320
               Width           =   975
            End
            Begin VB.Label Lab������ 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   3240
               TabIndex        =   210
               Top             =   960
               Width           =   975
            End
            Begin VB.Label Lab������ 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   840
               TabIndex        =   209
               Top             =   960
               Width           =   975
            End
            Begin VB.Label Lab��ʳϰ�� 
               BackColor       =   &H00FFFFFF&
               Height          =   375
               Index           =   1
               Left            =   120
               TabIndex        =   208
               Top             =   480
               Width           =   4455
            End
            Begin VB.Label Label91 
               AutoSize        =   -1  'True
               Caption         =   "��������"
               Height          =   180
               Left            =   2520
               TabIndex        =   48
               Top             =   960
               Width           =   720
            End
            Begin VB.Label Label90 
               AutoSize        =   -1  'True
               Caption         =   "��������"
               Height          =   180
               Left            =   120
               TabIndex        =   47
               Top             =   960
               Width           =   720
            End
            Begin VB.Label Label83 
               AutoSize        =   -1  'True
               Caption         =   "֧/��"
               Height          =   180
               Left            =   4320
               TabIndex        =   46
               Top             =   960
               Width           =   450
            End
            Begin VB.Label Label87 
               AutoSize        =   -1  'True
               Caption         =   "ML/��"
               Height          =   180
               Left            =   1920
               TabIndex        =   45
               Top             =   960
               Width           =   450
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               Caption         =   "�����ס��������ʳϰ�ߡ��̾��Ⱥ�������"
               Height          =   180
               Index           =   0
               Left            =   120
               TabIndex        =   44
               Top             =   240
               Width           =   3420
            End
            Begin VB.Label Label88 
               Caption         =   "���䣺"
               Height          =   255
               Left            =   120
               TabIndex        =   43
               Top             =   1320
               Width           =   720
            End
            Begin VB.Label Label92 
               Caption         =   "��"
               Height          =   255
               Left            =   1920
               TabIndex        =   42
               Top             =   1320
               Width           =   495
            End
            Begin VB.Label Label93 
               Caption         =   "���䣺"
               Height          =   255
               Left            =   2520
               TabIndex        =   41
               Top             =   1320
               Width           =   735
            End
            Begin VB.Label Label94 
               Caption         =   "��"
               Height          =   255
               Left            =   4320
               TabIndex        =   40
               Top             =   1320
               Width           =   495
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "����ʷ(����ż����ʷ)"
            ForeColor       =   &H000080FF&
            Height          =   2295
            Index           =   1
            Left            =   120
            TabIndex        =   25
            Top             =   1320
            Width           =   5775
            Begin VB.Label Lab��Ů���� 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   3720
               TabIndex        =   207
               Top             =   2040
               Width           =   1335
            End
            Begin VB.Label LabŮ���������� 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   2280
               TabIndex        =   206
               Top             =   2040
               Width           =   1335
            End
            Begin VB.Label Lab�к��������� 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   2280
               TabIndex        =   205
               Top             =   1680
               Width           =   1335
            End
            Begin VB.Label LabŮ�� 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   960
               TabIndex        =   204
               Top             =   2040
               Width           =   375
            End
            Begin VB.Label Lab��Ů�� 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   960
               TabIndex        =   203
               Top             =   1680
               Width           =   375
            End
            Begin VB.Label Lab����ԭ�� 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   960
               TabIndex        =   202
               Top             =   1200
               Width           =   4215
            End
            Begin VB.Label Lab���� 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   201
               Top             =   1200
               Width           =   615
            End
            Begin VB.Label Lab��̥ 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   4560
               TabIndex        =   200
               Top             =   600
               Width           =   615
            End
            Begin VB.Label Lab��̥ 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   3480
               TabIndex        =   199
               Top             =   600
               Width           =   615
            End
            Begin VB.Label Lab���� 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   2640
               TabIndex        =   198
               Top             =   600
               Width           =   615
            End
            Begin VB.Label Lab��� 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   1800
               TabIndex        =   197
               Top             =   600
               Width           =   615
            End
            Begin VB.Label Lab��� 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   960
               TabIndex        =   196
               Top             =   600
               Width           =   615
            End
            Begin VB.Label Lab�д� 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   195
               Top             =   600
               Width           =   615
            End
            Begin VB.Label Label89 
               AutoSize        =   -1  'True
               Caption         =   "��Ů����״����"
               Height          =   300
               Left            =   3600
               TabIndex        =   38
               Top             =   1680
               Width           =   1260
            End
            Begin VB.Label Label86 
               AutoSize        =   -1  'True
               Caption         =   "��̥��"
               Height          =   180
               Left            =   4560
               TabIndex        =   37
               Top             =   360
               Width           =   540
            End
            Begin VB.Label Label85 
               AutoSize        =   -1  'True
               Caption         =   "��̥��"
               Height          =   180
               Left            =   3480
               TabIndex        =   36
               Top             =   360
               Width           =   540
            End
            Begin VB.Label Label84 
               AutoSize        =   -1  'True
               Caption         =   "�дΣ�"
               Height          =   180
               Left            =   120
               TabIndex        =   35
               Top             =   360
               Width           =   540
            End
            Begin VB.Label Label82 
               AutoSize        =   -1  'True
               Caption         =   "�����"
               Height          =   180
               Left            =   960
               TabIndex        =   34
               Top             =   360
               Width           =   540
            End
            Begin VB.Label Label81 
               AutoSize        =   -1  'True
               Caption         =   "�����"
               Height          =   180
               Left            =   1800
               TabIndex        =   33
               Top             =   360
               Width           =   540
            End
            Begin VB.Label Label80 
               AutoSize        =   -1  'True
               Caption         =   "������"
               Height          =   180
               Left            =   2640
               TabIndex        =   32
               Top             =   360
               Width           =   540
            End
            Begin VB.Label Label79 
               AutoSize        =   -1  'True
               Caption         =   "���в���ԭ��"
               Height          =   180
               Left            =   960
               TabIndex        =   31
               Top             =   960
               Width           =   1260
            End
            Begin VB.Label Label75 
               Caption         =   "�����к���"
               Height          =   255
               Left            =   120
               TabIndex        =   30
               Top             =   1680
               Width           =   975
            End
            Begin VB.Label Label76 
               Caption         =   "����Ů����"
               Height          =   255
               Left            =   120
               TabIndex        =   29
               Top             =   2040
               Width           =   975
            End
            Begin VB.Label Label77 
               Caption         =   "�������ڣ�"
               Height          =   225
               Left            =   1440
               TabIndex        =   28
               Top             =   1680
               Width           =   975
            End
            Begin VB.Label Label78 
               Caption         =   "�������ڣ�"
               Height          =   225
               Left            =   1440
               TabIndex        =   27
               Top             =   2040
               Width           =   975
            End
            Begin VB.Label Label95 
               Caption         =   "������"
               Height          =   255
               Left            =   120
               TabIndex        =   26
               Top             =   960
               Width           =   855
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "����ʷ"
            ForeColor       =   &H000080FF&
            Height          =   1095
            Index           =   1
            Left            =   120
            TabIndex        =   19
            Top             =   120
            Width           =   10935
            Begin VB.Label Lab��ż���� 
               BackColor       =   &H00FFFFFF&
               Height          =   375
               Index           =   1
               Left            =   5880
               TabIndex        =   194
               Top             =   600
               Width           =   4335
            End
            Begin VB.Label Lab��żְҵ 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   3600
               TabIndex        =   193
               Top             =   720
               Width           =   1455
            End
            Begin VB.Label Lab��ż���� 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   3960
               TabIndex        =   192
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Lab���� 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   960
               TabIndex        =   191
               Top             =   720
               Width           =   1095
            End
            Begin VB.Label Lab��� 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   960
               TabIndex        =   190
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Lab 
               AutoSize        =   -1  'True
               Caption         =   "��ż����״����"
               Height          =   180
               Index           =   1
               Left            =   2760
               TabIndex        =   24
               Top             =   300
               Width           =   1260
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "��ż�Ӵ������������"
               Height          =   180
               Index           =   1
               Left            =   5880
               TabIndex        =   23
               Top             =   240
               Width           =   1800
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "��żְҵ��"
               Height          =   180
               Index           =   1
               Left            =   2760
               TabIndex        =   22
               Top             =   760
               Width           =   900
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "������ڣ�"
               Height          =   180
               Index           =   1
               Left            =   120
               TabIndex        =   21
               Top             =   760
               Width           =   900
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "�Ƿ��飺"
               Height          =   180
               Index           =   1
               Left            =   120
               TabIndex        =   20
               Top             =   300
               Width           =   900
            End
         End
      End
      Begin VB.Frame freRadiation 
         BorderStyle     =   0  'None
         Height          =   3735
         Left            =   -74760
         TabIndex        =   72
         Top             =   600
         Width           =   11895
         Begin VB.Frame Frame3 
            Caption         =   "����ʷ(����ż����ʷ)"
            ForeColor       =   &H000080FF&
            Height          =   2295
            Index           =   0
            Left            =   120
            TabIndex        =   95
            Top             =   1440
            Width           =   5775
            Begin VB.Label Lab��Ů���� 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   4080
               TabIndex        =   161
               Top             =   1920
               Width           =   1215
            End
            Begin VB.Label LabŮ���������� 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   2760
               TabIndex        =   160
               Top             =   1920
               Width           =   1215
            End
            Begin VB.Label Lab�к��������� 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   2760
               TabIndex        =   159
               Top             =   1560
               Width           =   1215
            End
            Begin VB.Label LabŮ�� 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   1080
               TabIndex        =   158
               Top             =   1920
               Width           =   735
            End
            Begin VB.Label Lab��Ů�� 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   1080
               TabIndex        =   157
               Top             =   1560
               Width           =   735
            End
            Begin VB.Label Lab����ԭ�� 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   2040
               TabIndex        =   156
               Top             =   1080
               Width           =   3255
            End
            Begin VB.Label Lab��λ���� 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   1080
               TabIndex        =   155
               Top             =   1080
               Width           =   735
            End
            Begin VB.Label Lab��̥ 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   154
               Top             =   1080
               Width           =   735
            End
            Begin VB.Label Lab��̥ 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   4560
               TabIndex        =   153
               Top             =   480
               Width           =   735
            End
            Begin VB.Label Lab���� 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   3720
               TabIndex        =   152
               Top             =   480
               Width           =   735
            End
            Begin VB.Label Lab���� 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   2880
               TabIndex        =   151
               Top             =   480
               Width           =   735
            End
            Begin VB.Label Lab��� 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   1920
               TabIndex        =   150
               Top             =   480
               Width           =   735
            End
            Begin VB.Label Lab��� 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   1080
               TabIndex        =   149
               Top             =   480
               Width           =   735
            End
            Begin VB.Label Lab�д� 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   147
               Top             =   480
               Width           =   735
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               Caption         =   "��Ů����״����"
               Height          =   180
               Index           =   0
               Left            =   4080
               TabIndex        =   109
               Top             =   1560
               Width           =   1260
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               Caption         =   "�����к���"
               Height          =   180
               Index           =   0
               Left            =   240
               TabIndex        =   108
               Top             =   1560
               Width           =   900
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               Caption         =   "������"
               Height          =   180
               Index           =   0
               Left            =   3720
               TabIndex        =   107
               Top             =   240
               Width           =   540
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               Caption         =   "��̥��"
               Height          =   180
               Index           =   0
               Left            =   4560
               TabIndex        =   106
               Top             =   240
               Width           =   540
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               Caption         =   "��̥��"
               Height          =   180
               Index           =   0
               Left            =   240
               TabIndex        =   105
               Top             =   840
               Width           =   540
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "�дΣ�"
               Height          =   180
               Index           =   0
               Left            =   240
               TabIndex        =   104
               Top             =   240
               Width           =   540
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               Caption         =   "��λ���"
               Height          =   180
               Index           =   0
               Left            =   1080
               TabIndex        =   103
               Top             =   840
               Width           =   900
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "�����"
               Height          =   180
               Index           =   0
               Left            =   1080
               TabIndex        =   102
               Top             =   240
               Width           =   540
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "�����"
               Height          =   180
               Index           =   0
               Left            =   1920
               TabIndex        =   101
               Top             =   240
               Width           =   540
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "������"
               Height          =   180
               Index           =   0
               Left            =   2880
               TabIndex        =   100
               Top             =   240
               Width           =   540
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               Caption         =   "���в���ԭ��"
               Height          =   180
               Index           =   0
               Left            =   2040
               TabIndex        =   99
               Top             =   840
               Width           =   1500
            End
            Begin VB.Label Label71 
               Caption         =   "�������ڣ�"
               Height          =   255
               Left            =   1920
               TabIndex        =   98
               Top             =   1560
               Width           =   975
            End
            Begin VB.Label Label72 
               Caption         =   "����Ů����"
               Height          =   255
               Left            =   240
               TabIndex        =   97
               Top             =   1920
               Width           =   975
            End
            Begin VB.Label Label73 
               Caption         =   "�������ڣ�"
               Height          =   255
               Left            =   1920
               TabIndex        =   96
               Top             =   1920
               Width           =   975
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "����ʷ"
            ForeColor       =   &H000080FF&
            Height          =   1095
            Index           =   0
            Left            =   120
            TabIndex        =   88
            Top             =   120
            Width           =   11055
            Begin MSComCtl2.DTPicker ctxtmarrydate1 
               CausesValidation=   0   'False
               Height          =   300
               Index           =   0
               Left            =   7560
               TabIndex        =   89
               Top             =   120
               Visible         =   0   'False
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   0   'False
               CustomFormat    =   "yyyy/MM"
               Format          =   60227584
               CurrentDate     =   41013
            End
            Begin VB.Label Lab��ż���� 
               BackColor       =   &H00FFFFFF&
               Height          =   495
               Index           =   0
               Left            =   5880
               TabIndex        =   146
               Top             =   480
               Width           =   4815
            End
            Begin VB.Label Lab��żְҵ 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   3960
               TabIndex        =   145
               Top             =   720
               Width           =   1575
            End
            Begin VB.Label Lab��ż���� 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   3960
               TabIndex        =   144
               Top             =   240
               Width           =   1575
            End
            Begin VB.Label Lab���� 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   960
               TabIndex        =   143
               Top             =   720
               Width           =   1575
            End
            Begin VB.Label Lab��� 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   960
               TabIndex        =   142
               Top             =   240
               Width           =   1575
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "�Ƿ��飺"
               Height          =   180
               Index           =   0
               Left            =   120
               TabIndex        =   94
               Top             =   300
               Width           =   900
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "������ڣ�"
               Height          =   180
               Index           =   0
               Left            =   120
               TabIndex        =   93
               Top             =   760
               Width           =   900
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "��żְҵ��"
               Height          =   180
               Index           =   0
               Left            =   2760
               TabIndex        =   92
               Top             =   760
               Width           =   900
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "��ż�Ӵ������������"
               Height          =   180
               Index           =   0
               Left            =   5880
               TabIndex        =   91
               Top             =   240
               Width           =   1800
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "��ż����״����"
               Height          =   180
               Index           =   0
               Left            =   2760
               TabIndex        =   90
               Top             =   300
               Width           =   1260
            End
         End
         Begin VB.ComboBox Combo11 
            Height          =   300
            ItemData        =   "frmInfromation.frx":00A8
            Left            =   5040
            List            =   "frmInfromation.frx":00B5
            TabIndex        =   87
            Text            =   "����"
            Top             =   360
            Width           =   855
         End
         Begin VB.Frame Frame19 
            Caption         =   "�̾�ʷ"
            ForeColor       =   &H000080FF&
            Height          =   2175
            Index           =   0
            Left            =   6000
            TabIndex        =   73
            Top             =   1560
            Width           =   5175
            Begin VB.Label Lab��ʳϰ�� 
               BackColor       =   &H00FFFFFF&
               Height          =   375
               Index           =   0
               Left            =   120
               TabIndex        =   175
               Top             =   1680
               Width           =   4215
            End
            Begin VB.Label Lab������ 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   3360
               TabIndex        =   174
               Top             =   840
               Width           =   975
            End
            Begin VB.Label Lab������ 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   960
               TabIndex        =   173
               Top             =   1200
               Width           =   975
            End
            Begin VB.Label Lab����ʱ�� 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   960
               TabIndex        =   172
               Top             =   840
               Width           =   975
            End
            Begin VB.Label Lab���� 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   3360
               TabIndex        =   171
               Top             =   480
               Width           =   975
            End
            Begin VB.Label Lab���� 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   960
               TabIndex        =   169
               Top             =   480
               Width           =   975
            End
            Begin VB.Label Lab���Ƴ̶� 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   3360
               TabIndex        =   168
               Top             =   120
               Width           =   975
            End
            Begin VB.Label Lab���̶̳� 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   960
               TabIndex        =   167
               Top             =   120
               Width           =   975
            End
            Begin VB.Label Label99 
               AutoSize        =   -1  'True
               Caption         =   "���̶̳ȣ�"
               Height          =   180
               Index           =   0
               Left            =   120
               TabIndex        =   86
               Top             =   195
               Width           =   900
            End
            Begin VB.Label Label100 
               AutoSize        =   -1  'True
               Caption         =   "���Ƴ̶ȣ�"
               Height          =   180
               Index           =   0
               Left            =   2520
               TabIndex        =   85
               Top             =   195
               Width           =   900
            End
            Begin VB.Label Label102 
               AutoSize        =   -1  'True
               Caption         =   "֧/��"
               Height          =   180
               Index           =   0
               Left            =   2040
               TabIndex        =   84
               Top             =   1200
               Width           =   450
            End
            Begin VB.Label Label103 
               AutoSize        =   -1  'True
               Caption         =   "��������"
               Height          =   180
               Index           =   0
               Left            =   2640
               TabIndex        =   83
               Top             =   915
               Width           =   720
            End
            Begin VB.Label Label104 
               AutoSize        =   -1  'True
               Caption         =   "��������"
               Height          =   180
               Index           =   0
               Left            =   240
               TabIndex        =   82
               Top             =   1215
               Width           =   720
            End
            Begin VB.Label Label105 
               AutoSize        =   -1  'True
               Caption         =   "���䣺"
               Height          =   180
               Index           =   0
               Left            =   2640
               TabIndex        =   81
               Top             =   540
               Width           =   540
            End
            Begin VB.Label Label106 
               AutoSize        =   -1  'True
               Caption         =   "���䣺"
               Height          =   180
               Index           =   0
               Left            =   360
               TabIndex        =   80
               Top             =   540
               Width           =   540
            End
            Begin VB.Label Label107 
               AutoSize        =   -1  'True
               Caption         =   "����ʱ����"
               Height          =   180
               Index           =   0
               Left            =   120
               TabIndex        =   79
               Top             =   855
               Width           =   900
            End
            Begin VB.Label Label108 
               AutoSize        =   -1  'True
               Caption         =   "��"
               Height          =   180
               Index           =   0
               Left            =   2040
               TabIndex        =   78
               Top             =   480
               Width           =   180
            End
            Begin VB.Label Label109 
               AutoSize        =   -1  'True
               Caption         =   "��"
               Height          =   180
               Index           =   0
               Left            =   4440
               TabIndex        =   77
               Top             =   480
               Width           =   180
            End
            Begin VB.Label Label110 
               AutoSize        =   -1  'True
               Caption         =   "��"
               Height          =   180
               Index           =   0
               Left            =   2040
               TabIndex        =   76
               Top             =   840
               Width           =   180
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               Caption         =   "�����ס��������ʳϰ�ߡ��̾��Ⱥ�������"
               Height          =   180
               Index           =   1
               Left            =   120
               TabIndex        =   75
               Top             =   1440
               Width           =   3420
            End
            Begin VB.Label Label101 
               AutoSize        =   -1  'True
               Caption         =   "ML/��"
               Height          =   300
               Index           =   0
               Left            =   4440
               TabIndex        =   74
               Top             =   840
               Width           =   810
            End
         End
      End
      Begin VB.Frame freOrdinary 
         BorderStyle     =   0  'None
         Height          =   3735
         Left            =   -74640
         TabIndex        =   49
         Top             =   600
         Width           =   11175
         Begin VB.Frame Frame4 
            Caption         =   "����ʷ"
            ForeColor       =   &H000080FF&
            Height          =   1095
            Index           =   1
            Left            =   6000
            TabIndex        =   71
            Top             =   2520
            Width           =   5055
            Begin VB.Label Lab����ʷ 
               BackColor       =   &H00FFFFFF&
               Height          =   735
               Index           =   2
               Left            =   120
               TabIndex        =   189
               Top             =   240
               Width           =   4815
            End
         End
         Begin VB.Frame Frame19 
            Caption         =   "�̾�ʷ"
            ForeColor       =   &H000080FF&
            Height          =   1455
            Index           =   1
            Left            =   6000
            TabIndex        =   58
            Top             =   240
            Width           =   5055
            Begin VB.Label Lab������ 
               BackColor       =   &H00FFFFFF&
               Height          =   180
               Index           =   2
               Left            =   960
               TabIndex        =   188
               Top             =   1200
               Width           =   975
            End
            Begin VB.Label Lab������ 
               BackColor       =   &H00FFFFFF&
               Height          =   180
               Index           =   2
               Left            =   3360
               TabIndex        =   187
               Top             =   960
               Width           =   975
            End
            Begin VB.Label Lab����ʱ�� 
               BackColor       =   &H00FFFFFF&
               Height          =   180
               Index           =   2
               Left            =   960
               TabIndex        =   186
               Top             =   960
               Width           =   975
            End
            Begin VB.Label Lab���� 
               BackColor       =   &H00FFFFFF&
               Height          =   180
               Index           =   2
               Left            =   3120
               TabIndex        =   185
               Top             =   600
               Width           =   1095
            End
            Begin VB.Label Lab���� 
               BackColor       =   &H00FFFFFF&
               Height          =   180
               Index           =   2
               Left            =   840
               TabIndex        =   184
               Top             =   600
               Width           =   1095
            End
            Begin VB.Label Lab���Ƴ̶� 
               BackColor       =   &H00FFFFFF&
               Height          =   180
               Index           =   2
               Left            =   3360
               TabIndex        =   183
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label Lab���̶̳� 
               BackColor       =   &H00FFFFFF&
               Height          =   180
               Index           =   2
               Left            =   960
               TabIndex        =   182
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label Label110 
               AutoSize        =   -1  'True
               Caption         =   "��"
               Height          =   180
               Index           =   1
               Left            =   2040
               TabIndex        =   70
               Top             =   960
               Width           =   180
            End
            Begin VB.Label Label109 
               AutoSize        =   -1  'True
               Caption         =   "��"
               Height          =   180
               Index           =   1
               Left            =   4440
               TabIndex        =   69
               Top             =   600
               Width           =   180
            End
            Begin VB.Label Label108 
               AutoSize        =   -1  'True
               Caption         =   "��"
               Height          =   180
               Index           =   1
               Left            =   2040
               TabIndex        =   68
               Top             =   600
               Width           =   180
            End
            Begin VB.Label Label107 
               AutoSize        =   -1  'True
               Caption         =   "����ʱ����"
               Height          =   180
               Index           =   1
               Left            =   120
               TabIndex        =   67
               Top             =   960
               Width           =   900
            End
            Begin VB.Label Label106 
               AutoSize        =   -1  'True
               Caption         =   "���䣺"
               Height          =   180
               Index           =   1
               Left            =   360
               TabIndex        =   66
               Top             =   600
               Width           =   540
            End
            Begin VB.Label Label105 
               AutoSize        =   -1  'True
               Caption         =   "���䣺"
               Height          =   180
               Index           =   1
               Left            =   2640
               TabIndex        =   65
               Top             =   600
               Width           =   540
            End
            Begin VB.Label Label104 
               AutoSize        =   -1  'True
               Caption         =   "��������"
               Height          =   180
               Index           =   1
               Left            =   240
               TabIndex        =   64
               Top             =   1215
               Width           =   720
            End
            Begin VB.Label Label103 
               AutoSize        =   -1  'True
               Caption         =   "��������"
               Height          =   180
               Index           =   1
               Left            =   2640
               TabIndex        =   63
               Top             =   960
               Width           =   720
            End
            Begin VB.Label Label102 
               AutoSize        =   -1  'True
               Caption         =   "֧/��"
               Height          =   180
               Index           =   1
               Left            =   2040
               TabIndex        =   62
               Top             =   1200
               Width           =   450
            End
            Begin VB.Label Label101 
               AutoSize        =   -1  'True
               Caption         =   "ML/��"
               Height          =   180
               Index           =   1
               Left            =   4440
               TabIndex        =   61
               Top             =   960
               Width           =   450
            End
            Begin VB.Label Label100 
               AutoSize        =   -1  'True
               Caption         =   "���Ƴ̶ȣ�"
               Height          =   180
               Index           =   1
               Left            =   2520
               TabIndex        =   60
               Top             =   240
               Width           =   900
            End
            Begin VB.Label Label99 
               AutoSize        =   -1  'True
               Caption         =   "���̶̳ȣ�"
               Height          =   180
               Index           =   1
               Left            =   120
               TabIndex        =   59
               Top             =   240
               Width           =   900
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "����ʷ"
            ForeColor       =   &H000080FF&
            Height          =   615
            Index           =   2
            Left            =   120
            TabIndex        =   56
            Top             =   240
            Width           =   5775
            Begin VB.Label Lab��� 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   2
               Left            =   1320
               TabIndex        =   213
               Top             =   240
               Width           =   1695
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "�Ƿ��飺"
               Height          =   180
               Index           =   2
               Left            =   480
               TabIndex        =   57
               Top             =   300
               Width           =   900
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "����ʷ(����ż����ʷ)"
            ForeColor       =   &H000080FF&
            Height          =   2655
            Index           =   2
            Left            =   120
            TabIndex        =   50
            Top             =   960
            Width           =   5775
            Begin VB.Label Lab�쳣̥ 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   2
               Left            =   1680
               TabIndex        =   181
               Top             =   1080
               Width           =   1335
            End
            Begin VB.Label Lab���� 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   2
               Left            =   3960
               TabIndex        =   180
               Top             =   720
               Width           =   1335
            End
            Begin VB.Label Lab���� 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   2
               Left            =   1680
               TabIndex        =   179
               Top             =   720
               Width           =   1335
            End
            Begin VB.Label Lab��� 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   2
               Left            =   3960
               TabIndex        =   178
               Top             =   360
               Width           =   1335
            End
            Begin VB.Label Lab��Ů�� 
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   2
               Left            =   1680
               TabIndex        =   177
               Top             =   360
               Width           =   1335
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "������"
               Height          =   180
               Index           =   1
               Left            =   3480
               TabIndex        =   55
               Top             =   720
               Width           =   540
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "�����"
               Height          =   180
               Index           =   1
               Left            =   3480
               TabIndex        =   54
               Top             =   360
               Width           =   540
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               Caption         =   "������"
               Height          =   180
               Index           =   1
               Left            =   1080
               TabIndex        =   53
               Top             =   720
               Width           =   540
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               Caption         =   "������Ů��Ŀ��"
               Height          =   180
               Index           =   1
               Left            =   480
               TabIndex        =   52
               Top             =   360
               Width           =   1260
            End
            Begin VB.Label Label�쳣̥ 
               AutoSize        =   -1  'True
               Caption         =   "�쳣̥��"
               Height          =   180
               Left            =   840
               TabIndex        =   51
               Top             =   1080
               Width           =   720
            End
         End
         Begin VB.Label Lab������ 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   2
            Left            =   6120
            TabIndex        =   231
            Top             =   2040
            Width           =   4935
         End
         Begin VB.Label Label36 
            Caption         =   "�����أ�"
            Height          =   255
            Left            =   6120
            TabIndex        =   230
            Top             =   1800
            Width           =   975
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "���һ�����¼��"
         ForeColor       =   &H000080FF&
         Height          =   3255
         Left            =   -74280
         TabIndex        =   124
         Top             =   1140
         Width           =   3615
         Begin VB.Label Lab���� 
            BackColor       =   &H80000009&
            Height          =   255
            Left            =   960
            TabIndex        =   233
            Top             =   2760
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label Label37 
            Caption         =   "����"
            Height          =   375
            Left            =   240
            TabIndex        =   232
            Top             =   2760
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label Lab����ѹ 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   960
            TabIndex        =   218
            Top             =   2280
            Width           =   1335
         End
         Begin VB.Label Lab����ѹ 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   960
            TabIndex        =   217
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label Lab���� 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   960
            TabIndex        =   216
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label Lab��� 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   960
            TabIndex        =   215
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label LabӪ�� 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   960
            TabIndex        =   214
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label55 
            Caption         =   "mmHg"
            Height          =   255
            Left            =   2400
            TabIndex        =   133
            Top             =   1800
            Width           =   375
         End
         Begin VB.Label Label53 
            Caption         =   "����ѹ"
            Height          =   255
            Left            =   240
            TabIndex        =   132
            Top             =   1800
            Width           =   615
         End
         Begin VB.Label Label51 
            Caption         =   "kg"
            Height          =   255
            Left            =   2400
            TabIndex        =   131
            Top             =   1320
            Width           =   375
         End
         Begin VB.Label Label42 
            Caption         =   "����"
            Height          =   255
            Left            =   240
            TabIndex        =   130
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label Label35 
            Caption         =   "cm"
            Height          =   255
            Index           =   0
            Left            =   2400
            TabIndex        =   129
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label34 
            Caption         =   "���"
            Height          =   255
            Left            =   240
            TabIndex        =   128
            Top             =   840
            Width           =   495
         End
         Begin VB.Label Label33 
            Caption         =   "Ӫ��"
            Height          =   255
            Left            =   240
            TabIndex        =   127
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label56 
            Caption         =   "mmHg"
            Height          =   255
            Left            =   2400
            TabIndex        =   126
            Top             =   2280
            Width           =   375
         End
         Begin VB.Label Label57 
            Caption         =   "����ѹ"
            Height          =   375
            Left            =   240
            TabIndex        =   125
            Top             =   2280
            Width           =   615
         End
      End
      Begin VB.TextBox ctxtOther 
         Height          =   495
         Left            =   -74640
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   123
         Top             =   5400
         Width           =   10935
      End
      Begin VB.Frame Frame1 
         Caption         =   "�¾�ʷ"
         ForeColor       =   &H000080FF&
         Height          =   975
         Left            =   -74640
         TabIndex        =   116
         Top             =   4320
         Width           =   5775
         Begin VB.Label Labͣ������ 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   3240
            TabIndex        =   166
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Labĩ���¾� 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   960
            TabIndex        =   165
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Lab���� 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   4080
            TabIndex        =   164
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Lab���� 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   2400
            TabIndex        =   163
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Lab���� 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   600
            TabIndex        =   162
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "������"
            Height          =   180
            Index           =   2
            Left            =   120
            TabIndex        =   122
            Top             =   240
            Width           =   540
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "���ڣ�"
            Height          =   180
            Index           =   2
            Left            =   1920
            TabIndex        =   121
            Top             =   240
            Width           =   540
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "���ڣ�"
            Height          =   180
            Index           =   2
            Left            =   3600
            TabIndex        =   120
            Top             =   240
            Width           =   540
         End
         Begin VB.Label Label4 
            Caption         =   "Label4"
            Height          =   15
            Index           =   2
            Left            =   720
            TabIndex        =   119
            Top             =   720
            Width           =   135
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "ĩ���¾���"
            Height          =   180
            Index           =   2
            Left            =   120
            TabIndex        =   118
            Top             =   600
            Width           =   900
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "ͣ�����䣺"
            Height          =   180
            Index           =   2
            Left            =   2400
            TabIndex        =   117
            Top             =   600
            Width           =   900
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "����ʷ"
         ForeColor       =   &H000080FF&
         Height          =   975
         Left            =   -68880
         TabIndex        =   114
         Top             =   4320
         Width           =   5175
         Begin VB.Label Lab����ʷ 
            BackColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   240
            TabIndex        =   176
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label Label27 
            Caption         =   "��ʾ:�����������Ŵ��Լ�����ѪҺ�������򲡡���Ѫѹ�����񾭾����Լ�������������˲���"
            Height          =   615
            Left            =   2520
            TabIndex        =   115
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "����¼��"
         ForeColor       =   &H000080FF&
         Height          =   855
         Left            =   -74280
         TabIndex        =   110
         Top             =   4800
         Width           =   5295
         Begin VB.Label Lab���� 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   840
            TabIndex        =   219
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label64 
            Caption         =   "��/��"
            Height          =   255
            Left            =   2400
            TabIndex        =   113
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label65 
            Caption         =   "����"
            Height          =   375
            Left            =   240
            TabIndex        =   112
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label66 
            Height          =   255
            Left            =   3120
            TabIndex        =   111
            Top             =   360
            Width           =   855
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid cgrd֢״ 
         Height          =   2175
         Left            =   -74880
         TabIndex        =   134
         Top             =   480
         Width           =   12375
         _cx             =   21828
         _cy             =   3836
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
         FormatString    =   $"frmInfromation.frx":00CB
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
         Height          =   4335
         Left            =   -74880
         TabIndex        =   135
         Top             =   480
         Width           =   12375
         _cx             =   21828
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
      Begin VSFlex8Ctl.VSFlexGrid cgrdzzxw 
         Height          =   4335
         Left            =   -74880
         TabIndex        =   136
         Top             =   480
         Width           =   12255
         _cx             =   21616
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
         FormatString    =   $"frmInfromation.frx":0162
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
         Height          =   4815
         Left            =   120
         TabIndex        =   137
         Top             =   480
         Width           =   12375
         _cx             =   21828
         _cy             =   8493
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
         FormatString    =   $"frmInfromation.frx":01E2
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
      Begin VB.Label Lab֢״ 
         BackColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   -74760
         TabIndex        =   222
         Top             =   480
         Width           =   5055
      End
      Begin VB.Label Lab��ʷ 
         BackColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   -74760
         TabIndex        =   221
         Top             =   600
         Width           =   3975
      End
      Begin VB.Label Labְҵʷ 
         BackColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   120
         TabIndex        =   220
         Top             =   600
         Width           =   5415
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "������"
         Height          =   180
         Index           =   0
         Left            =   -74640
         TabIndex        =   138
         Top             =   5460
         Width           =   540
      End
   End
   Begin VB.Label Label97 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label"
      Height          =   255
      Left            =   9720
      TabIndex        =   170
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label Label39 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label"
      Height          =   255
      Left            =   1560
      TabIndex        =   148
      Top             =   4680
      Width           =   735
   End
End
Attribute VB_Name = "frmInfromation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'�������
Private Sub Form_Load()
    
    Dim Index As Integer
    Dim baseresql As Object
    Set baseresql = dafuncGetData("select * From ְҵ�����_�����Ա������Ϣ�� where ϵͳ���='" & mstrϵͳ��� & "'")
    Lab���.Caption = baseresql("ϵͳ���")
    Lab����.Caption = baseresql("����")
    Lab�Ա�.Caption = baseresql("�Ա�")
    Lab����.Caption = baseresql("����")
    lab��λ.Caption = baseresql("��λ����")
    Lab�ֹ���.Caption = baseresql("�ֹ���")
    Lab��ְ��.Caption = baseresql("ְ���ְ��")
    LabelΣ������.Caption = baseresql("Σ������")
    
    
        '��ȡ��Ƭ
    Dim lobjRec As Object
    Set lobjRec = CreateObject("ְҵ������.clspersonexamed")
    lobjRec.ϵͳ��� = Trim(Lab���.Caption)
    Picture2.Picture = lobjRec.��Ƭ
    Picture2.Visible = True
     
    '��ʼ������ͨ�ö���ÿ�������Ӧһ������ͨ�ö���ʵ�������ɻ��ã��м��мǣ���
'    Set mobjGUI = New cls����ͨ�ö���
      Dim resql As Object
    Set resql = dafuncGetData("select * From ְҵ�����_�����Ա������Ϣ�� where ϵͳ���='" & mstrϵͳ��� & "'")
    
    If resql("��������") = "��ͨ���" Or resql("��������") = "ְҵ����" Then
        freOrdinary.Visible = True
        freNuclear.Visible = False
        freRadiation.Visible = False
        Index = 2
        
    ElseIf resql("��������") = "8023����" Or resql("��������") = "��˲���" Or resql("��������") = "��˲���YK" Then
        freNuclear.Visible = True
        freOrdinary.Visible = False
        freRadiation.Visible = False
        Index = 1
    ElseIf resql("��������") = "���佡��" Then
        freRadiation.Visible = True
        freOrdinary.Visible = False
        freNuclear.Visible = False
        Index = 0
    End If
    Label�������.Caption = resql("��������")
    If Len(resql("�绰����")) = 11 Then
        resql("�绰����") = Left(resql("�绰����"), 3) & "-" & Mid(resql("�绰����"), 4, 4) & "-" & Mid(resql("�绰����"), 8, 4)
    End If
    Label�绰.Caption = IIf(resql("�绰����") = "", "��", resql("�绰����"))
    Label�绰.FontSize = 14
        
    
    '����ʱ����ʾ��һ������
    SSTab1.Tab = 0
    Dim detsql As Object
    Set detsql = dafuncGetData("select * From ְҵ�����_��������ʷ�� where ϵͳ���='" & mstrϵͳ��� & "'")
    If Index = 0 Then   '���佡��
        
    Lab���(Index).Caption = detsql("�Ƿ���")
    Lab��ż����(Index).Caption = detsql("��ż����״��")
    Lab����(Index).Caption = detsql("�������")
    Lab��żְҵ(Index).Caption = detsql("��żְҵ")
    Lab��ż����(Index).Caption = detsql("��ż�Ӵ�����")
    Lab�д�(Index).Caption = detsql("�д�")
    Lab���(Index).Caption = detsql("���")
    Lab���(Index).Caption = detsql("���")
    Lab����(Index).Caption = detsql("����")
    Lab����(Index).Caption = detsql("��Ȼ����")
    Lab��̥(Index).Caption = detsql("��̥")
    Lab��̥(Index).Caption = detsql("��̥")
    Lab��λ����(Index).Caption = detsql("��λ����")
    Lab����ԭ��(Index).Caption = detsql("���в���ԭ��")
    Lab��Ů��(Index).Caption = detsql("������Ů��Ŀ")
    LabŮ��(Index).Caption = detsql("����Ů��")
    Lab�к���������(Index).Caption = detsql("�к���������")
    LabŮ����������(Index).Caption = detsql("Ů����������")
    Lab��Ů����(Index).Caption = detsql("��Ů����״��")
    Lab���̶̳�(Index).Caption = detsql("���̶̳�")
    Lab���Ƴ̶�(Index).Caption = detsql("���Ƴ̶�")
    Lab����(Index).Caption = detsql("����")
    Lab����(Index).Caption = detsql("����")
    Lab����ʱ��(Index).Caption = detsql("����ʱ��")
    Lab������(Index).Caption = detsql("������")
    Lab������(Index).Caption = detsql("������")
    Lab��ʳϰ��(Index).Caption = detsql("�������")
    End If
    
    
    If Index = 1 Then    '�˲���
    Lab���̶̳�(Index).Caption = detsql("���̶̳�")
    Lab���Ƴ̶�(Index).Caption = detsql("���Ƴ̶�")
    Lab����ʱ��(Index).Caption = detsql("����ʱ��")
    Lab���(Index).Caption = detsql("�Ƿ���")
    Lab��ż����(Index).Caption = detsql("��ż����״��")
    Lab����(Index).Caption = detsql("�������")
    Lab��żְҵ(Index).Caption = detsql("��żְҵ")
    Lab��ż����(Index).Caption = detsql("��ż�Ӵ�����")
    Lab�д�(Index).Caption = detsql("�д�")
    Lab���(Index).Caption = detsql("���")
    Lab���(Index).Caption = detsql("���")
    Lab����(Index).Caption = detsql("����")
    Lab����(Index).Caption = detsql("��Ȼ����")
    Lab��̥(Index).Caption = detsql("��̥")
    Lab��̥(Index).Caption = detsql("��̥")
    Lab����ԭ��(Index).Caption = detsql("���в���ԭ��")
    Lab��Ů��(Index).Caption = detsql("������Ů��Ŀ")
    LabŮ��(Index).Caption = detsql("����Ů��")
    Lab�к���������(Index).Caption = detsql("�к���������")
    LabŮ����������(Index).Caption = detsql("Ů����������")
    Lab��Ů����(Index).Caption = detsql("��Ů����״��")
    Lab����(Index).Caption = detsql("����")
    Lab����(Index).Caption = detsql("����")
    Lab������(Index).Caption = detsql("������")
    Lab������(Index).Caption = detsql("������")
        If Not IsNull(detsql("�������")) Then      '�жϾ�ס�ص���Ϣ
        Lab��ʳϰ��(Index).Caption = detsql("�������")
        End If
    End If
    
    
    If Index = 2 Then   'ְҵ����
    Dim sql As Object
    Set sql = dafuncGetData("select * From ְҵ�����_�����Ա������Ϣ�� where ϵͳ���='" & mstrϵͳ��� & "'")
    Lab������(Index).Caption = sql("������")
    
    Lab���(Index).Caption = detsql("�Ƿ���")
    Lab��Ů��(Index).Caption = detsql("������Ů��Ŀ")
    Lab���(Index).Caption = detsql("���")
    Lab����(Index).Caption = detsql("����")
'    Lab����(Index).Caption = detsql("����")
    Lab�쳣̥(Index).Caption = detsql("�쳣̥")
    Lab���̶̳�(Index).Caption = detsql("���̶̳�")
    Lab���Ƴ̶�(Index).Caption = detsql("���Ƴ̶�")
    Lab����(Index).Caption = detsql("����")
    Lab����(Index).Caption = detsql("����")
    Lab������(Index).Caption = detsql("������")
    Lab������(Index).Caption = detsql("������")
    Lab����ʱ��(Index).Caption = detsql("����ʱ��")
    Lab����ʷ(Index).Caption = detsql("����ʷ")
    End If
    
    Lab����.Caption = detsql("����")
    Lab����.Caption = detsql("����")
    Lab����.Caption = detsql("����")
    Labĩ���¾�.Caption = detsql("ĩ���¾�")
    Labͣ������.Caption = detsql("ͣ������")
    Lab����ʷ.Caption = detsql("����ʷ")


'ְҵʷ
    SSTab1.Tab = 1
   Dim lstrWhere As String
    Dim lstrSql As String
        
        lstrSql = "select * From ְҵ�����_ְҵʷ�� where ϵͳ���='" & mstrϵͳ��� & "'"
        dasubSetQueryTimeout 6000
        Set lobjRec = dafuncGetData(lstrSql)
        
        If Not lobjRec.EOF Then
            With cgrdְҵʷ
                Set .DataSource = lobjRec
                clblInfo = .rows - 1
                .Col = 0
                .Sort = flexSortGenericDescending
                .AutoSize 0, .cols - 1, 0, 0
                .ExplorerBar = flexExSort
                .DataMode = flexDMFree
            End With
            
            If cgrdְҵʷ.rows > 1 Then
                Dim i As Long
                Set mcolIndex = New Collection
                For i = 0 To cgrdְҵʷ.cols - 1
                    mcolIndex.Add i, cgrdְҵʷ.TextMatrix(0, i)
                Next
            End If

        Else
            '��ְҵʷû������ʱ����ʾ����ְҵʷ��  2015-10-26
            cgrdְҵʷ.Visible = False
            Labְҵʷ.Caption = "��ְҵʷ"
            Labְҵʷ.FontSize = 22
            cgrdְҵʷ.rows = 1
        End If
    
    
  '������ʷ
 SSTab1.Tab = 2
        lstrSql = "select * From ְҵ�����_������ʷ�� where ϵͳ���='" & mstrϵͳ��� & "'"
        dasubSetQueryTimeout 6000
        Set lobjRec = dafuncGetData(lstrSql)
        If Not lobjRec.EOF Then
            With cgrd��ʷ
                Set .DataSource = lobjRec
                clblInfo = .rows - 1
                .Col = 0
                .Sort = flexSortGenericDescending
                .AutoSize 0, .cols - 1, 0, 0
                .ExplorerBar = flexExSort
                .DataMode = flexDMFree
            End With
              
            If cgrd��ʷ.rows > 1 Then
                Set mcolIndex = New Collection
                For i = 0 To cgrd��ʷ.cols - 1
                    mcolIndex.Add i, cgrd��ʷ.TextMatrix(0, i)
                Next
            End If
                
        Else
         '��ְҵʷû������ʱ����ʾ���޼�����ʷ��  2015-10-26
            cgrd��ʷ.Visible = False
            Lab��ʷ.Caption = "�޼�����ʷ"
            Lab��ʷ.FontSize = 22
            cgrd��ʷ.rows = 1
        End If
    
    '�Ծ�֢״
   SSTab1.Tab = 3
'        lstrSql = "select ϵͳ���,���,֢״,�̶�,����ʱ�� From ְҵ�����_�Ծ�֢״�� where ϵͳ���='" & mstrϵͳ��� & "'"
'       lstrSql = "select * From ְҵ�����_�Ծ�֢״�� where ϵͳ���='" & mstrϵͳ��� & "'"
   Dim symresql As Object
    Set symresql = dafuncGetData("select �������� From ְҵ�����_�����Ա������Ϣ�� where ϵͳ���='" & mstrϵͳ��� & "'")
    If symresql("��������") = "8023����" Or resql("��������") = "��˲���" Or resql("��������") = "��˲���YK" Then
      lstrSql = "select ϵͳ���,���,֢״,�̶�,����ʱ�� From ְҵ�����_�Ծ�֢״�� where ϵͳ���='" & mstrϵͳ��� & "'and ����ʱ��!='' "
      dasubSetQueryTimeout 6000
        Set lobjRec = dafuncGetData(lstrSql)
        If Not lobjRec.EOF Then
            With cgrd֢״
                Set .DataSource = lobjRec
                clblInfo = .rows - 1
                .Col = 0
                .Sort = flexSortGenericDescending
                .AutoSize 0, .cols - 1, 0, 0
                .ExplorerBar = flexExSort
                .DataMode = flexDMFree
            End With
            
            If cgrd֢״.rows > 1 Then
                Set mcolIndex = New Collection
                For i = 0 To cgrd֢״.cols - 1
                    mcolIndex.Add i, cgrd֢״.TextMatrix(0, i)
                Next
             End If
                         
        Else
            cgrd֢״.Visible = False
            Lab֢״.Caption = "���Ծ�֢״"
            Lab֢״.FontSize = 22
            cgrd֢״.rows = 1
        End If
    Else
    '��Ҫ���ҽʦ  2015-10-26
    lstrSql = "select ϵͳ���,���,֢״,�̶�,����ʱ�� From ְҵ�����_�Ծ�֢״�� where ϵͳ���='" & mstrϵͳ��� & "'"
          dasubSetQueryTimeout 6000
        Set lobjRec = dafuncGetData(lstrSql)
        If Not lobjRec.EOF Then
            With cgrd֢״
                Set .DataSource = lobjRec
                clblInfo = .rows - 1
                .Col = 0
                .Sort = flexSortGenericDescending
                .AutoSize 0, .cols - 1, 0, 0
                .ExplorerBar = flexExSort
                .DataMode = flexDMFree
            End With
            
            If cgrd֢״.rows > 1 Then
                Set mcolIndex = New Collection
                For i = 0 To cgrd֢״.cols - 1
                    mcolIndex.Add i, cgrd֢״.TextMatrix(0, i)
                Next
             End If
            Else
            cgrd֢״.Visible = False
            Lab֢״.Caption = "���Ծ�֢״"
            Lab֢״.FontSize = 22
            cgrd֢״.rows = 1
        End If
    End If
        
        '���һ�����
    SSTab1.Tab = 4
        Set detsql = dafuncGetData("select ����� From ְҵ�����_�����Ϣ_�ܼ��߸�����Ϣ¼��� where ϵͳ���='" & mstrϵͳ��� & "'and �����Ŀ='13017'")
        LabӪ��.Caption = detsql("�����")
        Set detsql = dafuncGetData("select ����� From ְҵ�����_�����Ϣ_�ܼ��߸�����Ϣ¼��� where ϵͳ���='" & mstrϵͳ��� & "'and �����Ŀ='13018'")
        Lab���.Caption = detsql("�����")
        Set detsql = dafuncGetData("select ����� From ְҵ�����_�����Ϣ_�ܼ��߸�����Ϣ¼��� where ϵͳ���='" & mstrϵͳ��� & "'and �����Ŀ='13019'")
        Lab����.Caption = detsql("�����")
        Set detsql = dafuncGetData("select ����� From ְҵ�����_�����Ϣ_�ܼ��߸�����Ϣ¼��� where ϵͳ���='" & mstrϵͳ��� & "'and �����Ŀ='13020'")
        Lab����ѹ.Caption = detsql("�����")
        Set detsql = dafuncGetData("select ����� From ְҵ�����_�����Ϣ_�ܼ��߸�����Ϣ¼��� where ϵͳ���='" & mstrϵͳ��� & "'and �����Ŀ='13021'")
        Lab����ѹ.Caption = detsql("�����")
        Set detsql = dafuncGetData("select ����� From ְҵ�����_�����Ϣ_�ڿ� where ϵͳ���='" & mstrϵͳ��� & "'and �����Ŀ='02002'")
        If IsNull(detsql("�����")) Then
        Lab����.Caption = "δ¼��"
        Else
        Lab����.Caption = detsql("�����")
        End If
        Set detsql = dafuncGetData("select ����� From ְҵ�����_�����Ϣ_�ڿ� where ϵͳ���='" & mstrϵͳ��� & "'and �����Ŀ='02019'")
        If detsql.RecordCount > 0 Then
        Label37.Visible = True
        Lab����.Visible = True
           If detsql("�����") = "" Then
           Lab����.Caption = "/"
           Else
           Lab����.Caption = detsql("�����")
           End If
        End If
    '���ص�����ѡ�  2015-10-26
        SSTab1.Tab = 5
        SSTab1.TabVisible(5) = False
        
End Sub


