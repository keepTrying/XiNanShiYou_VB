VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frm��ӡ��ʽ���� 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "����֤��ӡ��ʽ����"
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   10290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton ccmdCopy 
      Caption         =   "����(&C)"
      Height          =   400
      Left            =   3600
      TabIndex        =   69
      ToolTipText     =   "���Ƶ�ǰ�����õĲ�����Ȼ�����ѡ��������ʽ�����ճ��������ǰ���Ƚ��б��档"
      Top             =   7560
      Width           =   1215
   End
   Begin VB.CommandButton ccmdPaste 
      Caption         =   "ճ��(&P)"
      Enabled         =   0   'False
      Height          =   400
      Left            =   4920
      TabIndex        =   68
      ToolTipText     =   "���ղŸ��ƵĴ�ӡ����ճ������ǰ��ѡ��ʽ��ͬʱ���浽ϵͳ��"
      Top             =   7560
      Width           =   1215
   End
   Begin VB.CommandButton ccmdExit 
      Caption         =   "����(&X)"
      Height          =   400
      Left            =   8880
      TabIndex        =   66
      Top             =   7560
      Width           =   1215
   End
   Begin VB.CommandButton ccmdPreview 
      Caption         =   "�Դ�(&P)"
      Height          =   400
      Left            =   2280
      TabIndex        =   65
      ToolTipText     =   "�Դ�ǰ���Ƚ��б���"
      Top             =   7560
      Width           =   1215
   End
   Begin VB.CommandButton ccmdSave 
      Caption         =   "����(&S)"
      Height          =   400
      Left            =   960
      TabIndex        =   64
      Top             =   7560
      Width           =   1215
   End
   Begin VB.TextBox ctxtBase 
      Height          =   300
      Index           =   29
      Left            =   9240
      TabIndex        =   62
      Tag             =   "��Ƭ��"
      Top             =   4800
      Width           =   615
   End
   Begin VB.TextBox ctxtBase 
      Height          =   300
      Index           =   28
      Left            =   9240
      TabIndex        =   60
      Tag             =   "��Ƭ��"
      Top             =   4440
      Width           =   615
   End
   Begin VB.TextBox ctxtBase 
      Height          =   300
      Index           =   27
      Left            =   9240
      TabIndex        =   58
      Tag             =   "��Ƭy"
      Top             =   4080
      Width           =   615
   End
   Begin VB.TextBox ctxtBase 
      Height          =   300
      Index           =   26
      Left            =   9240
      TabIndex        =   56
      Tag             =   "��Ƭx"
      Top             =   3720
      Width           =   615
   End
   Begin VB.TextBox ctxtBase 
      Height          =   300
      Index           =   25
      Left            =   9240
      TabIndex        =   54
      Tag             =   "��Ƭ��y2"
      Top             =   5880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox ctxtBase 
      Height          =   300
      Index           =   24
      Left            =   9240
      TabIndex        =   52
      Tag             =   "��Ƭ��x2"
      Top             =   5520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame Frame3 
      Height          =   135
      Left            =   120
      TabIndex        =   51
      Top             =   7200
      Width           =   10815
   End
   Begin VB.TextBox ctxtBase 
      Height          =   300
      Index           =   23
      Left            =   9240
      TabIndex        =   49
      Tag             =   "��Ƭ��y1"
      Top             =   3240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox ctxtBase 
      Height          =   300
      Index           =   22
      Left            =   9240
      TabIndex        =   47
      Tag             =   "��Ƭ��x1"
      Top             =   2880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox ctxtBase 
      Height          =   300
      Index           =   21
      Left            =   5280
      TabIndex        =   46
      Tag             =   "��֤��λy"
      Top             =   6360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox ctxtBase 
      Height          =   300
      Index           =   20
      Left            =   3480
      TabIndex        =   44
      Tag             =   "��֤��λx"
      Top             =   6360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox ctxtBase 
      Height          =   300
      Index           =   19
      Left            =   1200
      TabIndex        =   42
      Tag             =   "�������"
      Top             =   6420
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox ctxtBase 
      Height          =   300
      Index           =   18
      Left            =   1200
      TabIndex        =   40
      Tag             =   "�����"
      Top             =   6060
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox ctxtBase 
      Height          =   300
      Index           =   17
      Left            =   1200
      TabIndex        =   38
      Tag             =   "����"
      Top             =   5700
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox ctxtBase 
      Height          =   300
      Index           =   16
      Left            =   7080
      TabIndex        =   36
      Tag             =   "����"
      Top             =   6360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox ctxtBase 
      Height          =   300
      Index           =   15
      Left            =   7320
      TabIndex        =   34
      Tag             =   "�Ա�"
      Top             =   2400
      Width           =   615
   End
   Begin VB.TextBox ctxtBase 
      Height          =   300
      Index           =   14
      Left            =   1200
      TabIndex        =   32
      Tag             =   "����"
      Top             =   5340
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox ctxtBase 
      Height          =   300
      Index           =   13
      Left            =   1200
      TabIndex        =   30
      Tag             =   "����֤��"
      Top             =   4980
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   240
      TabIndex        =   29
      Top             =   720
      Width           =   9855
   End
   Begin VB.ComboBox ccmbֽ�� 
      Height          =   300
      ItemData        =   "frm��ӡ��ʽ����.frx":0000
      Left            =   1200
      List            =   "frm��ӡ��ʽ����.frx":0002
      TabIndex        =   14
      Text            =   "ccmbֽ��"
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox ctxtBase 
      Height          =   300
      Index           =   3
      Left            =   1080
      TabIndex        =   13
      Tag             =   "�����ʼ"
      Top             =   3900
      Width           =   735
   End
   Begin VB.TextBox ctxtBase 
      Height          =   300
      Index           =   4
      Left            =   1080
      TabIndex        =   12
      Tag             =   "������ʼ"
      Top             =   2580
      Width           =   855
   End
   Begin VB.TextBox ctxtBase 
      Height          =   300
      Index           =   1
      Left            =   5880
      TabIndex        =   11
      Tag             =   "���ź���"
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox ctxtBase 
      Height          =   300
      Index           =   0
      Left            =   4080
      TabIndex        =   10
      Tag             =   "�����ݼ��"
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox ctxtBase 
      Height          =   300
      Index           =   2
      Left            =   1080
      TabIndex        =   9
      Tag             =   "�м��"
      Top             =   4320
      Width           =   615
   End
   Begin VB.TextBox ctxtBase 
      Height          =   300
      Index           =   7
      Left            =   2160
      TabIndex        =   8
      Tag             =   "����"
      Text            =   "����"
      Top             =   1980
      Width           =   975
   End
   Begin VB.TextBox ctxtBase 
      Height          =   300
      Index           =   8
      Left            =   4200
      TabIndex        =   7
      Tag             =   "�����С"
      Text            =   "10"
      Top             =   1980
      Width           =   375
   End
   Begin VB.TextBox ctxtBase 
      Height          =   300
      Index           =   9
      Left            =   6960
      TabIndex        =   6
      Tag             =   "��������"
      Text            =   "����"
      Top             =   1680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox ctxtBase 
      Height          =   300
      Index           =   10
      Left            =   8640
      TabIndex        =   5
      Tag             =   "���������С"
      Text            =   "14"
      Top             =   1680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox ctxtBase 
      Height          =   300
      Index           =   11
      Left            =   7200
      TabIndex        =   4
      Tag             =   "����֤����x"
      Text            =   "0"
      Top             =   2040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox ctxtBase 
      Height          =   300
      Index           =   12
      Left            =   8640
      TabIndex        =   3
      Tag             =   "����֤����y"
      Text            =   "0"
      Top             =   2040
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox ccmb��ʽ 
      Height          =   300
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   2535
   End
   Begin MSComDlg.CommonDialog ccmdFile 
      Left            =   7200
      Top             =   7080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   2520
      TabIndex        =   27
      Top             =   2820
      Width           =   5015
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2610
         Left            =   0
         Picture         =   "frm��ӡ��ʽ����.frx":0004
         ScaleHeight     =   2580
         ScaleWidth      =   5010
         TabIndex        =   70
         Top             =   600
         Width           =   5040
         Begin VB.Line Line3 
            BorderColor     =   &H00000080&
            BorderWidth     =   2
            X1              =   0
            X2              =   360
            Y1              =   0
            Y2              =   300
         End
         Begin VB.Line Line4 
            BorderColor     =   &H00000080&
            BorderWidth     =   2
            X1              =   1320
            X2              =   1200
            Y1              =   180
            Y2              =   0
         End
         Begin VB.Line Line15 
            BorderColor     =   &H00000080&
            BorderWidth     =   2
            X1              =   2400
            X2              =   4080
            Y1              =   720
            Y2              =   -480
         End
         Begin VB.Line Line16 
            BorderColor     =   &H00000080&
            BorderWidth     =   2
            Visible         =   0   'False
            X1              =   2400
            X2              =   3840
            Y1              =   1200
            Y2              =   2640
         End
         Begin VB.Line Line20 
            BorderColor     =   &H00000080&
            BorderWidth     =   2
            Visible         =   0   'False
            X1              =   480
            X2              =   360
            Y1              =   2280
            Y2              =   2520
         End
         Begin VB.Line Line22 
            BorderColor     =   &H00000080&
            BorderWidth     =   2
            Visible         =   0   'False
            X1              =   3620
            X2              =   5040
            Y1              =   260
            Y2              =   -100
         End
         Begin VB.Line Line23 
            BorderColor     =   &H00000080&
            BorderWidth     =   2
            Visible         =   0   'False
            X1              =   4720
            X2              =   5040
            Y1              =   1720
            Y2              =   2040
         End
         Begin VB.Line Line25 
            BorderColor     =   &H00000080&
            BorderWidth     =   2
            X1              =   3670
            X2              =   4995
            Y1              =   340
            Y2              =   660
         End
         Begin VB.Line Line28 
            BorderColor     =   &H00000080&
            BorderWidth     =   2
            X1              =   4760
            X2              =   5040
            Y1              =   1240
            Y2              =   1240
         End
         Begin VB.Line Line29 
            BorderColor     =   &H00000080&
            BorderWidth     =   2
            X1              =   4440
            X2              =   5040
            Y1              =   1680
            Y2              =   1560
         End
         Begin VB.Line Line13 
            BorderColor     =   &H00000080&
            BorderWidth     =   2
            X1              =   360
            X2              =   -120
            Y1              =   645
            Y2              =   645
         End
         Begin VB.Line Line17 
            BorderColor     =   &H00000080&
            BorderWidth     =   2
            Visible         =   0   'False
            X1              =   -120
            X2              =   360
            Y1              =   1005
            Y2              =   1005
         End
         Begin VB.Line Line18 
            BorderColor     =   &H00000080&
            BorderWidth     =   2
            Visible         =   0   'False
            X1              =   360
            X2              =   -120
            Y1              =   1365
            Y2              =   1365
         End
         Begin VB.Line Line19 
            BorderColor     =   &H00000080&
            BorderWidth     =   2
            Visible         =   0   'False
            X1              =   360
            X2              =   -120
            Y1              =   1725
            Y2              =   1725
         End
         Begin VB.Line Line11 
            BorderColor     =   &H00000080&
            BorderWidth     =   2
            Visible         =   0   'False
            X1              =   360
            X2              =   0
            Y1              =   360
            Y2              =   360
         End
      End
      Begin VB.Line Line31 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         X1              =   1200
         X2              =   840
         Y1              =   600
         Y2              =   0
      End
      Begin VB.Line Line34 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Visible         =   0   'False
         X1              =   4960
         X2              =   4680
         Y1              =   390
         Y2              =   520
      End
      Begin VB.Line Line32 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         X1              =   3480
         X2              =   3960
         Y1              =   600
         Y2              =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Visible         =   0   'False
         X1              =   2640
         X2              =   3000
         Y1              =   240
         Y2              =   0
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "��  ��  ֤  ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1320
         TabIndex        =   28
         Top             =   240
         Width           =   2040
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��쵥λy��"
      Height          =   180
      Index           =   23
      Left            =   4320
      TabIndex        =   71
      Top             =   6480
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "xΪ�������꣬������ڡ������ʼ�������λ�ã�yΪ�������꣬������ڡ�������ʼ�������λ��"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   240
      TabIndex        =   67
      Top             =   6840
      Width           =   9915
   End
   Begin VB.Line Line30 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      X1              =   3120
      X2              =   3360
      Y1              =   2280
      Y2              =   2820
   End
   Begin VB.Line Line27 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      Index           =   1
      X1              =   8280
      X2              =   7560
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line Line26 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      Index           =   1
      X1              =   8280
      X2              =   7560
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��Ƭ��"
      Height          =   180
      Index           =   31
      Left            =   8280
      TabIndex        =   63
      Top             =   4920
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��Ƭ�ߣ�"
      Height          =   180
      Index           =   30
      Left            =   8280
      TabIndex        =   61
      Top             =   4560
      Width           =   720
   End
   Begin VB.Line Line27 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      Index           =   0
      X1              =   8280
      X2              =   7560
      Y1              =   4320
      Y2              =   4080
   End
   Begin VB.Line Line26 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      Index           =   0
      X1              =   8250
      X2              =   7560
      Y1              =   3945
      Y2              =   4080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��Ƭ Y��"
      Height          =   180
      Index           =   29
      Left            =   8280
      TabIndex        =   59
      Top             =   4200
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��Ƭ X��"
      Height          =   180
      Index           =   28
      Left            =   8280
      TabIndex        =   57
      Top             =   3840
      Width           =   720
   End
   Begin VB.Line Line14 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      Index           =   5
      Visible         =   0   'False
      X1              =   8280
      X2              =   7560
      Y1              =   6000
      Y2              =   5520
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��Ƭ�� Y2��"
      Height          =   180
      Index           =   27
      Left            =   8280
      TabIndex        =   55
      Top             =   6000
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Line Line14 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      Index           =   4
      Visible         =   0   'False
      X1              =   8260
      X2              =   7560
      Y1              =   5780
      Y2              =   5420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��Ƭ�� X2��"
      Height          =   180
      Index           =   26
      Left            =   8280
      TabIndex        =   53
      Top             =   5640
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Line Line14 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      Index           =   3
      Visible         =   0   'False
      X1              =   8280
      X2              =   7440
      Y1              =   3480
      Y2              =   3240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��Ƭ�� Y1��"
      Height          =   180
      Index           =   25
      Left            =   8280
      TabIndex        =   50
      Top             =   3360
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Line Line14 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      Index           =   2
      Visible         =   0   'False
      X1              =   8280
      X2              =   7560
      Y1              =   3000
      Y2              =   3240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��Ƭ�� X1��"
      Height          =   180
      Index           =   24
      Left            =   8280
      TabIndex        =   48
      Top             =   3000
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Line Line21 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   2760
      X2              =   2880
      Y1              =   6480
      Y2              =   6000
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��쵥λX��"
      Height          =   180
      Index           =   22
      Left            =   2520
      TabIndex        =   45
      Top             =   6480
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      Index           =   3
      Visible         =   0   'False
      X1              =   1800
      X2              =   2520
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��֤����X��"
      Height          =   180
      Index           =   21
      Left            =   240
      TabIndex        =   43
      Top             =   6540
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      Index           =   2
      Visible         =   0   'False
      X1              =   1800
      X2              =   2520
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�����X��"
      Height          =   180
      Index           =   20
      Left            =   240
      TabIndex        =   41
      Top             =   6180
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      Index           =   1
      Visible         =   0   'False
      X1              =   1800
      X2              =   2520
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "���X��"
      Height          =   180
      Index           =   19
      Left            =   240
      TabIndex        =   39
      Top             =   5820
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Line Line14 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      Index           =   1
      Visible         =   0   'False
      X1              =   6360
      X2              =   6600
      Y1              =   6000
      Y2              =   6480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "����X��"
      Height          =   180
      Index           =   18
      Left            =   6480
      TabIndex        =   37
      Top             =   6480
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Line Line14 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      Index           =   0
      X1              =   6720
      X2              =   6480
      Y1              =   2640
      Y2              =   2820
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�Ա�X��"
      Height          =   180
      Index           =   17
      Left            =   6720
      TabIndex        =   35
      Top             =   2520
      Width           =   630
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      Index           =   0
      X1              =   1800
      X2              =   2520
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "����X��"
      Height          =   180
      Index           =   16
      Left            =   240
      TabIndex        =   33
      Top             =   5460
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   1800
      X2              =   2520
      Y1              =   3780
      Y2              =   3780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "����֤��X��"
      Height          =   180
      Index           =   15
      Left            =   240
      TabIndex        =   31
      Top             =   5100
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   5520
      X2              =   6000
      Y1              =   2880
      Y2              =   2280
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      X1              =   1920
      X2              =   2520
      Y1              =   2820
      Y2              =   3420
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   5520
      X2              =   6000
      Y1              =   2880
      Y2              =   1920
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ֽ�����ͣ�"
      Height          =   180
      Index           =   7
      Left            =   240
      TabIndex        =   26
      Top             =   1140
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�����ʼ��"
      Height          =   180
      Index           =   6
      Left            =   120
      TabIndex        =   25
      Top             =   3960
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "������ʼ��"
      Height          =   180
      Index           =   5
      Left            =   120
      TabIndex        =   24
      Top             =   2580
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "���ź��ࣺ"
      Height          =   180
      Index           =   4
      Left            =   4800
      TabIndex        =   23
      Top             =   1140
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�����ݼ�ࣺ"
      Height          =   180
      Index           =   3
      Left            =   3000
      TabIndex        =   22
      Top             =   1140
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�м�ࣺ"
      Height          =   180
      Index           =   8
      Left            =   240
      TabIndex        =   21
      Top             =   4380
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "���壺"
      Height          =   180
      Index           =   9
      Left            =   1560
      TabIndex        =   20
      Top             =   2040
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�����С��"
      Height          =   180
      Index           =   10
      Left            =   3240
      TabIndex        =   19
      Top             =   2040
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�������壺"
      Height          =   180
      Index           =   11
      Left            =   6000
      TabIndex        =   18
      Top             =   1800
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��С��"
      Height          =   180
      Index           =   12
      Left            =   7920
      TabIndex        =   17
      Top             =   1800
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "����֤����X��"
      Height          =   180
      Index           =   13
      Left            =   6000
      TabIndex        =   16
      Top             =   2160
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "����Y��"
      Height          =   180
      Index           =   14
      Left            =   7920
      TabIndex        =   15
      Top             =   2160
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "ע�⣺����������������Ժ���Ϊ��λ��"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   240
      Left            =   5400
      TabIndex        =   2
      Top             =   240
      Width           =   4845
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ѡ���ʽ��"
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   300
      Width           =   900
   End
End
Attribute VB_Name = "frm��ӡ��ʽ����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlng���ư�ʽ As Long


Private Sub ccmb��ʽ_Click()
    Dim lobjPrintSeting As New ClsPrintSeting '������ӡ����,��ȡ��ӡ������Ϣ
    Dim ltxtBase As TextBox
    Dim i As Long
    
    On Error GoTo errhandler
    
    lobjPrintSeting.��ʽ = ccmb��ʽ.ItemData(ccmb��ʽ.ListIndex)
    
    Select Case ccmb��ʽ.ListIndex
        Case 0
            Label1(3).Visible = False
            ctxtBase(0).Visible = False
            Label1(4).Visible = False
            ctxtBase(1).Visible = False
        Case 1
            Label1(3).Visible = True
            ctxtBase(0).Visible = True
            Label1(4).Visible = False
            ctxtBase(1).Visible = False
        Case 2
            Label1(3).Visible = True
            ctxtBase(0).Visible = True
            Label1(4).Visible = True
            ctxtBase(1).Visible = True
    End Select
        
    ccmbֽ��.ListIndex = -1
    For i = 0 To ccmbֽ��.ListCount - 1
        If ccmbֽ��.ItemData(i) = lobjPrintSeting.ֽ������ Then
            ccmbֽ��.ListIndex = i
            Exit For
        End If
    Next
    If ccmbֽ��.ListIndex = -1 Then
        ccmbֽ��.Text = lobjPrintSeting.ֽ������
    End If
    
    For Each ltxtBase In ctxtBase
        ltxtBase.Text = lobjPrintSeting.����ֵ(ltxtBase.Tag)
    Next
    
    If mlng���ư�ʽ <> -1 And mlng���ư�ʽ <> ccmb��ʽ.ItemData(ccmb��ʽ.ListIndex) Then
        ccmdPaste.Enabled = True
    Else
        ccmdPaste.Enabled = False
    End If
    Exit Sub
errhandler:
    sfsub������ "����֤��ӡ��ʽ����A", "frm��ӡ��ʽ����", "ccmb��ʽ_Click", Err.Number, Err.Description, False

End Sub

Private Sub ccmdCopy_Click()
    mlng���ư�ʽ = ccmb��ʽ.ItemData(ccmb��ʽ.ListIndex)
    ccmdPaste.Enabled = False
    
End Sub

Private Sub ccmdExit_Click()
    Unload Me
    
End Sub

Private Sub ccmdPaste_Click()
    On Error GoTo errhandler
    
    If mlng���ư�ʽ <> -1 And mlng���ư�ʽ <> ccmb��ʽ.ItemData(ccmb��ʽ.ListIndex) Then
        '���Ƹ�ʽ��
        sub���Ƹ�ʽ mlng���ư�ʽ, ccmb��ʽ.ItemData(ccmb��ʽ.ListIndex)
        
        ccmb��ʽ_Click
        
        MsgBox "���Ƴɹ���", vbOKOnly + vbInformation, "ϵͳ��ʾ"
    End If
    
    Exit Sub
errhandler:
    sfsub������ "����֤��ӡ��ʽ����A", "frm��ӡ��ʽ����", "ccmdPaste_Click", Err.Number, Err.Description, False
    
End Sub

Private Sub ccmdPreview_Click()
    Dim lobjSet As New ClsPrintSeting
    On Error GoTo errhandler
    
    lobjSet.��ʽ = ccmb��ʽ.ItemData(ccmb��ʽ.ListIndex)
    
    lobjSet.sub���Դ�ӡ
    Exit Sub
errhandler:
    sfsub������ "����֤��ӡ��ʽ����A", "ccmdPreview_Click", "ccmdSave_Click", Err.Number, Err.Description, False
    
End Sub

Private Sub ccmdSave_Click()
    Dim lobjPrintSeting As New ClsPrintSeting '������ӡ����,��ȡ��ӡ������Ϣ
    Dim ltxtBase As TextBox
    Dim i As Long
    
    On Error GoTo errhandler
    
    lobjPrintSeting.��ʽ = ccmb��ʽ.ItemData(ccmb��ʽ.ListIndex)
    lobjPrintSeting.ֽ������ = ccmbֽ��.ItemData(IIf(ccmbֽ��.ListIndex = -1, 0, ccmbֽ��.ListIndex))
    For Each ltxtBase In ctxtBase
        lobjPrintSeting.����ֵ(ltxtBase.Tag) = ltxtBase.Text
    Next
    lobjPrintSeting.sub����
    MsgBox "����ɹ���", vbInformation, "ϵͳ��ʾ"
    Exit Sub
errhandler:
    sfsub������ "����֤��ӡ��ʽ����A", "frm��ӡ��ʽ����", "ccmdSave_Click", Err.Number, Err.Description, False
    
End Sub


Private Sub Form_Load()

    On Error GoTo errhandler
    
    ccmb��ʽ.Clear
    
    ccmb��ʽ.AddItem "����"
    ccmb��ʽ.ItemData(ccmb��ʽ.ListCount - 1) = 2
    ccmb��ʽ.AddItem "1*5(���ϵ���)"
    ccmb��ʽ.ItemData(ccmb��ʽ.ListCount - 1) = 1
    ccmb��ʽ.AddItem "2*5(������)"
    ccmb��ʽ.ItemData(ccmb��ʽ.ListCount - 1) = 0
    
    '���ֽ�����͡�
    ccmbֽ��.Clear
    ccmbֽ��.AddItem "A4"
    ccmbֽ��.ItemData(ccmbֽ��.ListCount - 1) = vbPRPSA4
    ccmbֽ��.AddItem "СA4"
    ccmbֽ��.ItemData(ccmbֽ��.ListCount - 1) = vbPRPSA4Small
    ccmbֽ��.AddItem "A3"
    ccmbֽ��.ItemData(ccmbֽ��.ListCount - 1) = vbPRPSA3
    ccmbֽ��.AddItem "B5"
    ccmbֽ��.ItemData(ccmbֽ��.ListCount - 1) = vbPRPSB5
    ccmbֽ��.AddItem "Legal"
    ccmbֽ��.ItemData(ccmbֽ��.ListCount - 1) = vbPRPSLegal
    ccmbֽ��.AddItem "11x17"
    ccmbֽ��.ItemData(ccmbֽ��.ListCount - 1) = vbPRPS11x17
    ccmbֽ��.AddItem "10x14"
    ccmbֽ��.ItemData(ccmbֽ��.ListCount - 1) = vbPRPS10x14
    ccmbֽ��.AddItem "�Զ���"
    ccmbֽ��.ItemData(ccmbֽ��.ListCount - 1) = vbPRPSUser
    
    ccmb��ʽ.ListIndex = 0
    
    mlng���ư�ʽ = -1
    ccmdPaste.Enabled = False
    
    Exit Sub
errhandler:
    sfsub������ "����֤��ӡ��ʽ����A", "frm��ӡ��ʽ����", "ccmb��ʽ_Click", Err.Number, Err.Description, False
    
End Sub

'���ܣ����Ʋ��������뵥���š�
Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 39 Or KeyAscii = -23636 Then
        '���������롰'������������
        KeyAscii = 0
    End If

End Sub
Sub sub���Ƹ�ʽ(ByVal paraԴ��ʽ As Long, ByVal paraĿ�İ�ʽ As Long)
    Dim lobjԴ��ʽ As New ClsPrintSeting
    Dim lobjĿ�ĸ�ʽ As ClsPrintSeting
    
    On Error GoTo errhandler
    If paraԴ��ʽ <> paraĿ�İ�ʽ Then
        lobjԴ��ʽ.��ʽ = paraԴ��ʽ
        Set lobjĿ�ĸ�ʽ = lobjԴ��ʽ.Clone(paraĿ�İ�ʽ)
        lobjĿ�ĸ�ʽ.sub����
    End If
    
    Exit Sub
errhandler:
    sfsub������ "����֤��ӡ��ʽ����A", "cls��ʽ����", "sub���Ƹ�ʽ", Err.Number, Err.Description, True
    Exit Sub
    
End Sub


