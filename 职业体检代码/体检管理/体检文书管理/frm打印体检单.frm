VERSION 5.00
Object = "{D9347025-9612-11D1-9D75-00C04FCC8CDC}#1.0#0"; "MSBCODE9.OCX"
Begin VB.Form frm��ӡ��쵥 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000009&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   14625
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   14625
   ScaleWidth      =   9870
   ShowInTaskbar   =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame cfram��쵥 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6120
      Left            =   300
      TabIndex        =   28
      Top             =   570
      Width           =   9345
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�� �� �ţ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   90
         TabIndex        =   48
         Top             =   915
         Width           =   1050
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��    ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   90
         TabIndex        =   47
         Top             =   1267
         Width           =   1050
      End
      Begin VB.Label label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���֤�ţ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   90
         TabIndex        =   46
         Top             =   1971
         Width           =   1050
      End
      Begin VB.Label clblSysNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   1425
         TabIndex        =   45
         Top             =   915
         Width           =   105
      End
      Begin VB.Label clblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   1425
         TabIndex        =   44
         Top             =   1267
         Width           =   105
      End
      Begin VB.Label clblIDCard 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1425
         TabIndex        =   43
         Top             =   1971
         Width           =   105
      End
      Begin VB.Image cimgPhoto 
         Height          =   1845
         Left            =   7080
         Stretch         =   -1  'True
         Top             =   1005
         Width           =   1410
      End
      Begin VB.Label clblUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ݹ�ҵ԰���������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   6120
         TabIndex        =   42
         Top             =   5640
         Width           =   2100
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "лл������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   3360
         TabIndex        =   41
         Top             =   5250
         Width           =   1050
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "5.���д����������Ϊ��ȡ�����������ͨ����λ��������Դ��˾��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   90
         TabIndex        =   40
         Top             =   4440
         Width           =   6510
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "4.�����뵱����ɣ�������Ч��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   90
         TabIndex        =   39
         Top             =   4110
         Width           =   3150
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3.���Ѫѹƫ���ߣ�����Ϣһ���ӣ���⼸�Σ�ѡ���е�ֵ���롣����������쵱��ѪѹΪ׼"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   90
         TabIndex        =   38
         Top             =   3810
         Width           =   8610
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2.�����ֻ���װ����Һ���򱭷ŵ�ָ���ĵط������鵥ѹ�����漴��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   90
         TabIndex        =   37
         Top             =   3525
         Width           =   6510
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         Index           =   1
         X1              =   105
         X2              =   9000
         Y1              =   3435
         Y2              =   3420
      End
      Begin VB.Label clbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������   �����   ���   �ĵ�ͼ   �ۿ�   �ڿ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   90
         TabIndex        =   36
         Top             =   3150
         Width           =   8685
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   1
         X1              =   105
         X2              =   9000
         Y1              =   3045
         Y2              =   3045
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1.������������ƾ����쵥�����¸����ҽ������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   90
         TabIndex        =   35
         Top             =   2775
         Width           =   4620
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��쵥"
         BeginProperty Font 
            Name            =   "����"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3270
         TabIndex        =   34
         Top             =   195
         Width           =   1365
      End
      Begin BARCODELibCtl.BarCodeCtrl cbccMain 
         Height          =   765
         Index           =   0
         Left            =   6600
         TabIndex        =   33
         Top             =   120
         Width           =   3015
         Style           =   7
         SubStyle        =   -1
         Validation      =   0
         LineWeight      =   3
         Direction       =   0
         ShowData        =   1
         Value           =   "123456 Code-128"
         ForeColor       =   0
         BackColor       =   16777215
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ���ƣ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   90
         TabIndex        =   32
         Top             =   1619
         Width           =   1050
      End
      Begin VB.Label clbl��λ���� 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1425
         TabIndex        =   31
         Top             =   1619
         Width           =   105
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�շѱ�ţ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   90
         TabIndex        =   30
         Top             =   2325
         Width           =   1050
      End
      Begin VB.Label clbl�շ����� 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1425
         TabIndex        =   29
         Top             =   2325
         Width           =   105
      End
   End
   Begin VB.Frame cframѪҺ����鵥 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   1725
      Left            =   300
      TabIndex        =   18
      Top             =   11880
      Width           =   9465
      Begin VB.Line Line3 
         BorderWidth     =   2
         Index           =   5
         X1              =   75
         X2              =   9360
         Y1              =   30
         Y2              =   30
      End
      Begin BARCODELibCtl.BarCodeCtrl cbccMain 
         Height          =   765
         Index           =   3
         Left            =   6720
         TabIndex        =   26
         Top             =   120
         Width           =   3015
         Style           =   7
         SubStyle        =   -1
         Validation      =   0
         LineWeight      =   3
         Direction       =   0
         ShowData        =   1
         Value           =   "123456 Code-128"
         ForeColor       =   0
         BackColor       =   16777215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�� �� �ţ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   5
         Left            =   90
         TabIndex        =   25
         Top             =   870
         Width           =   1050
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��    ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   5
         Left            =   90
         TabIndex        =   24
         Top             =   1125
         Width           =   1050
      End
      Begin VB.Label label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��    ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   5
         Left            =   90
         TabIndex        =   23
         Top             =   1395
         Width           =   1050
      End
      Begin VB.Label clblSysNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   1200
         TabIndex        =   22
         Top             =   870
         Width           =   105
      End
      Begin VB.Label clblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   1200
         TabIndex        =   21
         Top             =   1125
         Width           =   105
      End
      Begin VB.Label clblSex 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   1200
         TabIndex        =   20
         Top             =   1395
         Width           =   105
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ѪҺ���鵥"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2715
         TabIndex        =   19
         Top             =   360
         Width           =   1650
      End
   End
   Begin VB.Frame cfram�����������鵥 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   1740
      Left            =   240
      TabIndex        =   9
      Top             =   9480
      Width           =   9465
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����������鵥"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2715
         TabIndex        =   17
         Top             =   360
         Width           =   2310
      End
      Begin VB.Label clblSex 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   1200
         TabIndex        =   16
         Top             =   1395
         Width           =   105
      End
      Begin VB.Label clblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   1200
         TabIndex        =   15
         Top             =   1125
         Width           =   105
      End
      Begin VB.Label clblSysNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   1200
         TabIndex        =   14
         Top             =   810
         Width           =   105
      End
      Begin VB.Label label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��    ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   90
         TabIndex        =   13
         Top             =   1395
         Width           =   1050
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��    ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   90
         TabIndex        =   12
         Top             =   1125
         Width           =   1050
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�� �� �ţ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   90
         TabIndex        =   11
         Top             =   870
         Width           =   1050
      End
      Begin BARCODELibCtl.BarCodeCtrl cbccMain 
         Height          =   765
         Index           =   2
         Left            =   6720
         TabIndex        =   10
         Top             =   120
         Width           =   3015
         Style           =   7
         SubStyle        =   -1
         Validation      =   0
         LineWeight      =   3
         Direction       =   0
         ShowData        =   1
         Value           =   "123456 Code-128"
         ForeColor       =   0
         BackColor       =   16777215
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         Index           =   4
         X1              =   90
         X2              =   9240
         Y1              =   30
         Y2              =   30
      End
   End
   Begin VB.Frame cfram�򳣹���鵥 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   1965
      Left            =   300
      TabIndex        =   0
      Top             =   6960
      Width           =   9345
      Begin VB.Label clblSex 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   1200
         TabIndex        =   27
         Top             =   1395
         Width           =   105
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         Index           =   1
         X1              =   90
         X2              =   9240
         Y1              =   30
         Y2              =   30
      End
      Begin BARCODELibCtl.BarCodeCtrl cbccMain 
         Height          =   765
         Index           =   1
         Left            =   6600
         TabIndex        =   8
         Top             =   120
         Width           =   3015
         Style           =   7
         SubStyle        =   -1
         Validation      =   0
         LineWeight      =   3
         Direction       =   0
         ShowData        =   1
         Value           =   "123456 Code-128"
         ForeColor       =   0
         BackColor       =   16777215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�� �� �ţ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   90
         TabIndex        =   7
         Top             =   870
         Width           =   1050
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��    ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   90
         TabIndex        =   6
         Top             =   1125
         Width           =   1050
      End
      Begin VB.Label label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��    ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   90
         TabIndex        =   5
         Top             =   1395
         Width           =   1050
      End
      Begin VB.Label clblSysNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   1200
         TabIndex        =   4
         Top             =   810
         Width           =   105
      End
      Begin VB.Label clblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   1200
         TabIndex        =   3
         Top             =   1125
         Width           =   105
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�򳣹���鵥"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2715
         TabIndex        =   2
         Top             =   360
         Width           =   1995
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ע�������鵥�����򳣹���飬ѹ��װ����Һ�����£�����ָ���ط�����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   90
         TabIndex        =   1
         Top             =   1665
         Width           =   6930
      End
   End
End
Attribute VB_Name = "frm��ӡ��쵥"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'���ߣ��˺㡣

Public pobj�������� As Object 'recordset[ϵͳ��ţ����������֤�ţ��Ա𣬵�λ����]

Public pbln�Ƿ����򳣹� As Boolean
Public pbln�Ƿ��д������ As Boolean
Public pbln�Ƿ�����Ѫ As Boolean

'���ܣ����������ݡ�
'���ߣ��˺㡣
Private Sub Form_Load()
    On Error GoTo errHandler
    Dim i As Integer
    Label1.Left = (Me.Width - Label1.Width) / 2
    Label5.Left = (Me.Width - Label5.Width) / 2
    Label7.Left = (Me.Width - Label7.Width) / 2
    Label17.Left = (Me.Width - Label17.Width) / 2
    
    '���������ݡ�
    For i = 0 To clblSysNo.Count - 1
        clblSysNo(i).Caption = pobj��������("ϵͳ���")
    Next i
    For i = 0 To clblName.Count - 1
        clblName(i).Caption = pobj��������("����")
    Next i
    clblIDCard.Caption = pobj��������("������ݺ���")
    For i = 0 To clblSex.Count - 1
        clblSex(i).Caption = pobj��������("�Ա�")
    Next i

    For i = 0 To cbccMain.Count - 1
        cbccMain(i).Value = pobj��������("ϵͳ���")
    Next i
    clbl�շ�����.Caption = pobj��������("�շ�����")
    clbl��λ����.Caption = pobj��������("��λ����")
    
    '����վ����
    clblUnit(1).Caption = um����վ��
    
    '�����Ҫ���ü�顣
    If Not pbln�Ƿ����򳣹� Then
        cfram�򳣹���鵥.Visible = False
    End If
    If Not pbln�Ƿ��д������ Then
        cfram�����������鵥.Visible = False
    End If
    If Not pbln�Ƿ�����Ѫ Then
        cframѪҺ����鵥.Visible = False
    End If
    
    '���������󣬻�ȡ��Ƭ��
    Dim lobj��� As Object
    Set lobj��� = CreateObject("������.clsMedicalExam")
    lobj���.ϵͳ��� = pobj��������("ϵͳ���")
    
    '��ʾ��Ƭ��
    cimgPhoto.Picture = lobj���.�����Ա.��Ƭ
    
    '�޸ģ�2002-8-16�������̬��ʾ���ҡ�
    On Error Resume Next
    clbl����.Caption = pobj��������("����������")
    
    Exit Sub
errHandler:
    sfsub������ "����������", "frm��ӡ��쵥3", "Form_Load", Err.Number, Err.Description, False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    
End Sub
