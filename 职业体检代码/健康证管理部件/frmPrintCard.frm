VERSION 5.00
Begin VB.Form frmPrintCard 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4785
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7770
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "����"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   7770
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame FrameM 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2625
      Index           =   0
      Left            =   120
      TabIndex        =   17
      Top             =   1080
      Width           =   5010
      Begin VB.Label clbl������� 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   38
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Line Line5 
         X1              =   2880
         X2              =   3240
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label cLab���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���䣺"
         BeginProperty Font 
            Name            =   "����"
            Size            =   7.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   2400
         TabIndex        =   37
         Top             =   1080
         Width           =   450
      End
      Begin VB.Line Line4 
         X1              =   1200
         X2              =   2160
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line3 
         X1              =   1440
         X2              =   3240
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Label cLbl��쵥λ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������쵥λ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   7.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   360
         TabIndex        =   36
         Top             =   2160
         Width           =   1050
      End
      Begin VB.Line Line1 
         X1              =   1200
         X2              =   3240
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label cinfo��ע 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "����"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   150
         Left            =   3720
         TabIndex        =   35
         Top             =   2040
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label cInfo���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "����"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   150
         Left            =   3960
         TabIndex        =   34
         Top             =   2160
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label cInfo��λ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "����"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   150
         Left            =   2280
         TabIndex        =   33
         Top             =   2040
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label cInfo�Թܱ�� 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   300
         TabIndex        =   32
         Top             =   1920
         Width           =   90
      End
      Begin VB.Label cInfo��֤���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "����"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   150
         Left            =   1200
         TabIndex        =   31
         Top             =   1680
         Width           =   90
      End
      Begin VB.Image Cphoto 
         Height          =   1425
         Left            =   3660
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1080
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00E38B5B&
         BorderWidth     =   3
         Height          =   1455
         Index           =   0
         Left            =   3660
         Top             =   360
         Width           =   1110
      End
      Begin VB.Label cLab��֤���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��֤���ڣ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   7.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   360
         TabIndex        =   30
         Top             =   1800
         Width           =   750
      End
      Begin VB.Line cLne��֤���� 
         Index           =   0
         X1              =   1200
         X2              =   3240
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label cLab����֤�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��   �ţ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   7.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Index           =   0
         Left            =   360
         TabIndex        =   29
         Top             =   360
         Width           =   675
      End
      Begin VB.Label cLab���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��   ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   7.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Index           =   0
         Left            =   360
         TabIndex        =   28
         Top             =   1080
         Width           =   675
      End
      Begin VB.Label cInfo���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "����"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   150
         Left            =   1200
         TabIndex        =   27
         Top             =   960
         Width           =   90
      End
      Begin VB.Label cInfo���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "����"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   150
         Left            =   2880
         TabIndex        =   26
         Top             =   960
         Width           =   90
      End
      Begin VB.Line cLne�Ա� 
         Index           =   0
         X1              =   2880
         X2              =   3240
         Y1              =   800
         Y2              =   800
      End
      Begin VB.Label cInfo�Ա� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "����"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   150
         Left            =   2880
         TabIndex        =   25
         Top             =   600
         Width           =   90
      End
      Begin VB.Label cLab�Ա� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   7.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Index           =   0
         Left            =   2400
         TabIndex        =   24
         Top             =   720
         Width           =   450
      End
      Begin VB.Line cLne���� 
         Index           =   0
         X1              =   1200
         X2              =   2160
         Y1              =   800
         Y2              =   800
      End
      Begin VB.Label cInfo���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "����"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   150
         Left            =   1200
         TabIndex        =   23
         Top             =   600
         Width           =   90
      End
      Begin VB.Label cLab���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��    ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   7.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Index           =   0
         Left            =   360
         TabIndex        =   22
         Top             =   660
         Width           =   750
      End
      Begin VB.Label cinfo֤�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1200
         TabIndex        =   21
         Top             =   240
         Width           =   60
      End
      Begin VB.Label cLbl��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��    �죺"
         BeginProperty Font 
            Name            =   "����"
            Size            =   7.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Index           =   0
         Left            =   360
         TabIndex        =   20
         Top             =   1395
         Width           =   750
      End
      Begin VB.Label cLbl�ϸ� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ϸ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   7.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Index           =   0
         Left            =   3720
         TabIndex        =   19
         Top             =   2160
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label Cinfo��쵥λ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "����"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   1440
         TabIndex        =   18
         Top             =   2040
         Width           =   90
      End
      Begin VB.Line cLne��� 
         Index           =   0
         X1              =   1200
         X2              =   3240
         Y1              =   1560
         Y2              =   1560
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "ѡ��ģ��"
      Height          =   2295
      Left            =   5280
      TabIndex        =   7
      Top             =   1080
      Width           =   2415
      Begin VB.Frame cframPos 
         Caption         =   "��ʼ��"
         Height          =   735
         Left            =   60
         TabIndex        =   11
         Top             =   1560
         Visible         =   0   'False
         Width           =   2295
         Begin VB.ComboBox ccmbIndex 
            Height          =   300
            ItemData        =   "frmPrintCard.frx":0000
            Left            =   1320
            List            =   "frmPrintCard.frx":0013
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   360
            Width           =   600
         End
         Begin VB.ComboBox ccmbSide 
            Height          =   300
            ItemData        =   "frmPrintCard.frx":0026
            Left            =   120
            List            =   "frmPrintCard.frx":0030
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   360
            Width           =   600
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��"
            Height          =   180
            Index           =   1
            Left            =   2000
            TabIndex        =   15
            Top             =   480
            Width           =   180
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ߵ�"
            Height          =   180
            Index           =   0
            Left            =   800
            TabIndex        =   14
            Top             =   480
            Width           =   360
         End
      End
      Begin VB.OptionButton optģ������ 
         Caption         =   "���Ŵ�ӡ"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Width           =   1095
      End
      Begin VB.OptionButton optģ������ 
         Caption         =   "1 * 5(���ϵ���)"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   1695
      End
      Begin VB.OptionButton optģ������ 
         Caption         =   "2 * 5(������)"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Value           =   -1  'True
         Width           =   1695
      End
   End
   Begin VB.CommandButton cComNext 
      Caption         =   "��һ��>&>"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cComPrev 
      Caption         =   "&<<��һ��"
      Enabled         =   0   'False
      Height          =   375
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cComCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(Esc)"
      Height          =   375
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cComPrintCard 
      Caption         =   "��ӡ(&P)"
      Default         =   -1  'True
      Height          =   375
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label cInfo��Ч��ֹ���� 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   840
      TabIndex        =   16
      Top             =   3480
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   6600
      Picture         =   "frmPrintCard.frx":003C
      Top             =   120
      Width           =   480
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   7800
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "ϵͳ�ڴ˴�Ԥ������֤��Ч�����ṩ���Ĵ�ӡȷ�ϡ������ͨ������""��һ��""��""��һ��""Ԥ�������㽫Ҫ��ӡ�Ľ���֤��"
      Height          =   495
      Left            =   390
      TabIndex        =   5
      Top             =   120
      Width           =   5295
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "����Ϊ����֤�Ĵ�ӡģ�棬����""��ӡ""ȷ�ϣ�����""ȡ��""�˳���"
      Height          =   180
      Left            =   480
      TabIndex        =   2
      Top             =   3840
      Width           =   5040
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   120
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   7755
   End
End
Attribute VB_Name = "frmPrintCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Cards As New Collection
Private m��֤��λ As String             '���������¼��֤��λ

Dim intIndex As Long

Public pblnPrint As Boolean

Private Sub cComCancel_Click()

On Error GoTo errordeal

    Me.Hide
    Set Cards = New Collection
    pblnPrint = False
Exit Sub

errordeal:
    MsgBox Err.Description, vbInformation, "����֤ϵͳ"
    On Error Resume Next
End Sub

Private Sub cComNext_Click()

On Error GoTo errordeal

    intIndex = intIndex + 1
    FillForm intIndex
    
Exit Sub

errordeal:
    MsgBox Err.Description, vbInformation, "����֤ϵͳ"
    On Error Resume Next

End Sub

Private Sub cComPrev_Click()

On Error GoTo errordeal

    intIndex = intIndex - 1
    FillForm intIndex

Exit Sub

errordeal:
    MsgBox Err.Description, vbInformation, "����֤ϵͳ"
    On Error Resume Next

End Sub

Private Sub FillForm(ByVal Index As Long)
On Error GoTo errordeal
    
    Dim lobjRectemp As Object       '���������¼��ʱ��¼��
    
    If Index >= 1 And Index <= Cards.Count Then
        cInfo����.Caption = Cards(Index).����
        cInfo�Ա�.Caption = Cards(Index).�Ա�
        cInfo����.Caption = Cards(Index).����
        
        cInfo����.Caption = Cards(Index).����
        Cinfo��쵥λ.Caption = Cards(Index).��֤��λ
        cInfo��֤����.Caption = Cards(Index).��֤����
        
        clbl�������.Caption = IIf(Cards(Index).������� = "��", "�޴�ҵ����֢", Cards(Index).�������)
        cinfo֤��.Caption = Cards(Index).����֤��
        cInfo��λ.Caption = Cards(Index).��λ����
        cInfo����.Caption = Cards(Index).����
        cinfo��ע.Caption = IIf(Cards(Index).������� = "����", "", Cards(Index).�������)
        Cphoto.Picture = Cards(Index).��Ƭ
    End If
    
    If Index = 1 Then
        cComPrev.Enabled = False
        cComNext.Enabled = True
    ElseIf Index = Cards.Count Then
        cComNext.Enabled = False
        cComPrev.Enabled = True
    ElseIf Index > 1 And Index < Cards.Count Then
        cComPrev.Enabled = True
        cComNext.Enabled = True
    ElseIf Cards.Count = 1 Then
        cComPrev.Enabled = False
        cComNext.Enabled = False
    End If

Exit Sub

errordeal:
    MsgBox Err.Description, vbInformation, "����֤ϵͳ"
    On Error Resume Next
    Resume
End Sub

Private Sub cComPrintCard_Click()

On Error GoTo errordeal


    Me.Hide
    
    pblnPrint = True
Exit Sub

errordeal:
    MsgBox Err.Description, vbInformation, "����֤ϵͳ"
    On Error Resume Next

End Sub



Private Sub Form_Activate()
'    intIndex = 1
'    Select Case Cards.Count
'    Case 1
'        cComNext.Enabled = False
'        cComPrev.Enabled = False
'    Case Is > 1
'        cComNext.Enabled = True
'        cComPrev.Enabled = False
'    End Select
'
'    If Cards.Count <= 5 Then
'        optģ������(1).Value = True
'    ElseIf Cards.Count > 5 Then
'        optģ������(0).Value = True
'    End If
'
'    If Cards.Count >= 1 Then
'        FillForm intIndex
'    End If
'
'    '�޸ģ�2002-5-23�����Ӵ�ӡλ�����ã���
'    If Cards.Count < 10 Then
'        cframPos.Visible = True
'        cframPos.Enabled = True
'    Else
'        cframPos.Visible = False
'        cframPos.Enabled = False
'    End If
End Sub

Private Sub Form_Load()

On Error GoTo errordeal
    
    intIndex = 1
    Select Case Cards.Count
    Case 1
        cComNext.Enabled = False
        cComPrev.Enabled = False
    Case Is > 1
        cComNext.Enabled = True
        cComPrev.Enabled = False
    End Select
    If Cards.Count = 1 Then
        optģ������(2).Value = True
    ElseIf Cards.Count <= 5 Then
        optģ������(1).Value = True
    ElseIf Cards.Count > 5 Then
        optģ������(0).Value = True
    End If
    
    If Cards.Count >= 1 Then
        FillForm intIndex
    End If
    
    '�޸ģ�2002-5-23�����Ӵ�ӡλ�����ã���
    ccmbSide.ListIndex = 0
    ccmbIndex.ListIndex = 0
    
    If Cards.Count < 10 Then
        cframPos.Visible = True
        cframPos.Enabled = True
    Else
        cframPos.Visible = False
        cframPos.Enabled = False
    End If
    
    
Exit Sub

errordeal:
    MsgBox Err.Description, vbInformation, "����֤ϵͳ"
    On Error Resume Next

End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error GoTo errordeal

    Set Cards = Nothing

Exit Sub

errordeal:
    MsgBox Err.Description, vbInformation, "����֤ϵͳ"
    On Error Resume Next

End Sub

'�޸ģ�2002-5-23�����Ӵ�ӡλ�����ã���
Private Sub optģ������_Click(Index As Integer)
    Dim i As Long
    Dim lstrCN As String
    If Index = 0 Then
        If Cards.Count < 10 Then
            cframPos.Visible = True
            cframPos.Enabled = True
        Else
            cframPos.Visible = False
            cframPos.Enabled = False
        End If
        
    Else
        cframPos.Visible = False
        cframPos.Enabled = False
        
    End If
    
    lstrCN = Cards(1).����֤��
    If optģ������(0).Value Then
        For i = 2 To Cards.Count
            lstrCN = Format(Val(lstrCN) + 1, String(Len(lstrCN), "0"))
            Cards(i).����֤�� = lstrCN
        Next
        FillForm intIndex
    ElseIf optģ������(1).Value Then
        For i = 2 To Cards.Count
            lstrCN = Format(Val(lstrCN) + 2, String(Len(lstrCN), "0"))
            Cards(i).����֤�� = lstrCN
        Next
        FillForm intIndex
    End If
End Sub
