VERSION 5.00
Object = "{D9347025-9612-11D1-9D75-00C04FCC8CDC}#1.0#0"; "MSBCODE9.OCX"
Begin VB.Form frmPrintPVCCard 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3585
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6945
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   6945
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton ccmdNext 
      Caption         =   "��һ��(&>)"
      Height          =   495
      Left            =   5520
      TabIndex        =   24
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton ccmdPrve 
      Caption         =   "��һ��(&<)"
      Height          =   495
      Left            =   5520
      TabIndex        =   23
      Top             =   840
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   240
      Picture         =   "frmPrintPVCCard.frx":0000
      ScaleHeight     =   135
      ScaleWidth      =   5055
      TabIndex        =   22
      Top             =   3120
      Width           =   5055
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   240
      Picture         =   "frmPrintPVCCard.frx":0433
      ScaleHeight     =   495
      ScaleWidth      =   5055
      TabIndex        =   21
      Top             =   240
      Width           =   5055
   End
   Begin VB.CommandButton ccmdCancel 
      Caption         =   "�˳�(&E)"
      Height          =   495
      Left            =   5520
      TabIndex        =   2
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton ccmdPrint 
      Caption         =   "��ӡ(&P)"
      Height          =   495
      Left            =   5520
      TabIndex        =   1
      Top             =   2280
      Width           =   1215
   End
   Begin VB.PictureBox PICCard 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   240
      ScaleHeight     =   3015
      ScaleWidth      =   5055
      TabIndex        =   0
      Top             =   240
      Width           =   5055
      Begin VB.Label clblѪ�� 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3720
         TabIndex        =   25
         Top             =   2280
         Width           =   615
      End
      Begin VB.Image cPhoto 
         Height          =   1605
         Left            =   3360
         Stretch         =   -1  'True
         Top             =   600
         Width           =   1200
      End
      Begin VB.Label cinfo��֤����2 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   960
         TabIndex        =   20
         Top             =   2230
         Width           =   2655
      End
      Begin VB.Label cinfo��֤����1 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   960
         TabIndex        =   19
         Top             =   2040
         Width           =   4335
      End
      Begin VB.Label cinfo��� 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   600
         TabIndex        =   18
         Top             =   1770
         Width           =   3375
      End
      Begin VB.Label cinfo��Ч��ֹ 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   960
         TabIndex        =   17
         Top             =   1515
         Width           =   2775
      End
      Begin VB.Label cinfo��Ч���� 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   960
         TabIndex        =   16
         Top             =   1245
         Width           =   2775
      End
      Begin VB.Label cinfo���� 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   2760
         TabIndex        =   15
         Top             =   990
         Width           =   735
      End
      Begin VB.Label cinfo���� 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   960
         TabIndex        =   14
         Top             =   990
         Width           =   975
      End
      Begin VB.Label cinfo�Ա� 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   2760
         TabIndex        =   13
         Top             =   720
         Width           =   735
      End
      Begin VB.Label cinfo���� 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   600
         TabIndex        =   12
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   720
         TabIndex        =   10
         Top             =   1515
         Width           =   195
      End
      Begin VB.Label cLab����֤�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ţ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   9
         Top             =   1770
         Width           =   585
      End
      Begin VB.Label cLab��֤���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��֤������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   8
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label cLab���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ҵ���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   7
         Top             =   990
         Width           =   975
      End
      Begin VB.Label cLab���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���䣺"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   2280
         TabIndex        =   6
         Top             =   990
         Width           =   585
      End
      Begin VB.Label cLab�Ա� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   2280
         TabIndex        =   5
         Top             =   720
         Width           =   585
      End
      Begin VB.Label cLab���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   585
      End
      Begin VB.Label cLbl��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ч���ޣ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   3
         Top             =   1245
         Width           =   975
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000002&
         Height          =   1605
         Left            =   3360
         Top             =   600
         Width           =   1200
      End
      Begin BARCODELibCtl.BarCodeCtrl cBar����֤��� 
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   2400
         Width           =   4335
         Style           =   7
         SubStyle        =   -1
         Validation      =   0
         LineWeight      =   1
         Direction       =   0
         ShowData        =   0
         Value           =   "123456 Code-128"
         ForeColor       =   0
         BackColor       =   16777215
      End
   End
End
Attribute VB_Name = "frmPrintPVCCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Cards As New Collection
Private m��֤��λ As String             '���������¼��֤��λ

Dim intIndex As Long

Private Sub Command2_Click()

End Sub

Private Sub ccmdCancel_Click()
On Error GoTo errordeal

'    IFPrint = False
    Me.Hide
    dafuncGetData ("update ϵͳ����_ϵͳ������ɼ�¼�� set ��ǰֵ=��ǰֵ-" & Cards.Count & " where ҵ������='����֤����' and �������='����֤���'")
    Set Cards = New Collection
    Unload Me
Exit Sub

errordeal:
    MsgBox Err.Description, vbInformation, "����֤ϵͳ"
    On Error Resume Next
End Sub


Private Sub FillForm(ByVal Index As Long)
On Error GoTo errordeal
   Dim lstr��Ч���� As String
   Dim lstrǰ׺ As String
   Dim lobjtemp As Object
    m��֤��λ = um����վ��
    Set lobjtemp = dafuncGetData("select top 1 �غ���ַ from ����֤_ҵ�����ñ�")
    If lobjtemp.RecordCount > 0 Then
    lstrǰ׺ = lobjtemp(0)
    Else
        lstrǰ׺ = ""
    End If
    If Index >= 1 And Index <= Cards.Count Then
        cinfo����.Caption = Cards(Index).����
        cinfo�Ա�.Caption = Cards(Index).�Ա�
        cinfo����.Caption = Cards(Index).����
        cinfo����.Caption = Cards(Index).����
        
        cinfo��Ч���� = Left(Cards(Index).��֤����, 4) + "��" + Mid(Cards(Index).��֤����, 6, 2) + "��" + Right(Cards(Index).��֤����, 2) + "��"
        lstr��Ч���� = DateAdd("d", -1, DateAdd("yyyy", 1, Cards(Index).��֤����))
        cinfo��Ч��ֹ = Left(lstr��Ч����, 4) + "��" + Mid(lstr��Ч����, 6, 2) + "��" + Right(lstr��Ч����, 2) + "��"
        cinfo���.Caption = "��" + "(" + Left(Cards(Index).��֤����, 4) + ")" + lstrǰ׺ + "-" + Cards(Index).����֤��
'        cinfo���.Caption = "��" + lstrǰ׺ + "(" + Left(Cards(Index).Date, 4) + ")��" + "00000000" + "��"
        cinfo��֤����1.Caption = Left(m��֤��λ, 12)
        If Len(m��֤��λ) > 12 Then
            cinfo��֤����2.Caption = Right(m��֤��λ, Len(m��֤��λ) - 12)
        End If
        
'        cBar����֤���.Value = Right(Cards(Index).SN, 8)
        cBar����֤���.Value = IIf(Cards(Index).�������֤��� = "", Cards(Index).���ϵͳ���, Cards(Index).�������֤���)
'        cBar����֤���.Value = "AFEdU/ZW13N0UDAA"

        cPhoto.Picture = Cards(Index).��Ƭ
'        clabSysNo.Caption = Cards(Index).���ϵͳ���
        clblѪ��.Caption = Cards(Index).����
    End If
    

Exit Sub

errordeal:
    MsgBox Err.Description, vbInformation, "����֤ϵͳ"
    On Error Resume Next

End Sub

Private Sub ccmdNext_Click()
intIndex = intIndex + 1
If intIndex = Cards.Count Then
    ccmdNext.Enabled = False
End If
If Cards.Count > 1 Then
    ccmdPrve.Enabled = True
   
End If

 FillForm intIndex
End Sub

Private Sub ccmdPrint_Click()
On Error GoTo errordeal
Dim i As Integer
'    IFPrint = True
Me.Hide
Me.BackColor = vbWhite
PICCard.Top = 140
PICCard.Left = 150
Picture1.Visible = False
Picture2.Visible = False

For i = 1 To Cards.Count
    FillForm i

    Me.PrintForm
    
    dafuncGetData "Update ����֤����_��֤������Ϣ��  Set ״̬='�Ѵ�ӡ',����֤�� ='" & Cards(i).����֤�� & "',���֤�� ='" & Cards(i).���֤�� & "',��֤����='" & Cards(i).��֤���� & "',��Ч����='" & Cards(i).��Ч���� & "', ��֤��λ='" & Cards(i).��֤��λ & "' where ϵͳ��� ='" & Cards(i).ϵͳ��� & "'"
Next
    
Set Cards = New Collection
Exit Sub

errordeal:
    MsgBox Err.Description, vbInformation, "����֤ϵͳ"
    On Error Resume Next
End Sub

Private Sub ccmdPrve_Click()
intIndex = intIndex - 1
If intIndex = 1 Then
    ccmdPrve.Enabled = False
End If
If intIndex < Cards.Count Then
    ccmdNext.Enabled = True
   
End If

 FillForm intIndex
End Sub

Private Sub Form_Load()
On Error GoTo errordeal
    
    intIndex = 1

    
    If Cards.Count >= 1 Then
        FillForm intIndex
    End If
    ccmdPrve.Enabled = False
    If Cards.Count = 1 Then
        ccmdNext.Enabled = False
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
