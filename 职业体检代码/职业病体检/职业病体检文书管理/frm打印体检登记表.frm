VERSION 5.00
Object = "{D9347025-9612-11D1-9D75-00C04FCC8CDC}#1.0#0"; "MSBCODE9.OCX"
Begin VB.Form frm��ӡ���ǼǱ� 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9945
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   9945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin BARCODELibCtl.BarCodeCtrl cbccMain 
      Height          =   765
      Left            =   4320
      TabIndex        =   10
      Top             =   240
      Width           =   3015
      Style           =   7
      SubStyle        =   0
      Validation      =   0
      LineWeight      =   3
      Direction       =   0
      ShowData        =   1
      Value           =   "123456 Code-128"
      ForeColor       =   0
      BackColor       =   16777215
   End
   Begin VB.Label clbl���� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8400
      TabIndex        =   9
      Top             =   2280
      Width           =   120
   End
   Begin VB.Label clbl��ҵ��� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7200
      TabIndex        =   8
      Top             =   2880
      Width           =   120
   End
   Begin VB.Label clblUnit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2250
      TabIndex        =   7
      Top             =   2880
      Width           =   120
   End
   Begin VB.Label clblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6480
      TabIndex        =   6
      Top             =   2280
      Width           =   120
   End
   Begin VB.Label clblIDCard 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2265
      TabIndex        =   5
      Top             =   2340
      Width           =   120
   End
   Begin VB.Label clblPhotoNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7320
      TabIndex        =   4
      Top             =   1755
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label clblSysNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4560
      TabIndex        =   3
      Top             =   1755
      Width           =   120
   End
   Begin VB.Label clblDay 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3030
      TabIndex        =   2
      Top             =   1755
      Width           =   120
   End
   Begin VB.Label clblMonth 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2190
      TabIndex        =   1
      Top             =   1755
      Width           =   120
   End
   Begin VB.Label clblYear 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1305
      TabIndex        =   0
      Top             =   1755
      Width           =   120
   End
   Begin VB.Image cimgPhoto 
      Height          =   1500
      Left            =   8160
      Top             =   240
      Width           =   1185
   End
End
Attribute VB_Name = "frm��ӡ���ǼǱ�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public pobj�������� As Object 'recordset[ϵͳ��ţ����������֤�ţ��Ա𣬵�λ����]

Private Sub Form_Load()
    Dim lstrTmp As String
    
    On Error GoTo errhandler
    
    '���������ݡ�
    clblYear.Caption = Left(Format(pobj��������("�������"), "yyyy-mm-dd"), 4)
    clblMonth.Caption = Format(Month(pobj��������("�������")), "00")
    clblDay.Caption = Format(Day(pobj��������("�������")), "00")
    
    clblSysNo.Caption = pobj��������("ϵͳ���")
    clblPhotoNo.Caption = pobj��������("�����������")
    
    clblIDCard.Caption = IIf(IsNull(pobj��������("������ݺ���")), "", pobj��������("������ݺ���"))
    
    clblName.Caption = IIf(IsNull(pobj��������("����")), "", pobj��������("����"))
    
    clblUnit.Caption = IIf(IsNull(pobj��������("��λ����")), "", pobj��������("��λ����"))
    
    
    clbl���� = IIf(IsNull(pobj��������("����")), "", pobj��������("����"))
    
    lstrTmp = IIf(IsNull(pobj��������("��������")), "", pobj��������("��������"))
    If lstrTmp <> "" Then
        If Right(lstrTmp, 2) = "����" Then
            lstrTmp = Left(lstrTmp, Len(lstrTmp) - 2)
        End If
    End If
    clbl��ҵ���.Caption = lstrTmp
    
    cbccMain.Value = pobj��������("ϵͳ���")
    
    '���������󣬻�ȡ��Ƭ��
    Dim lobj��� As Object
    Set lobj��� = CreateObject("������.clsMedicalExam")
    lobj���.ϵͳ��� = pobj��������("ϵͳ���")
    
    '��ʾ��Ƭ��
    If Not lobj���.�����Ա.��Ƭ Is Nothing Then
        cimgPhoto.Picture = lobj���.�����Ա.��Ƭ
    End If
    
    Exit Sub
errhandler:
    sfsub������ "ְҵ���������", "frm��ӡ���ǼǱ�", "Form_Load", Err.Number, Err.Description, False
    Exit Sub
    Resume
End Sub


