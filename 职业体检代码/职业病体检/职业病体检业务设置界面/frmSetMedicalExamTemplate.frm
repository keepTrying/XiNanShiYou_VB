VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSetMedicalExamTemplate 
   Caption         =   "��������"
   ClientHeight    =   7980
   ClientLeft      =   615
   ClientTop       =   1260
   ClientWidth     =   10410
   ClipControls    =   0   'False
   Icon            =   "frmSetMedicalExamTemplate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form4"
   ScaleHeight     =   7980
   ScaleWidth      =   10410
   Begin VB.Frame cfraMedicalTemplateName 
      Appearance      =   0  'Flat
      Caption         =   "��������(ѡ�п����޸�)"
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   6915
      Left            =   0
      TabIndex        =   8
      Top             =   960
      Width           =   2955
      Begin VB.ListBox clstTemplate 
         BackColor       =   &H00FFFFFF&
         Height          =   6540
         Left            =   90
         TabIndex        =   6
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      Caption         =   "��������"
      ForeColor       =   &H80000008&
      Height          =   6915
      Left            =   3000
      TabIndex        =   10
      Top             =   960
      Width           =   9105
      Begin VB.ComboBox tijian_human_leixing 
         Height          =   300
         Left            =   2640
         TabIndex        =   37
         Top             =   480
         Width           =   2415
      End
      Begin VB.ComboBox tijian_leibie 
         Height          =   300
         Left            =   5280
         TabIndex        =   47
         Top             =   480
         Width           =   1695
      End
      Begin VB.CheckBox cchkAnnual 
         Caption         =   "�Ƿ�����"
         Height          =   180
         Left            =   3480
         TabIndex        =   42
         Top             =   1440
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox ctxtLetter 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6960
         MaxLength       =   10
         TabIndex        =   3
         Top             =   1320
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.TextBox ctxtName 
         Height          =   300
         Left            =   120
         MaxLength       =   50
         TabIndex        =   0
         Text            =   "����"
         Top             =   480
         Width           =   2295
      End
      Begin VB.CheckBox cchkAgain 
         Caption         =   "�Ƿ񸴲�����"
         Height          =   180
         Left            =   2040
         TabIndex        =   4
         Top             =   1200
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox ctxtNo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7680
         MaxLength       =   2
         TabIndex        =   2
         Top             =   960
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.ComboBox ccmbSheet 
         Height          =   300
         Left            =   4920
         TabIndex        =   1
         Top             =   1080
         Visible         =   0   'False
         Width           =   855
      End
      Begin TabDlg.SSTab ctabMain 
         Height          =   5430
         Left            =   60
         TabIndex        =   5
         Top             =   1560
         Width           =   8940
         _ExtentX        =   15769
         _ExtentY        =   9578
         _Version        =   393216
         Style           =   1
         Tabs            =   4
         Tab             =   1
         TabsPerRow      =   4
         TabHeight       =   520
         TabCaption(0)   =   "  �����Ŀ  "
         TabPicture(0)   =   "frmSetMedicalExamTemplate.frx":0442
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame1(0)"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "  ������  "
         TabPicture(1)   =   "frmSetMedicalExamTemplate.frx":045E
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Frame1(3)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "  ��������Ϣ  "
         TabPicture(2)   =   "frmSetMedicalExamTemplate.frx":047A
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame1(2)"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "  �շѱ�׼����ϴ������  "
         TabPicture(3)   =   "frmSetMedicalExamTemplate.frx":0496
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Frame1(1)"
         Tab(3).ControlCount=   1
         Begin VB.Frame Frame1 
            Appearance      =   0  'Flat
            ClipControls    =   0   'False
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   5055
            Index           =   0
            Left            =   -75000
            TabIndex        =   32
            Top             =   310
            Width           =   7890
            Begin MSComctlLib.TreeView ctrwSelectedItem 
               Height          =   4365
               Left            =   120
               TabIndex        =   39
               TabStop         =   0   'False
               Top             =   480
               Width           =   3255
               _ExtentX        =   5741
               _ExtentY        =   7699
               _Version        =   393217
               Indentation     =   529
               LabelEdit       =   1
               LineStyle       =   1
               Style           =   7
               Appearance      =   1
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
            Begin VB.CommandButton ccmdDeleteItem 
               Caption         =   ">>"
               Enabled         =   0   'False
               Height          =   285
               Left            =   3435
               Style           =   1  'Graphical
               TabIndex        =   34
               Top             =   1080
               Width           =   650
            End
            Begin VB.CommandButton ccmdAddItem 
               Caption         =   "<<"
               Enabled         =   0   'False
               Height          =   285
               Left            =   3435
               Style           =   1  'Graphical
               TabIndex        =   33
               Top             =   600
               Width           =   650
            End
            Begin MSComctlLib.TreeView ctrwAllItem 
               Height          =   4365
               Left            =   4200
               TabIndex        =   38
               TabStop         =   0   'False
               Top             =   480
               Width           =   3615
               _ExtentX        =   6376
               _ExtentY        =   7699
               _Version        =   393217
               Indentation     =   529
               LabelEdit       =   1
               LineStyle       =   1
               Style           =   7
               Appearance      =   1
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
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��ѡ��������Ŀ��˫�����Լ��룩"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   5
               Left            =   4200
               TabIndex        =   36
               Top             =   240
               Width           =   2880
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ѡ���������Ŀ��˫������ȥ����"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   4
               Left            =   120
               TabIndex        =   35
               Top             =   240
               Width           =   2700
            End
         End
         Begin VB.Frame Frame1 
            Appearance      =   0  'Flat
            ClipControls    =   0   'False
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   4905
            Index           =   3
            Left            =   0
            TabIndex        =   27
            Top             =   310
            Width           =   8865
            Begin MSComctlLib.TreeView ctrwAllConclusion 
               Height          =   4215
               Left            =   4320
               TabIndex        =   40
               Top             =   480
               Width           =   4455
               _ExtentX        =   7858
               _ExtentY        =   7435
               _Version        =   393217
               HideSelection   =   0   'False
               Indentation     =   529
               LabelEdit       =   1
               LineStyle       =   1
               Style           =   7
               FullRowSelect   =   -1  'True
               Appearance      =   1
            End
            Begin VB.CommandButton ccmdDeleteConclusion 
               Caption         =   ">>"
               Height          =   350
               Left            =   3600
               Style           =   1  'Graphical
               TabIndex        =   29
               Top             =   1320
               Width           =   650
            End
            Begin VB.CommandButton ccmdAddConclusion 
               Caption         =   "<<"
               Height          =   350
               Left            =   3600
               Style           =   1  'Graphical
               TabIndex        =   28
               Top             =   780
               Width           =   650
            End
            Begin MSComctlLib.TreeView ctrwSelectedConclusion 
               Height          =   4215
               Left            =   120
               TabIndex        =   41
               Top             =   480
               Width           =   3375
               _ExtentX        =   5953
               _ExtentY        =   7435
               _Version        =   393217
               HideSelection   =   0   'False
               Indentation     =   529
               LabelEdit       =   1
               LineStyle       =   1
               Style           =   7
               FullRowSelect   =   -1  'True
               Appearance      =   1
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "���п�ѡ��������"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   11
               Left            =   4560
               TabIndex        =   31
               Top             =   240
               Width           =   1620
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ѡ���Ľ���"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   10
               Left            =   120
               TabIndex        =   30
               Top             =   240
               Width           =   900
            End
         End
         Begin VB.Frame Frame1 
            Appearance      =   0  'Flat
            ClipControls    =   0   'False
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   5055
            Index           =   2
            Left            =   -75000
            TabIndex        =   18
            Top             =   310
            Width           =   7905
            Begin VB.CommandButton ccmdDown 
               Caption         =   "����"
               Height          =   350
               Left            =   3300
               Style           =   1  'Graphical
               TabIndex        =   24
               Top             =   2280
               Width           =   650
            End
            Begin VB.CommandButton ccmdUp 
               Caption         =   "����"
               Height          =   350
               Left            =   3300
               Style           =   1  'Graphical
               TabIndex        =   23
               Top             =   1800
               Width           =   650
            End
            Begin VB.ListBox clstSelectedBase 
               Height          =   4050
               Left            =   60
               Style           =   1  'Checkbox
               TabIndex        =   22
               Top             =   480
               Width           =   3135
            End
            Begin VB.ListBox clstAllBase 
               BackColor       =   &H00FFFFFF&
               Height          =   3840
               Left            =   4080
               TabIndex        =   21
               Top             =   480
               Width           =   3435
            End
            Begin VB.CommandButton ccmdAddBase 
               Caption         =   "<<"
               Height          =   350
               Left            =   3300
               Style           =   1  'Graphical
               TabIndex        =   20
               Top             =   780
               Width           =   650
            End
            Begin VB.CommandButton ccmdDeleteBase 
               Caption         =   ">>"
               Height          =   350
               Left            =   3300
               Style           =   1  'Graphical
               TabIndex        =   19
               Top             =   1200
               Width           =   650
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ѡ����Ӧ�����Ŀ����¼��ǰ�򹳣�"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   6
               Left            =   120
               TabIndex        =   26
               Top             =   240
               Width           =   3060
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��ѡ�ĸ�����Ŀ"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   7
               Left            =   4080
               TabIndex        =   25
               Top             =   240
               Width           =   1260
            End
         End
         Begin VB.Frame Frame1 
            Appearance      =   0  'Flat
            ClipControls    =   0   'False
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   5025
            Index           =   1
            Left            =   -75000
            TabIndex        =   15
            Top             =   310
            Width           =   7665
            Begin VB.ListBox clstDisposalIdea 
               Height          =   4260
               Left            =   3720
               Style           =   1  'Checkbox
               TabIndex        =   44
               Top             =   480
               Width           =   3255
            End
            Begin VB.TextBox ctxtCharge 
               BackColor       =   &H00FFFFFF&
               Height          =   270
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   17
               Top             =   480
               Width           =   3195
            End
            Begin VB.ListBox clstCharge 
               Height          =   4020
               Left            =   120
               TabIndex        =   16
               Top             =   840
               Width           =   3195
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ѡ����ϴ������"
               Height          =   180
               Index           =   0
               Left            =   3780
               TabIndex        =   45
               Top             =   240
               Width           =   1440
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ѡ���շѱ�׼"
               Height          =   180
               Index           =   1
               Left            =   120
               TabIndex        =   43
               Top             =   240
               Width           =   1080
            End
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����Ա����"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   9
         Left            =   2640
         TabIndex        =   48
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   8
         Left            =   5280
         TabIndex        =   46
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Թ���ĸ��"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   6120
         TabIndex        =   14
         Top             =   1320
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��쵥���ƣ�"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   3720
         TabIndex        =   13
         Top             =   1080
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������ƣ�"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������ţ�"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   6600
         TabIndex        =   11
         Top             =   960
         Visible         =   0   'False
         Width           =   1080
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   2520
      Top             =   480
   End
   Begin MSComctlLib.StatusBar csbMain 
      Align           =   2  'Align Bottom
      Height          =   390
      Left            =   0
      TabIndex        =   9
      Top             =   7590
      Visible         =   0   'False
      Width           =   10410
      _ExtentX        =   18362
      _ExtentY        =   688
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17833
            Key             =   "Msg"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar ctbMain 
      Align           =   1  'Align Top
      Height          =   615
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   10410
      _ExtentX        =   18362
      _ExtentY        =   1085
      ButtonWidth     =   609
      ButtonHeight    =   926
      Appearance      =   1
      ImageList       =   "cimg��ťͼ��"
      _Version        =   393216
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   12
      Left            =   0
      TabIndex        =   49
      Top             =   0
      Width           =   720
   End
End
Attribute VB_Name = "frmSetMedicalExamTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'���ߣ�����

Private mobj����ģ�� As Object               '��ǰ���ڲ���������ģ�塣
Private WithEvents mobjGUI  As cls����ͨ�ö��� '����ͨ�ö���������ʼ����������
Attribute mobjGUI.VB_VarHelpID = -1

Private mblnInUse As Boolean                   '��Ӧ����pblnInUse��
Private mblnSys As Boolean

'���ܣ�������ǰ�����Ƿ��Ѽ��ء�
Public Property Get pblnInUse() As Boolean
    pblnInUse = mblnInUse
End Property

Private Sub cchkAgain_Click()
    On Error Resume Next
    If cchkAgain.Value = 1 Then
        cchkAnnual.Value = 0
        cchkAnnual.Enabled = False
    Else
        cchkAnnual.Enabled = True
    End If
    
End Sub

Private Sub cchkAgain_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        clstDisposalIdea.SetFocus
    End If
End Sub

Private Sub ccmbSheet_GotFocus()
    On Error Resume Next
    If ccmbSheet.Text = "" And ccmbSheet.ListCount > 0 Then
        ccmbSheet.ListIndex = 0
    End If
End Sub

Private Sub clstCharge_Click()
    Dim lstrTemp As String
    
    On Error Resume Next
    If mblnSys Then Exit Sub
    
    lstrTemp = ""
    If clstCharge.ListIndex <> -1 Then
        lstrTemp = clstCharge.List(clstCharge.ListIndex)
        If InStr(lstrTemp, " ") > 0 Then lstrTemp = Left(lstrTemp, InStr(lstrTemp, " ") - 1)
    End If
    ctxtCharge.Text = lstrTemp
    
End Sub



Private Sub ctabMain_Click(PreviousTab As Integer)
    subResizeTab
End Sub

Private Sub ctrwAllConclusion_DblClick()
    On Error Resume Next
    If ccmdAddConclusion.Enabled Then
        If Not ctrwAllConclusion.SelectedItem.Parent Is Nothing Then
            ccmdAddConclusion_click
        End If
    End If
End Sub

Private Sub ctrwAllConclusion_NodeClick(ByVal Node As MSComctlLib.Node)
    ccmdAddConclusion.Enabled = True
End Sub

Private Sub ctrwAllItem_DblClick()
    On Error Resume Next
    ccmdAddItem_Click
End Sub



Private Sub ctrwAllItem_NodeClick(ByVal Node As MSComctlLib.Node)
    On Error Resume Next
    ccmdAddItem.Enabled = True
End Sub

Private Sub ctrwSelectedConclusion_DblClick()
    On Error Resume Next
    If ccmdDeleteConclusion.Enabled Then
        If Not ctrwSelectedConclusion.SelectedItem.Parent Is Nothing Then
            ccmdDeleteConclusion_click
        End If
    End If
End Sub

Private Sub ctrwSelectedConclusion_NodeClick(ByVal Node As MSComctlLib.Node)
    ccmdDeleteConclusion.Enabled = True
End Sub

Private Sub ctrwSelectedItem_DblClick()
    On Error Resume Next
    ccmdDeleteItem_Click

End Sub

Private Sub ctrwSelectedItem_NodeClick(ByVal Node As MSComctlLib.Node)
    On Error Resume Next
    ccmdDeleteItem.Enabled = True
End Sub

Private Sub ctxtLetter_KeyPress(KeyAscii As Integer)
    '�������뺺�֣�������Ctrl-V��
    On Error Resume Next
    If KeyAscii < 0 Or KeyAscii = 22 Then
        KeyAscii = 0
    End If
End Sub

Private Sub ctxtNo_KeyPress(KeyAscii As Integer)
    '�������뺺�֣�������Ctrl-V��
    On Error Resume Next
    If KeyAscii < 0 Or KeyAscii = 22 Then
        KeyAscii = 0
    End If

End Sub

Private Sub Form_Activate()
    On Error Resume Next
    ctxtName.SetFocus
End Sub

'���ܣ����ش���ʱ,��ʼ�������ϵĿؼ�(״̬��,������,���ı���.)
'      ��Ϊ������Ŀ�б��,�����Ŀ�б��,�������б��'���б���Ŀ���ݻ��������Ĳ�ͬ���б仯,
'      ����������ģ���б���click�¼��г�ʼ����
Private Sub Form_Load()
    Dim lcolInfo As Collection
    
    On Error GoTo errHandler
    If mblnInUse Then Exit Sub
    
    MousePointer = 11
    csbMain.Panels(1).Text = "�������ڳ�ʼ�������Ժ�.."
    
    '������ʱ���ܲ�����
    Me.Enabled = False
        
    '��������ͨ�ö���ͨ���ö����ʼ����������
    Set lcolInfo = New Collection
    With lcolInfo
        .Add "����(&A)102"
        .Add "|"
        .Add "����"
        .Add "ɾ��"
        .Add "����(&C)118"
        .Add "|"
        .Add "�˳�"
    End With
    Set mobjGUI = New cls����ͨ�ö���
    With mobjGUI
        Set .Form = Me
        Set .c������ = ctbMain
        .subInitialize lcolInfo, ""
    End With
    
    ctbMain.Buttons(4).Enabled = False
    ctbMain.Buttons(5).Enabled = False
    
    '2012-05-22 ���� ������
    '����Ȩ������
    Dim lobjTmp As Object
    Set lobjTmp = CreateObject("ְҵ������.clsPermissionConfigure")
    If lobjTmp.func���Ҳ���Ȩ��(um�û����, "ְҵ�����_ҵ������_��������_����") = False Then
        ctbMain.Buttons(1).Visible = False
    End If
    If lobjTmp.func���Ҳ���Ȩ��(um�û����, "ְҵ�����_ҵ������_��������_����") = False Then
        ctbMain.Buttons(3).Visible = False
        ctbMain.Buttons(2).Visible = False
    End If
    If lobjTmp.func���Ҳ���Ȩ��(um�û����, "ְҵ�����_ҵ������_��������_����") = False Then
        ctbMain.Buttons(5).Visible = False
        ctbMain.Buttons(6).Visible = False
    End If
    If lobjTmp.func���Ҳ���Ȩ��(um�û����, "ְҵ�����_ҵ������_��������_ɾ��") = False Then
        ctbMain.Buttons(4).Visible = False
    End If
    Set lobjTmp = Nothing
    '2012-05-22 ������
    '��׼�治�ṩ�շѹ��ܡ�
    ctabMain.TabVisible(3) = False
    '���ι��� 'german
    ctabMain.TabVisible(2) = False
    ctabMain.TabVisible(1) = False
    
    '���³�ʼ�������ڶ�ʱ������ɡ�
    Timer1.Enabled = True
    
    mblnInUse = True
    mblnSys = False
    
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����ý���", "frmSetMedicalExamTemplate", "form_load", 6666, lstrError, False
    MousePointer = 0
    csbMain.Panels(1).Text = ""
    '�ָ�������Բ�����
    Me.Enabled = True
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    Frame2.Width = Me.ScaleWidth - Frame2.Left - 60
    cfraMedicalTemplateName.Height = Me.ScaleHeight - cfraMedicalTemplateName.Top - 60
    Frame2.Height = Me.ScaleHeight - Frame2.Top - 60
    ctabMain.Width = Frame2.Width - ctabMain.Left - 60
    ctabMain.Height = Frame2.Height - ctabMain.Top - 60
    
    subResizeTab
End Sub

Private Sub subResizeTab()
    On Error Resume Next
    Select Case ctabMain.Tab
    Case 0
        Frame1(0).Width = ctabMain.Width - Frame1(0).Left - 60
        Frame1(0).Height = ctabMain.Height - Frame1(0).Top - 60
        ctrwSelectedItem.Height = Frame1(0).Height - ctrwSelectedItem.Top - 120
        ctrwAllItem.Height = Frame1(0).Height - ctrwAllItem.Top - 120
        
    Case 1
        Frame1(3).Width = ctabMain.Width - Frame1(3).Left - 60
        Frame1(3).Height = ctabMain.Height - Frame1(3).Top - 60
        ctrwSelectedConclusion.Height = Frame1(3).Height - ctrwSelectedConclusion.Top - 120
        ctrwAllConclusion.Height = Frame1(3).Height - ctrwAllConclusion.Top - 120
    Case 2
        Frame1(2).Width = ctabMain.Width - Frame1(2).Left - 60
        Frame1(2).Height = ctabMain.Height - Frame1(2).Top - 60
        clstSelectedBase.Height = Frame1(2).Height - clstSelectedBase.Top - 120
        clstAllBase.Height = Frame1(2).Height - clstAllBase.Top - 120
    Case 3
        Frame1(1).Width = ctabMain.Width - Frame1(1).Left - 60
        Frame1(1).Height = ctabMain.Height - Frame1(1).Top - 60
        clstCharge.Height = Frame1(1).Height - clstCharge.Top - 120
        clstDisposalIdea.Height = Frame1(1).Height - clstDisposalIdea.Top - 120
    End Select
End Sub

'���ܣ�Ϊ����ߴ���ļ����ٶȣ��Ѳ��ֳ�ʼ���������ڶ�ʱ������ɡ�
Private Sub Timer1_Timer()
    Dim lobj����ģ�弯  As Object
    Dim lobj�����Ŀ�� As Object
    Dim lobjItem As Object
    Dim lobjRec As Object           'ִ��sql���Ľ����
    Dim lcolInfo As New Collection
    Dim objFrame As Frame
    Dim i As Integer
    
    On Error GoTo errHandler
    
    Timer1.Enabled = False
    
    '��������ģ�����
    Set mobj����ģ�� = CreateObject("ְҵ������.ClsMedicalExamTemplate")
    
    '��������ģ�弯���󣬴Ӷ���ȡ��������ģ�����ơ�
    Set lobj����ģ�弯 = CreateObject("ְҵ������.ClsMedicalExamTemplateSet")
    'lobj����ģ�弯.�������� = 1
    Set lcolInfo = lobj����ģ�弯.Ԫ�ؼ�
    
    '��ʼ�����������б��
    clstTemplate.Clear
    For i = 1 To lcolInfo.Count
        clstTemplate.AddItem lcolInfo(i)
    Next i
    
    'ͨ���ֵ�����ȡ������ϴ������������ʼ����ϴ�������б��
    clstDisposalIdea.Clear
    Set lobjRec = pobjDict.Fetch("��ϴ�������ֵ�")
    If lobjRec Is Nothing Then
        Err.Raise 6666, , "ʹ���ֵ�������.Fetch��������������ע�ᡰ�ֵ����.dll��"
    Else
        While Not lobjRec.EOF
            clstDisposalIdea.AddItem lobjRec("����")
            lobjRec.movenext
        Wend
    End If
    '���õ�ǰ����ģ��Ϊ�������ƿ�ĵ�һ�
    If clstTemplate.ListCount = 0 Then
        '�������в��־����ɲ�����
        For Each objFrame In Frame1
            objFrame.Enabled = False
        Next
        ccmbSheet.Enabled = False
        ctxtNo.Enabled = False
        ctxtLetter.Enabled = False
        cchkAgain.Enabled = False
        ctbMain.Buttons(3).Enabled = False
        ctbMain.Buttons(4).Enabled = False
        ctbMain.Buttons(5).Enabled = False
    Else
        'clstTemplate.ListIndex = 0
    End If
    
    '��ȡ������쵥���͡�
    Set lcolInfo = pobjҵ�����.func��ȡ������쵥���� 'CreateObject("ְҵ������.clsManageMedicalExam")
    ccmbSheet.Clear
    For i = 1 To lcolInfo.Count
        ccmbSheet.AddItem lcolInfo(i)
    Next
    
    '�޸ģ�2001-11-2�������ѡ�������Ŀ��ʾ��һ�����С�
    Set lobj�����Ŀ�� = CreateObject("ְҵ������.clsTestItemSet")
    
    '��ȡ���������ࡢ�����Ŀ��
    Set lobjRec = pobjDict.Fetch("ְҵ���������ֵ�")
    '��ʾ��������ctrvItem�У����нڵ��key=������id����
    ctrwAllItem.Nodes.Clear
    Do While Not lobjRec.EOF
        'ͨ��"lobj�����Ŀ��"���λ�ȡ������������Ŀ��
        lobj�����Ŀ��.������ = lobjRec("InnerID")
        Set lobjItem = lobj�����Ŀ��.�����Ŀ
        If Not lobjItem.EOF Then
            ctrwAllItem.Nodes.Add , , "I" & lobjRec("InnerID"), lobjRec("���") & " " & lobjRec("����")
        End If
        '��ʾ�����Ŀ��ctrvItem�У����нڵ��key=����,parent=�����ࣩ��
        Do While Not lobjItem.EOF
            ctrwAllItem.Nodes.Add "I" & lobjRec("InnerID"), tvwChild, "I" & lobjItem("����"), lobjItem("����") & " " & lobjItem("����")
            lobjItem.movenext
        Loop
        
        lobjRec.movenext
    Loop
    
    'ͨ���ֵ���󣬳�ʼ�����������۴���
    Set lobjRec = pobjDict.Fetch("�������ֵ�", "Parent=0")
    If lobjRec Is Nothing Then
        Err.Raise 6666, , "ʹ���ֵ�������.Fetch������ȡ�������ֵ�����ʱ����������ע�ᡰ�ֵ����.dll��"
    End If
    ctrwAllConclusion.Nodes.Clear
    While Not lobjRec.EOF
        'key:R+InnderID��
        ctrwAllConclusion.Nodes.Add , , "R" & lobjRec("InnerID").Value, lobjRec("����").Value
        lobjRec.movenext
    Wend
    
    '��ȡ�����ۡ�
    Set lobjRec = pobjDict.Fetch("�������ֵ�", "Parent<>0")
    If lobjRec Is Nothing Then
        Err.Raise 6666, , "ʹ���ֵ�������.Fetch������ȡ�������ֵ�����ʱ����������ע�ᡰ�ֵ����.dll��"
    End If
    While Not lobjRec.EOF
        'key:I+InnerID��
        On Error Resume Next
        ctrwAllConclusion.Nodes.Add "R" & lobjRec("Parent").Value, tvwChild, "I" & lobjRec("InnerID").Value, lobjRec("����").Value
        On Error GoTo errHandler
        lobjRec.movenext
    Wend
    
    
    If ctabMain.TabVisible(3) Then
        '��ȡ�����շѱ�׼�����ƣ��ܶ
        Set lcolInfo = pobjҵ�����.��������շѱ�׼
        mblnSys = True
        clstCharge.Clear
        For i = 1 To lcolInfo.Count
            clstCharge.AddItem Format(lcolInfo(i)("����"), String(50, " ")) & " " & lcolInfo(i)("�ܶ�")
        Next
        mblnSys = False
    End If
    
    '�޸ģ�2003-6-27���������ҵ�����á��Ƿ�ʹ����쵥����
    If pobjҵ�����.ҵ������("�Ƿ�ʹ����쵥") <> "��" Then
        'Label1(0).Visible = True   '��ʱ����-----------------------
        'ccmbSheet.Visible = True   '��ʱ����-----------------------
        'Label2.Top = 1320          '��ʱ����-----------------------
    End If

    
    '��ʼ��������״̬��
    Dim lblnCancel As Boolean
    mobjGUI_BeforeOperate "����", lblnCancel
    
    '�ָ�������Բ�����
    Me.Enabled = True
    MousePointer = 0
    csbMain.Panels(1).Text = ""
    ctxtName.SetFocus
    
    'german
    '��ֲ��:��Х��
    '����:3.16
    '���ֵ���л��ѡȡ������
    Set lobjRec = pobjDict.FetchEx("��������ֵ�")
    tijian_leibie.Clear
    tijian_leibie.AddItem ""
    For i = 1 To lobjRec.recordcount
        tijian_leibie.AddItem lobjRec("����")
        tijian_leibie.ItemData(tijian_leibie.NewIndex) = lobjRec("���")
        lobjRec.movenext
    Next
    tijian_leibie.ListIndex = 0
    
    Set lobjRec = pobjDict.FetchEx("���������ֵ�")
    tijian_human_leixing.Clear
    tijian_human_leixing.AddItem ""
    For i = 1 To lobjRec.recordcount
        tijian_human_leixing.AddItem lobjRec("����")
        tijian_human_leixing.ItemData(tijian_human_leixing.NewIndex) = lobjRec("���")
        lobjRec.movenext
    Next
    tijian_human_leixing.ListIndex = 0
    
    Exit Sub
    
errHandler:
    Dim lstrError  As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����ý���", "frmSetMedicalExamTemplate", "Timer1_Timer", 6666, lstrError, False
    mblnSys = False
    MousePointer = 0
    csbMain.Panels(1).Text = ""
    '�ָ�������Բ�����
    Me.Enabled = True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 39 Then
        '���������롰'����
        KeyAscii = 0
    End If

End Sub
'����:���һ��Ŀ�����Ŀ
Private Sub ccmdAddItem_Click()
    Dim lstrCode As String '�����Ŀ�ı��롣
    Dim i As Long
    Dim llngIndex As Long
    Dim llngChildren As Long
    
    On Error GoTo errHandler
    
    If ctrwAllItem.SelectedItem Is Nothing Then Exit Sub
    
    If ctrwAllItem.SelectedItem.Parent Is Nothing Then
        '���һ����Ŀ��
        llngChildren = ctrwAllItem.SelectedItem.Children  '�ӽڵ�����
        llngIndex = ctrwAllItem.SelectedItem.Child.Index  '��һ���ӽڵ��������
        For i = llngIndex To llngIndex + llngChildren - 1
            lstrCode = ctrwAllItem.Nodes(i).Key                 '�����Ŀ�ı��롣
            lstrCode = Right(lstrCode, Len(lstrCode) - 1)
            sub��ӵ��������Ŀ lstrCode
        Next
    Else
        '��ӵ�����Ŀ��
        lstrCode = ctrwAllItem.SelectedItem.Key
        lstrCode = Right(lstrCode, Len(lstrCode) - 1)
        
        sub��ӵ��������Ŀ lstrCode
    End If
    ccmdAddItem.Enabled = False
    ccmdDeleteItem.Enabled = False
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "����ģ������", "frmSetMedicalExamTemplate", "ccmdAddItem_Click", 6666, lstrError, False
End Sub

'���룺paraItemCode �����Ŀ���롣
Private Sub sub��ӵ��������Ŀ(ByVal paraItemCode As String)
    Dim lobj�����Ŀ As Object
    
    '����������������Ŀ��
    mobj����ģ��.Sub��������Ŀ paraItemCode
    
    Set lobj�����Ŀ = mobj����ģ��.�����Ŀ��(paraItemCode)
    
    '��ѡ�������Ŀ���������Ŀ�������ظ�����Ĵ���
    On Error Resume Next
    '����һ���ֽڵ�.
    ctrwSelectedItem.Nodes.Add , , "I" & lobj�����Ŀ.������, ctrwAllItem.Nodes("I" & lobj�����Ŀ.������).Text
    '������Ŀ��
    ctrwSelectedItem.Nodes.Add "I" & lobj�����Ŀ.������, tvwChild, "I" & lobj�����Ŀ.����, lobj�����Ŀ.���� & " " & lobj�����Ŀ.����

End Sub
Private Sub ccmdDeleteItem_Click()
    'ɾ��ѡ�е�һ��Ŀ�����Ŀ����ǰ���ģ��������Ŀ�б��
    Dim lintSelectedIndex As Integer
    Dim lstrCode As String
    Dim lobjParent As Node
    Dim llngChildren  As Long
    Dim llngIndex As Long
    Dim i As Long
    
    On Error GoTo errHandler
    If ctrwSelectedItem.SelectedItem Is Nothing Then Exit Sub
    
    If ctrwSelectedItem.SelectedItem.Parent Is Nothing Then
        'ɾ��һ�ࡣ
        llngChildren = ctrwSelectedItem.SelectedItem.Children  '�ӽڵ�����
        llngIndex = ctrwSelectedItem.SelectedItem.Child.Index  '��һ���ӽڵ��������
        For i = llngIndex To llngIndex + llngChildren - 1
            lstrCode = ctrwSelectedItem.Nodes(i).Key                 '�����Ŀ�ı��롣
            lstrCode = Right(lstrCode, Len(lstrCode) - 1)
            mobj����ģ��.Subɾ�������Ŀ lstrCode
        Next
        
         'ɾ��������ڵ㡣
        ctrwSelectedItem.Nodes.Remove ctrwSelectedItem.SelectedItem.Key
    Else
        'ɾ��������Ŀ��
        lstrCode = ctrwSelectedItem.SelectedItem.Key
        lstrCode = Right(lstrCode, Len(lstrCode) - 1)
        mobj����ģ��.Subɾ�������Ŀ lstrCode
        
        'ɾ��ѡ����Ŀ������Ŀ��
        Set lobjParent = ctrwSelectedItem.SelectedItem.Parent
        ctrwSelectedItem.Nodes.Remove ctrwSelectedItem.SelectedItem.Key
        
        '��ĳ���౻ɾ����ϣ�ɾ��������ڵ㡣
        If lobjParent.Children = 0 Then
            ctrwSelectedItem.Nodes.Remove lobjParent.Key
        End If
    End If
    ccmdDeleteItem.Enabled = False
    
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "����ģ������", "frmSetMedicalExamTemplate", "ccmdDeleteItem_Click", 6666, lstrError, False
End Sub

'���ܣ����һ�������ۡ�
Private Sub ccmdAddConclusion_click()
    Dim lobjNode As Node
    Dim llngChildren  As Long
    Dim llngIndex As Long
    Dim i As Long
    
    On Error GoTo errHandler
    If ctrwAllConclusion.SelectedItem Is Nothing Then Exit Sub
    
    
    Set lobjNode = ctrwAllConclusion.SelectedItem
    
    If lobjNode.Parent Is Nothing Then
        'ѡ����ࡣ
        
        llngChildren = lobjNode.Children  '�ӽڵ�����
        
        If llngChildren > 0 Then
            llngIndex = lobjNode.Child.Index  '��һ���ӽڵ��������
            On Error Resume Next
            '����ϼ��ڵ㡣
            ctrwSelectedConclusion.Nodes.Add , , lobjNode.Key, lobjNode.Text
            
            '������Ӹô��������������ۡ�
            For i = llngIndex To llngIndex + llngChildren - 1
                Set lobjNode = ctrwAllConclusion.Nodes(i)
                '�ڶ��������ѡ�е������ۡ�
                mobj����ģ��.sub��������� Right(lobjNode.Key, Len(lobjNode.Key) - 1)
                '���Ҷ�ڵ㡣
                ctrwSelectedConclusion.Nodes.Add lobjNode.Parent.Key, tvwChild, lobjNode.Key, lobjNode.Text
            Next
        End If
    Else 'ѡ��Ҷ�ڵ㡣
        
        '�ڶ��������ѡ�е������ۡ�
        mobj����ģ��.sub��������� Right(lobjNode.Key, Len(lobjNode.Key) - 1)
        
        '��ѡ���������б������ӡ�
        On Error Resume Next
        '����ϼ��ڵ㡣
        ctrwSelectedConclusion.Nodes.Add , , lobjNode.Parent.Key, lobjNode.Parent.Text
        '���Ҷ�ڵ㡣
        ctrwSelectedConclusion.Nodes.Add lobjNode.Parent.Key, tvwChild, lobjNode.Key, lobjNode.Text
        On Error GoTo errHandler
        
    End If
    
    '��ѡ�������۳���10������ʾ��
    If ctrwSelectedConclusion.Nodes.Count >= 20 Then
        sffuncMsg "��ע�⣬��ѡ������ô��������ۣ����ܵ���������¼��������󱣴��ʱ���ܳ������Զ��������ۣ���������ѡ�����������µ������ۡ�", sf����
    End If
    
    ccmdAddConclusion.Enabled = False
    
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "����ģ������", "frmSetMedicalExamTemplate", "ccmdAddConclusion_click", 6666, lstrError, False
End Sub
'���ܣ�ɾ��һ�������ۡ�
Private Sub ccmdDeleteConclusion_click()
    Dim lobjNode As Node
    Dim llngChildren As Long
    Dim llngIndex As Long
    Dim i As Long
    
    On Error GoTo errHandler
    If ctrwSelectedConclusion.SelectedItem Is Nothing Then Exit Sub
    
    Set lobjNode = ctrwSelectedConclusion.SelectedItem
    
    If lobjNode.Parent Is Nothing Then
        'ɾ�����ࡣ
        
        llngChildren = lobjNode.Children  '�ӽڵ�����
        If llngChildren > 0 Then
            llngIndex = lobjNode.Child.Index  '��һ���ӽڵ��������
            
            '������Ӹô��������������ۡ�
            For i = llngIndex To llngIndex + llngChildren - 1
                '�ڶ�����ɾ��ѡ�е������ۡ�
                mobj����ģ��.subɾ�������� Right(ctrwSelectedConclusion.Nodes(i).Key, Len(ctrwSelectedConclusion.Nodes(i).Key) - 1)
            Next
        End If
    Else 'ɾ��Ҷ�ڵ㡣
    
        '�ڶ��������ѡ�е������ۡ�
        mobj����ģ��.subɾ�������� Right(lobjNode.Key, Len(lobjNode.Key) - 1)
        
        '���ϼ��ڵ��Ҷ�ڵ㱻ɾ�����ˣ���������Ҫ��ɾ����
        If lobjNode.Parent.Children = 1 Then
            Set lobjNode = lobjNode.Parent
        End If
        
    End If
    
    On Error Resume Next
    '��ѡ����������ɾ����
    ctrwSelectedConclusion.Nodes.Remove lobjNode.Key
    
    ccmdDeleteConclusion.Enabled = False
    
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "����ģ������", "frmSetMedicalExamTemplate", "ccmdDeleteConclusion_click", 6666, lstrError, False
    
End Sub

'���ܣ���Ӹ�����Ŀ��
Private Sub ccmdAddBase_click()
    Dim llngAllIndex As Long
    On Error GoTo errHandler
    
    '�ڶ�������Ӹ�����Ŀ��
    mobj����ģ��.sub��Ӹ�����Ŀ clstAllBase.List(clstAllBase.ListIndex), False
    
    '��ѡ�и�����Ŀ�б���������Ŀ��
    With clstSelectedBase
        .AddItem clstAllBase.List(clstAllBase.ListIndex)
        .Selected(.NewIndex) = False
        .ListIndex = .NewIndex
    End With
    
    '�����и�����Ŀ�б����ɾ����Ŀ��
    With clstAllBase
        llngAllIndex = .ListIndex
        .RemoveItem llngAllIndex
        If .ListCount = 0 Then
            ccmdAddBase.Enabled = False
        ElseIf .ListCount > llngAllIndex Then
            .ListIndex = llngAllIndex
        Else
            .ListIndex = .ListCount - 1
        End If
    End With
    
    ccmdDeleteBase.Enabled = True
    
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "����ģ������", "frmSetMedicalExamTemplate", "ccmdAddBase_click", 6666, lstrError, False
End Sub

'���ܣ�ɾ��������Ŀ��
Private Sub ccmdDeleteBase_click()
    On Error GoTo errHandler
    
    '�ڶ�����ɾ��������Ŀ��
    mobj����ģ��.subɾ��������Ŀ clstSelectedBase.List(clstSelectedBase.ListIndex)
    
    '�����и�����Ŀ�б����ɾ����Ŀ��
    clstAllBase.AddItem clstSelectedBase.List(clstSelectedBase.ListIndex)
    clstAllBase.ListIndex = clstAllBase.NewIndex
    
    '��ѡ�и�����Ŀ�б����ɾ����Ŀ��
    clstSelectedBase.RemoveItem clstSelectedBase.ListIndex
    If clstSelectedBase.ListCount > 0 Then
        clstSelectedBase.ListIndex = 0
    Else
        ccmdDeleteBase.Enabled = False
    End If
    ccmdAddBase.Enabled = True
    
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "����ģ������", "frmSetMedicalExamTemplate", "ccmdDeleteBase_click", 6666, lstrError, False
End Sub

Private Sub clstAllConclusion_DblClick()
    On Error Resume Next
    ccmdAddConclusion_click
End Sub



Private Sub clstSelectedBase_Click()
    On Error GoTo errHandler
    
    '�ı�ĳһ������Ŀ�� "�Ƿ��¼"״̬��
    With clstSelectedBase
        mobj����ģ��.sub��Ӹ�����Ŀ .List(.ListIndex), .Selected(.ListIndex)
    End With
    
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "����ģ������", "frmSetMedicalExamTemplate", "clstSelectedBase_Click", 6666, lstrError, False
End Sub

'���ܣ�ͬɾ��������Ŀ��
Private Sub clstSelectedBase_DblClick()
    On Error Resume Next
    ccmdDeleteBase_click
End Sub

'���ܣ�ͬ��Ӹ�����Ŀ��
Private Sub clstAllBase_dblclick()
    On Error Resume Next
    ccmdAddBase_click
End Sub

'���ܣ�ͬɾ�������ۡ�
Private Sub clstSelectedConclusion_DblClick()
    On Error Resume Next
    ccmdDeleteConclusion_click
End Sub


'���ܣ�ͨ���Ե�ǰ����ģ��Ĺؼ�������ֵ,����������������,����ʾ�ڽ����ϡ�
Private Sub clstTemplate_Click()
    Dim objFrame As Frame
    Dim lobjRec As Object
    
    On Error GoTo errHandler
    If clstTemplate.ListIndex = -1 Then Exit Sub
    
    MousePointer = 11
    csbMain.Panels(1).Text = "���ڻ�ȡ����ģ��������Ϣ�����Ժ�..."
    
    'Ϊ����ģ����������������ʼ���������Ӷ���ȡ��������������ԡ�
    mobj����ģ��.������ = clstTemplate.List(clstTemplate.ListIndex)
    
    'german
    '�����:��Х��
    '���ܣ�Ԥ�Ƚ��Ѿ����ڵ���Ϣ��������ڽ����ϣ�Ӧ��������ӵĿؼ�
    Set lobjRec = dafuncGetData("select * from ְҵ�����_����ģ�������Ϣ�� where(��������='" + mobj����ģ��.������ + "')")
    If (lobjRec.recordcount > 0) Then
        tijian_leibie.Text = lobjRec("������")
        tijian_human_leixing.Text = lobjRec("�����Ա����")
    Else
        MsgBox "���ݶ�ȡ�������ش���", 16, "��Ϣ"
        Exit Sub
    End If
    
    '������ģ������������ʾ�ڽ����ϡ�
    subShowTemplate
    
    'ѡ��������ģ������,�����ϵĿؼ���Ϊ�ɲ���״̬.
    For Each objFrame In Frame1
        objFrame.Enabled = True
    Next
    'modify by lanchao 2015-03-15 ��Buttons(3).Enabled = false -->Buttons(3).Enabled = True
    ctbMain.Buttons(3).Enabled = True
    ctbMain.Buttons(4).Enabled = True
    ctbMain.Buttons(5).Enabled = True
    
    MousePointer = 0
    csbMain.Panels(1).Text = ""
    Exit Sub
    
errHandler:
    Dim lstrError  As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "����ģ������", "frmSetMedicalExamTemplate", "clstTemplate_Click", 6666, lstrError, False
    
    MousePointer = 0
    csbMain.Panels(1).Text = "��ȡ����ģ��������Ϣʧ�ܡ�"
    Exit Sub
    Resume
End Sub

Private Sub ctxtLetter_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errHandler
    If KeyCode = 13 Then
        cchkAgain.SetFocus
    End If
    Exit Sub
errHandler:
    
End Sub


Private Sub ctxtName_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        If ccmbSheet.Visible Then
            ccmbSheet.Enabled = True
            ccmbSheet.SetFocus
        Else
            ctxtNo.SetFocus
        End If
    End If

End Sub

'����������ģ������,��������������ģ�����仯������ʾ��ͬ������.���������������Ѿ�����,Ч����ͬ�ڵ��
'�����б���е�������.���������,����Ϊ�½�����.
Private Sub ctxtName_LostFocus()
    On Error Resume Next
    If clstTemplate.ListIndex = -1 Then
        Frame2.Caption = "��������" & Trim(ctxtName)
    Else
        Frame2.Caption = "�޸�����" & Trim(ctxtName)
    End If
    ctbMain.Buttons(3).Enabled = True
End Sub

Private Sub ctxtNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Err.Clear
    On Error GoTo errorHandler
    If KeyCode = 13 Then
        ctxtLetter.SetFocus
    End If
    Exit Sub
errorHandler:
    
End Sub

Private Sub ccmbSheet_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errHandler
    If KeyCode = 13 Then
        ctxtNo.SetFocus
    End If
    Exit Sub
errHandler:
    
End Sub

'���ܣ���ս��档
Private Sub subClear()
    clstTemplate.ListIndex = -1
    ctxtName = ""
    ctbMain.Buttons(3).Enabled = True
    ctbMain.Buttons(4).Enabled = False
    ctbMain.Buttons(5).Enabled = False
    ctxtName.Enabled = True
    ccmbSheet.Enabled = True
    ctxtNo.Enabled = True
    ctxtLetter.Enabled = True
    cchkAgain.Enabled = True
    mobj����ģ��.������ = ""
    Frame2.Caption = "��������"
    On Error Resume Next
    ctxtName.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set mobj����ģ�� = Nothing
    Set mobjGUI = Nothing
    mblnInUse = False
End Sub

'���ܣ�ѡ����츽����Ŀ���ơ�
Private Sub ccmdUp_Click()
    Dim lstrItem As String
    Dim lblnSelected As Boolean
    Dim llngIndex As Long
    
    On Error GoTo errHandler
    
    '��츽����Ŀ���ơ�
    With clstSelectedBase
        llngIndex = .ListIndex
        
        '�ڶ��������Ƹ�����Ŀ��
        mobj����ģ��.sub������Ŀ���� .List(llngIndex)
        
        If llngIndex > 0 Then
            '�ȼ�¼ѡ����Ŀ���ݡ��Ƿ�ѡ�С�
            lstrItem = .List(llngIndex)
            lblnSelected = .Selected(llngIndex)
            '���Ƴ���
            .RemoveItem llngIndex
            '�ټ��롣
            .AddItem lstrItem, llngIndex - 1
            If lblnSelected Then
                .Selected(llngIndex - 1) = True
            End If
        End If
    End With
    
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����ý���", "frmSetMedicalExamTemplate", "ccmdUp_Click", 6666, lstrError, False
End Sub

'���ܣ�ѡ����츽����Ŀ���ơ�
Private Sub ccmdDown_Click()
    Dim lstrItem As String
    Dim lblnSelected As Boolean
    Dim llngIndex As Long
    
    On Error GoTo errHandler
    
    '��츽����Ŀ���ơ�
    With clstSelectedBase
        llngIndex = .ListIndex
        
        '�ڶ��������Ƹ�����Ŀ��
        mobj����ģ��.sub������Ŀ���� .List(llngIndex)
        
        If llngIndex < .ListCount - 1 Then
            '�ȼ�¼ѡ����Ŀ���ݡ��Ƿ�ѡ�С�
            lstrItem = .List(llngIndex)
            lblnSelected = .Selected(llngIndex)
            '���Ƴ���
            .RemoveItem llngIndex
            '�ټ��롣
            .AddItem lstrItem, llngIndex + 1
            If lblnSelected Then
                .Selected(llngIndex + 1) = True
            End If
            .ListIndex = llngIndex + 1
        End If
    End With
    
    Exit Sub
errHandler:
    Dim lstrError  As String
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "ְҵ�����ý���", "frmSetMedicalExamTemplate", "ccmdUp_Click", 6666, lstrError, False

End Sub


'���ܣ����������ϵİ�ť��
Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    Dim objFrame As Object
    Dim lstrTemp As String
    Dim lstrError As String
    Dim i As Long
    
    On Error GoTo errHandler
    
    Select Case Operate
        Case "����"
            clstTemplate.ListIndex = -1
            ctxtName = ""
            For Each objFrame In Frame1
                objFrame.Enabled = True
            Next
            ctbMain.Buttons(3).Enabled = True
            ctbMain.Buttons(4).Enabled = False
            ctbMain.Buttons(5).Enabled = False
            ctxtName.Enabled = True
            ccmbSheet.Enabled = True
            ctxtNo.Enabled = True
            ctxtLetter.Enabled = True
            cchkAgain.Enabled = True
            mobj����ģ��.������ = ""
            subShowTemplate
            Frame2.Caption = "��������"
            On Error Resume Next
            ctxtName.SetFocus
            
        Case "ɾ��"
            'ѯ�ʡ�
            If sffuncMsg("��ȷ��Ҫɾ������ģ�塰" & mobj����ģ��.������ & "����", sfѯ��) Then
                '�ӿ���ɾ����
                mobj����ģ��.Subɾ��ģ��
                '�ӽ�����ɾ����
                clstTemplate.RemoveItem clstTemplate.ListIndex
                If clstTemplate.ListCount > 0 Then
                    clstTemplate.ListIndex = 0
                Else
                    '������
                    subClear
                    ctbMain.Buttons(4).Enabled = False
                    ctbMain.Buttons(5).Enabled = False
                    ctbMain.Buttons(3).Enabled = False
                    ctxtName.Text = ""
                    ctxtName.SetFocus
                    Frame2.Caption = "��������"
                End If
                
            End If
            Cancel = True
        Case "����"
            Dim lstrNewName As String '���ƺ�����������ơ�
            lstrNewName = Trim(ctxtName.Text) & "1"
            lstrNewName = InputBox("������������(������Ϻ󣬿����޸��������ã������밴���水ť���Ʋ������):", "��������", lstrNewName)
            lstrNewName = Trim(lstrNewName)
            If lstrNewName <> "" Then
                ctxtName.Text = lstrNewName
                mobj����ģ��.Sub����ģ�� lstrNewName
                clstTemplate.ListIndex = -1
                ctxtNo = ""
                ctbMain.Buttons(4).Enabled = False
                ctbMain.Buttons(5).Enabled = False
                Frame2.Caption = "��������" & lstrNewName
                csbMain.Panels("Msg") = "���ƺ�����������밴�����桱��ť���棬���Ʋ�����ϡ��� "
                
            End If
            
            Cancel = True
            
        Case "����"
            Dim lbln�Ƿ��Ѵ��� As Boolean '�жϱ�������������Ƿ����ڿ��д��ڡ�
            Dim lobjFrame As Variant
            lstrError = ""
            If ctrwSelectedItem.Nodes.Count = 0 Then
                lstrError = "û��ѡ�������Ŀ��" & Chr(13) & Chr(10)
            End If
'            If Trim(ctxtNo) = "" Then 'ȥ��ͷ����β���Ŀո�
'                lstrError = lstrError & "û������������š�" & Chr(13) & Chr(10)
'            End If
'            If Trim(ctxtLetter) = "" Then 'ȥ��ͷ����β���Ŀո�
'                lstrError = lstrError & "û�������Թ���ĸ��" & Chr(13) & Chr(10)
'            End If
'            If ctrwSelectedConclusion.Nodes.Count = 0 Then
'                lstrError = lstrError & "û��ѡ�������ۡ�" & Chr(13) & Chr(10)
'            End If
            '�ж����������Ƿ�Ψһ��
            Dim lobjTemp As Object
            Set lobjTemp = CreateObject("ְҵ������.ClsMedicalExamTemplate")
            
            'MsgBox CStr(clstTemplate.ListIndex), , "Message"
            
            If clstTemplate.ListIndex = -1 Then '������������ģ��Ļ�����ô��ֵΪ-1������ʱ��ģ���б��Ϊʧȥ����״̬
                '������
                lobjTemp.������ = Trim(ctxtName)
                If lobjTemp.�Ƿ��Ѵ��� Then
                    lstrError = lstrError & "���������Ѵ��ڡ�"
                Else
                    '�����������ơ�
                    If lstrError = "" Then
                        mobj����ģ��.sub������������ Trim(ctxtName) 'δ���� �Ƿ����
                        '�����ݿ��еı���ȫ���滻
                    End If
                End If
            ElseIf Trim(clstTemplate.Text) <> Trim(ctxtName.Text) Then
                lobjTemp.������ = Trim(ctxtName)
                If lobjTemp.�Ƿ��Ѵ��� Then
                    lstrError = lstrError & "���������Ѵ��ڡ�"
                Else
                    '�����������ơ�
                    If lstrError = "" Then
                        mobj����ģ��.sub������������ Trim(ctxtName) 'ְҵ������ - clsMedicalExamTemplate
                    End If
                End If
            End If
            If lstrError <> "" Then
                sffuncMsg "ϵͳ�޷����棬��Ϊ��" & Chr(13) & Chr(10) & lstrError, sf����
                Cancel = True
                Exit Sub
            End If
            
            Cancel = True
            If (tijian_human_leixing.Text = "") Then
                MsgBox "�����������Ա����ѡ��ú��ٱ���", 16, "��Ϣ"
                Exit Sub
            End If
                
            If (tijian_leibie.Text = "") Then
                MsgBox "������������ѡ��ú��ٱ���", 16, "��Ϣ"
                Exit Sub
            End If
            
            MousePointer = 11
            csbMain.Panels(1).Text = "���ڱ��棬���Ժ�..."
            
            '������ʱ���ܲ�����
            ctbMain.Enabled = False
            cfraMedicalTemplateName.Enabled = False
            For Each lobjFrame In Frame1
                lobjFrame.Enabled = False
            Next
            ctxtName.Enabled = False
            ctxtLetter.Enabled = False
            ctxtNo.Enabled = False
            cchkAgain.Enabled = False
            
            '��������ģ�����ԡ�
            ccmbSheet.ListIndex = 0 '----------��쵥Ĭ������------------
            With mobj����ģ��
                '.���� = Trim(ctxtNo.Text)
                .���� = "13" 'Ĭ�����δ���
                .��쵥���� = Trim(ctxtName.Text)
                '.�Թ���ĸ��� = Trim(ctxtLetter.Text)
                .�Թ���ĸ��� = "B"
                .�Ƿ񸴲����� = IIf(cchkAgain.Value = 1, True, False)
                
                '�޸ģ�2002-7-26�����ӡ��Ƿ��������ԣ���
                .�Ƿ����� = IIf(cchkAnnual.Value = 1, True, False)
                
                '��ȡ�շѱ�׼���ơ�
                .�շѱ�׼ = ctxtCharge.Text
                .��ϴ������ = Empty
                .�����Ա����_ger = tijian_human_leixing.Text 'german
                .������_ger = tijian_leibie.Text 'german

            End With
            '��ȡѡ�е���ϴ��������
            lstrTemp = ""
            For i = 0 To clstDisposalIdea.ListCount - 1
                If clstDisposalIdea.Selected(i) Then
                    lstrTemp = lstrTemp & clstDisposalIdea.List(i) & ","
                End If
            Next i
            mobj����ģ��.��ϴ������ = lstrTemp
            
            '��������ģ�塣�����½���������������Ƶ��б��С�
            lbln�Ƿ��Ѵ��� = mobj����ģ��.�Ƿ��Ѵ���
            
            mobj����ģ��.sub����ģ��
            
            If Not lbln�Ƿ��Ѵ��� Then
                clstTemplate.AddItem Trim(ctxtName.Text)
                mblnSys = True
                clstTemplate.ListIndex = clstTemplate.NewIndex
                mblnSys = False
                ctxtNo = mobj����ģ��.����
            Else
                '�ж��Ƿ��޸����������ơ�
                If mobj����ģ��.������ <> Trim(ctxtName.Text) Then
                    '�޸Ŀ����������ơ�
                    mobj����ģ��.sub������������ Trim(ctxtName.Text)
                    
                    '�޸��б����������ơ�
                    mblnSys = True
                    i = clstTemplate.ListIndex
                    clstTemplate.RemoveItem i
                    If i < clstTemplate.ListCount - 1 Then
                        clstTemplate.AddItem Trim(ctxtName.Text), i
                    Else
                        clstTemplate.AddItem Trim(ctxtName.Text)
                    End If
                    mblnSys = False
                End If
            End If
            
            MsgBox "����ɹ���", vbOKOnly + vbInformation, "ϵͳ��ʾ"
            subClear
            '�ָ�������Բ�����
            ctbMain.Enabled = True
            cfraMedicalTemplateName.Enabled = True
            For Each lobjFrame In Frame1
                lobjFrame.Enabled = True
            Next
            If clstTemplate.ListIndex >= 0 Then
                If Trim(clstTemplate.Text) <> Trim(ctxtName.Text) Then
                    i = clstTemplate.ListIndex
                    clstTemplate.RemoveItem i
                    If i = clstTemplate.ListCount Then
                        clstTemplate.AddItem Trim(ctxtName.Text)
                    Else
                        clstTemplate.AddItem Trim(ctxtName.Text), i
                    End If
                    clstTemplate.ListIndex = i
                End If
            End If
            ctxtName.Enabled = True
            ctxtLetter.Enabled = True
            ctxtNo.Enabled = True
            cchkAgain.Enabled = True
            
            MousePointer = 0
            csbMain.Panels(1).Text = "����ɹ���"
            Cancel = True
    End Select
    
    Exit Sub
errHandler:
    mblnSys = False
    lstrError = func������(Err.Number, Err.Description)
    sfsub������ "����ģ������", "frmSetmedicalExamTemplate", "mobjGUI_BeforeOperate", 6666, lstrError, False
    If Operate = "����" Then
        '�ָ�������Բ�����
        ctbMain.Enabled = True
        cfraMedicalTemplateName.Enabled = True
        For Each lobjFrame In Frame1
            lobjFrame.Enabled = True
        Next
        ctxtName.Enabled = True
        ctxtLetter.Enabled = True
        ctxtNo.Enabled = True
        cchkAgain.Enabled = True
        MousePointer = 0
        csbMain.Panels(1).Text = "����ʧ�ܣ�"
        Cancel = True
    End If
End Sub

'���ܣ���������ģ������������ʾ����ģ��������Ϣ��
Private Sub subShowTemplate()
    Dim lobj�����Ŀ�� As Object 'ְҵ�����󲿼�.clsTestItemSet
    Dim lobjRec As Object        'ִ��sql���Ľ����
    Dim lcolInfo As Collection   '����ģ�����ġ������ۼ�����������������Ŀ�������������Ŀ���������Լ��ϡ�
    Dim lcolItem As Collection   'lcolInfo�����е�Ԫ�ء�
    Dim lobj�����Ŀ As Object   'ְҵ�����󲿼�.ClsTestItem
    Dim lobjNode As Node
    Dim lstrIdea As String
    Dim lstrItem As String
    Dim lintPos As Long
    Dim i As Long
    
    On Error GoTo errHandler
    
    '��ս��档
'    subClear
    
    clstAllBase.Clear
    
    'Ϊ����ģ����������������ʼ���������Ӷ���ȡ��������������ԡ�
    With mobj����ģ��
        If .�Ƿ��Ѵ��� Then
            ctxtName.Text = .������
            ctxtNo.Text = .����
            ccmbSheet.Text = .��쵥����
            ctxtLetter.Text = .�Թ���ĸ���
            cchkAgain.Value = IIf(.�Ƿ񸴲�����, 1, 0)
            If cchkAgain.Value = 0 Then
                cchkAnnual.Value = IIf(.�Ƿ�����, 1, 0)
            Else
                cchkAnnual.Value = 0
            End If
            
            Frame2.Caption = "�޸�����" & ctxtName.Text
        Else
            Frame2.Caption = "��������" & ctxtName.Text
        End If
    End With
    
    
    'ͨ����������ҵ������ȡ������츽����Ŀ������ʼ��ʼ����츽����Ŀ�б��
    Set lobjRec = pobjҵ�����.������츽����Ŀ
    clstAllBase.Clear
    While Not lobjRec.EOF
        clstAllBase.AddItem lobjRec("������Ŀ").Value
        lobjRec.movenext
    Wend
    
    
    '��ʼ����ǰ�����ۿ�
    Set lcolInfo = mobj����ģ��.�����ۼ�
    ctrwSelectedConclusion.Nodes.Clear
    For Each lcolItem In lcolInfo
    
        '����һ���ֽڵ�.
        On Error Resume Next
        ctrwSelectedConclusion.Nodes.Add , , ctrwAllConclusion.Nodes("I" & lcolItem("������ID")).Parent.Key, ctrwAllConclusion.Nodes("I" & lcolItem("������ID")).Parent.Text
        '������Ŀ��
        ctrwSelectedConclusion.Nodes.Add ctrwAllConclusion.Nodes("I" & lcolItem("������ID")).Parent.Key, tvwChild, "I" & lcolItem("������ID"), lcolItem("����")
        On Error GoTo errHandler
    Next
        
    '��ʼ����ǰ���ģ�����츽����Ŀ�б��
    Set lcolInfo = mobj����ģ��.����������Ŀ��
    clstSelectedBase.Clear
    For Each lcolItem In lcolInfo
        clstSelectedBase.AddItem lcolItem("������Ŀ")
        clstSelectedBase.Selected(clstSelectedBase.NewIndex) = lcolItem("�Ƿ��¼")
        i = 0
        While i <= clstAllBase.ListCount - 1
            If clstAllBase.List(i) = lcolItem("������Ŀ") Then
                clstAllBase.RemoveItem i
            Else
                i = i + 1
            End If
        Wend
    Next
        
    '��ʼ����ǰ����ģ�������Ŀ�б��.
    '�޸ģ�2001-11-2�������ѡ�������Ŀ��ʾ��һ�����С�
    Set lcolInfo = mobj����ģ��.�����Ŀ��
    ctrwSelectedItem.Nodes.Clear
    For Each lobj�����Ŀ In lcolInfo
        '����һ���ֽڵ�.
        On Error Resume Next
        ctrwSelectedItem.Nodes.Add , , "I" & lobj�����Ŀ.������, ctrwAllItem.Nodes("I" & lobj�����Ŀ.������).Text
        '������Ŀ��
        ctrwSelectedItem.Nodes.Add "I" & lobj�����Ŀ.������, tvwChild, "I" & lobj�����Ŀ.����, lobj�����Ŀ.���� & " " & lobj�����Ŀ.����
        On Error GoTo errHandler
    Next
    ccmdAddItem.Enabled = False
    ccmdDeleteItem.Enabled = False
    
    '��ʼ����ǰ������ϴ������
    lstrIdea = Trim(mobj����ģ��.��ϴ������)
    If Right(lstrIdea, 1) <> "," Then lstrIdea = lstrIdea & ","
    For i = 0 To clstDisposalIdea.ListCount - 1
        If InStr(1, lstrIdea, clstDisposalIdea.List(i) & ",") > 0 Then
            clstDisposalIdea.Selected(i) = True
        Else
            clstDisposalIdea.Selected(i) = False
        End If
    Next
    
    '�����Ŀ,������Ŀ,�������б���С�
    '���������Ŀ��ѡ��,����Ӱ�ť��Ϊ������.���û��һ��Ŀ��ѡ��,��ɾ����ť��Ϊ������.
    If ctrwAllConclusion.SelectedItem Is Nothing Then
        ccmdAddConclusion.Enabled = False
    Else
        ccmdAddConclusion.Enabled = True
    End If
    
    If ctrwSelectedConclusion.SelectedItem Is Nothing Then
        ccmdDeleteConclusion.Enabled = False
    Else
        ccmdDeleteConclusion.Enabled = True
    End If
   
    If clstAllBase.ListCount > 0 Then
        clstAllBase.ListIndex = 0
        ccmdAddBase.Enabled = True
    Else
        ccmdAddBase.Enabled = False
    End If
    
    If clstSelectedBase.ListCount > 0 Then
        clstSelectedBase.ListIndex = 0
        ccmdDeleteBase.Enabled = True
    Else
        ccmdDeleteBase.Enabled = False
    End If
    
    
    '�����շѱ�׼��
    If mobj����ģ��.�շѱ�׼ = "" Then
        clstCharge.ListIndex = -1
        ctxtCharge.Text = ""
    Else
        i = gffuncItemIsInListBox(clstCharge, mobj����ģ��.�շѱ�׼)
        clstCharge.ListIndex = i
    End If
    Exit Sub
    
errHandler:
    sfsub������ "����ģ������", "frmSetMedicalExamTemplate", "subShowTemplate", Err.Number, Err.Description, True
End Sub

