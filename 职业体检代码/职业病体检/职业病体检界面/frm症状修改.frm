VERSION 5.00
Begin VB.Form frm症状修改 
   Caption         =   "症状修改"
   ClientHeight    =   3465
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8445
   LinkTopic       =   "Form1"
   ScaleHeight     =   3465
   ScaleWidth      =   8445
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1335
   End
   Begin VB.Frame Frame10 
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   7335
      Begin VB.TextBox Text44 
         Height          =   270
         Index           =   2
         Left            =   3480
         TabIndex        =   2
         Top             =   600
         Width           =   1695
      End
      Begin VB.ComboBox Combo28 
         Height          =   300
         Index           =   2
         ItemData        =   "frm症状修改.frx":0000
         Left            =   1920
         List            =   "frm症状修改.frx":0013
         TabIndex        =   1
         Text            =   "-"
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label37 
         Caption         =   "病程时间"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label38 
         Caption         =   "程   度"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label39 
         Caption         =   "项   目"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label76 
         Caption         =   "项    目"
         Height          =   255
         Index           =   2
         Left            =   600
         TabIndex        =   3
         Top             =   600
         Width           =   735
      End
   End
End
Attribute VB_Name = "frm症状修改"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo28_Click(Index As Integer)
If Combo28(2).Text <> "-" Then
 Text44(2).Text = "年"
 Else
 Text44(2).Text = ""
 End If
End Sub

Private Sub Command1_Click()
  
  dafuncGetData ("update 职业病体检_自觉症状表 set 程度='" & Combo28(2).Text & "',出现时间='" & Text44(2).Text & "' where  系统编号='" & frmCareerHstRegt.selectsysno & "' and 症状='" & Label76(2).Caption & "'")
  
  MsgBox ("修改成功！")
   frmCareerHstRegt.sub查询填充表格
   Command2_Click
End Sub

Private Sub Command2_Click()
 Unload frm症状修改
 
End Sub

Private Sub Form_Load()
Dim csysno As String
csysno = frmCareerHstRegt.selectsysno
Label76(2).Caption = frmCareerHstRegt.selectzz
Combo28(2).Text = frmCareerHstRegt.selectcd
Text44(2).Text = frmCareerHstRegt.selectcxrq
End Sub
