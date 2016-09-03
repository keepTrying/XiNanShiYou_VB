VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm查询 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "查询"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5100
   Icon            =   "frm查询.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CheckBox cchkType 
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   22
      Top             =   3480
      Width           =   255
   End
   Begin VB.CheckBox cchkType 
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   21
      Top             =   2520
      Width           =   255
   End
   Begin VB.CheckBox cchkType 
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   20
      Top             =   1920
      Width           =   255
   End
   Begin VB.CheckBox cchkType 
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   19
      Top             =   1320
      Width           =   255
   End
   Begin VB.CheckBox cchkType 
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   18
      Top             =   720
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CheckBox cchkType 
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   17
      Top             =   120
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CommandButton ccmdCancel 
      Caption         =   "取消(&C)"
      Height          =   375
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton ccmdOk 
      Caption         =   "确定(&O)"
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   120
      TabIndex        =   14
      Top             =   4080
      Width           =   4815
   End
   Begin VB.TextBox ctxt交费单位 
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   1920
      Width           =   3015
   End
   Begin VB.TextBox ctxt交费人 
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   1320
      Width           =   3015
   End
   Begin VB.TextBox ctxt收据号 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   720
      Width           =   3015
   End
   Begin VB.TextBox ctxt收费批号 
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.ComboBox Ccbo业务分类 
      Height          =   300
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   3480
      Width           =   3015
   End
   Begin MSComCtl2.DTPicker cdtp截止日期 
      Height          =   300
      Left            =   1680
      TabIndex        =   5
      Top             =   3000
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   529
      _Version        =   393216
      Format          =   52363265
      CurrentDate     =   36951
   End
   Begin MSComCtl2.DTPicker cdtp开始日期 
      Height          =   300
      Left            =   1680
      TabIndex        =   4
      Top             =   2520
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   529
      _Version        =   393216
      Format          =   52363265
      CurrentDate     =   36951
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "交费单位"
      Height          =   180
      Index           =   3
      Left            =   720
      TabIndex        =   13
      Top             =   1920
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "交费人"
      Height          =   180
      Index           =   2
      Left            =   720
      TabIndex        =   12
      Top             =   1320
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "止  号"
      Height          =   180
      Index           =   1
      Left            =   720
      TabIndex        =   11
      Top             =   720
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "起  号"
      Height          =   180
      Index           =   0
      Left            =   720
      TabIndex        =   10
      Top             =   120
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "业务分类"
      Height          =   180
      Left            =   720
      TabIndex        =   9
      Top             =   3480
      Width           =   720
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "至"
      Height          =   180
      Left            =   1200
      TabIndex        =   8
      Top             =   3000
      Width           =   180
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "日期范围"
      Height          =   180
      Left            =   720
      TabIndex        =   7
      Top             =   2520
      Width           =   720
   End
End
Attribute VB_Name = "frm查询"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public pstr收费批号 As String
Public pstr收据号 As String
Public pstr单位名称 As String
Public pstr交费人 As String
Public pstr开始日期 As String
Public pstr截止日期 As String
Public pstr业务分类 As String

Public pblnOk As Boolean

Private Sub Clab业务分类_Click()

End Sub

Private Sub ccmdCancel_Click()
    pblnOk = False
    Unload Me
   
End Sub

Private Sub ccmdOk_Click()
    If cchkType(0).Value = 1 Then
        pstr收费批号 = ctxt收费批号.Text
    Else
        pstr收费批号 = ""
    End If
    If cchkType(1).Value = 1 Then
        pstr收据号 = ctxt收据号.Text
    Else
        pstr收据号 = ""
    End If
    If cchkType(2).Value = 1 Then
        pstr交费人 = ctxt交费人.Text
    Else
        pstr交费人 = ""
    End If
    If cchkType(3).Value = 1 Then
        pstr单位名称 = ctxt交费单位.Text
    Else
        pstr单位名称 = ""
    End If
    If cchkType(4).Value = 1 Then
        pstr开始日期 = Format(cdtp开始日期.Value, "yyyy-mm-dd")
        pstr截止日期 = Format(cdtp截止日期.Value, "yyyy-mm-dd")
    Else
        pstr开始日期 = ""
        pstr截止日期 = ""
    End If
    If cchkType(5).Value = 1 Then
        pstr业务分类 = Ccbo业务分类.Text
    Else
        pstr业务分类 = ""
    End If
    
    pblnOk = True
    Unload Me
End Sub

Private Sub Form_Load()
    cdtp开始日期.Value = Date               '初始化开始日期输入框为本机日期
    cdtp截止日期.Value = Date               '初始化结束日期输入框为本机日期

    Dim lobjRec As Object
    On Error GoTo errhandler
    Set lobjRec = dafuncGetData("select 对应业务 from 收费管理_费用信息表 where isnull(对应业务,'')<>'' group by 对应业务  order by 对应业务 ")
        
    Ccbo业务分类.Clear
        
    Ccbo业务分类.AddItem ""
        
    Do While Not lobjRec.EOF
        Ccbo业务分类.AddItem lobjRec("对应业务").Value
        lobjRec.MoveNext
    Loop
    
    Ccbo业务分类.ListIndex = 0
    Exit Sub
errhandler:
    sfsub错误处理 "收费界面部件", "frm查询", "Form_Load", Err.Number, Err.Description, False
End Sub
