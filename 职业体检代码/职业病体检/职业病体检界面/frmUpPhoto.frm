VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmUpPhoto 
   Caption         =   "Form1"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12465
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   12465
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text系统编号 
      Height          =   375
      Left            =   1560
      TabIndex        =   12
      Top             =   120
      Width           =   3135
   End
   Begin MSComDlg.CommonDialog ccdg 
      Left            =   11880
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   975
      Left            =   2880
      TabIndex        =   7
      Top             =   8280
      Width           =   6015
      Begin VB.CommandButton Cmd退出 
         Caption         =   "退出"
         Height          =   375
         Left            =   3720
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Cmd显示 
         Caption         =   "显示"
         Height          =   375
         Left            =   1920
         TabIndex        =   9
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton Cmd上传 
         Caption         =   "上传"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   5655
      Left            =   360
      ScaleHeight     =   5595
      ScaleWidth      =   11115
      TabIndex        =   6
      Top             =   2280
      Width           =   11175
   End
   Begin VB.Frame Frame1 
      Caption         =   "上传图片"
      Height          =   1575
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   11295
      Begin VB.CommandButton Cmd浏览 
         Caption         =   "浏览"
         Height          =   375
         Left            =   9240
         TabIndex        =   5
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox Text图片名称 
         Height          =   375
         Left            =   1320
         TabIndex        =   4
         Text            =   "图片名称"
         Top             =   1080
         Width           =   7575
      End
      Begin VB.TextBox Text路径 
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         Text            =   "图片路径"
         Top             =   360
         Width           =   7575
      End
      Begin VB.Label Label2 
         Caption         =   "图片名"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "图片上传"
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Label Label3 
      Caption         =   "系统编号："
      Height          =   255
      Left            =   600
      TabIndex        =   11
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmUpPhoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd浏览_Click()
'ccdg.Filter = "All Files (*.*)|*.*|Excel file" & _
'            "(*.xls)|*.xls|Batch Files (*.bat)|*.bat"
ccdg.Filter = "all files(*.*)"
    ccdg.ShowOpen
    Picture1.Picture = LoadPicture(ccdg.FileName)
    Text路径.Text = ccdg.FileName
    Text图片名称.Text = CreateObject("Scripting.FileSystemObject").GetBaseName(ccdg.FileName) '只要文件名，不要路径和后缀名
End Sub

Private Sub Cmd上传_Click()
'    Dim rs As Object
'    Dim lngFile As Long
'    Dim lngCh As Long
'    Dim rgbArray() As Byte
'    Dim str As String
'    lngFile = FileLen(Text路径.Text)
'    ReDim rgbArray(lngFile)
'    lngCh = FreeFile
'    Open Text路径.Text For Binary As #lngCh
'    Get #lngCh, , rgbArray()
'    Close #lngCh

    dafuncGetData ("insert into 职业病体检_复查资料信息表 values('" & Text系统编号.Text & "','1', '" & Picture1.Picture & "')")

'   Call rs("select * from 职业病体检_复查资料信息表")
'    rs.AddNew
'    rs.Fields(1).Value = Text系统编号.Text
'    rs.Fields(2).Value = 1
'    rs.Fields(2).Value = rgbArray
'    rs.Update
    Text路径.Text = Empty
    Text图片名称.Text = Empty
    Picture1.Picture = Nothing
    Exit Sub
End Sub

Private Sub Cmd退出_Click()
    Unload Me
End Sub

Private Sub Cmd显示_Click()
'Set lobjRec = CreateObject("职业病对象.clspersonexamed")
    Set lobjRec = CreateObject("职业病界面.clsShowPhoto")
    lobjRec.系统编号 = Trim(Text系统编号.Text)
    Picture1.Picture = lobjRec.像片
    Picture1.Visible = True
End Sub

Private Sub Form_Load()
    Text路径.Text = ""
    Text图片名称.Text = ""
    Text系统编号.Text = FrmRegisterAgain.clblsysno.Text & "F"
'    Text系统编号.Text = FrmRegisterAgain.clblsysno.Text
End Sub
