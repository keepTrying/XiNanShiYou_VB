VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frm导照片 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "导出体检健康证照片"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7440
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton ccmdExit 
      Caption         =   "退出(&X)"
      Height          =   495
      Left            =   4560
      TabIndex        =   8
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton ccmdExport 
      Caption         =   "开始导出(&S)"
      Height          =   495
      Left            =   2760
      TabIndex        =   7
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox ctxtDataDay 
      Height          =   270
      Left            =   960
      TabIndex        =   4
      Text            =   "30"
      Top             =   960
      Width           =   735
   End
   Begin VB.CheckBox cchkClear 
      Caption         =   "导出后清空库中照片"
      Height          =   255
      Left            =   3360
      TabIndex        =   3
      Top             =   960
      Width           =   1935
   End
   Begin VB.CommandButton CommPotoPath 
      Caption         =   "....."
      Height          =   375
      Left            =   6720
      TabIndex        =   2
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox CtxtPotoPath 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Width           =   4935
   End
   Begin ComctlLib.StatusBar Cstau状态栏 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   9
      Top             =   2925
      Width           =   7830
      _ExtentX        =   13811
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   7409
            MinWidth        =   7409
            Text            =   ""
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "导出："
      Height          =   180
      Left            =   360
      TabIndex        =   6
      Top             =   960
      Width           =   540
   End
   Begin VB.Label Label3 
      Caption         =   "天之前的照片，"
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "照片存放路径："
      Height          =   240
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1320
   End
End
Attribute VB_Name = "frm导照片"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ccmdExit_Click()
    Unload Me
End Sub

Private Sub ccmdExport_Click()
    Dim lobj数据服务 As Object
    On Error GoTo errHandler
    MousePointer = 11
    ccmdExit.Enabled = False
    ccmdExport.Enabled = False
    Cstau状态栏.Panels.Item(1).Text = "正在导出体检健康证图片..."
    Plng数据期限 = Val(ctxtDataDay)
    PStr照片目录 = CtxtPotoPath
    
    Set lobj数据服务 = New Cls数据服务
    lobj数据服务.sub服务进程

    Cstau状态栏.Panels.Item(1).Text = "照片导出完毕！"
    MousePointer = 0
    ccmdExit.Enabled = True
    ccmdExport.Enabled = True
    Exit Sub
errHandler:
    MsgBox Error, vbOKOnly + vbExclamation, "系统提示"
End Sub

Private Sub CommPotoPath_Click()
On Error Resume Next

    Dim bi As BROWSEINFO 'declare the needed variables
    Dim rtn&, pidl&, path$, pos%
    Dim t%, Specin$, SpecOut$
    
    bi.hOwner = Me.hwnd 'centres the dialog on the screen
    bi.lpszTitle = "浏览目标文件..." 'set the title text
    bi.ulFlags = BIF_RETURNONLYFSDIRS 'the type of folder(s) to return
    pidl& = SHBrowseForFolder(bi) 'show the dialog box
      
    path = Space(512) 'sets the maximum characters
    t% = SHGetPathFromIDList(ByVal pidl&, ByVal path) 'gets the selected path
    
    pos% = InStr(path$, Chr$(0)) 'extracts the path from the string
    Specin = Left(path$, pos - 1) 'sets the extracted path to SpecIn
    
    If Right$(Specin, 1) = "\" Then 'makes sure that "\" is at the end of the path
       SpecOut = Specin             'if so then, do nothing
    Else                            'otherwise
       SpecOut = Specin
    End If


    If SpecOut <> "" Then
        CtxtPotoPath.Text = SpecOut
        PStr照片目录 = CtxtPotoPath.Text
    End If
End Sub
