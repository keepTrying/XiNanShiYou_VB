VERSION 5.00
Begin VB.Form frmTip 
   Caption         =   "日积月累"
   ClientHeight    =   3285
   ClientLeft      =   2370
   ClientTop       =   2400
   ClientWidth     =   5685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   5685
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CheckBox chkLoadTipsAtStartup 
      Caption         =   "在启动时显示提示(&S)"
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   2940
      Width           =   2055
   End
   Begin VB.CommandButton cmdNextTip 
      Caption         =   "下一条提示(&N)"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   2715
      Left            =   120
      Picture         =   "frmTip.frx":0000
      ScaleHeight     =   2655
      ScaleWidth      =   3675
      TabIndex        =   1
      Top             =   120
      Width           =   3735
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "您知道吗..."
         Height          =   255
         Left            =   540
         TabIndex        =   5
         Top             =   180
         Width           =   2655
      End
      Begin VB.Label lblTipText 
         BackColor       =   &H00FFFFFF&
         Height          =   1635
         Left            =   180
         TabIndex        =   4
         Top             =   840
         Width           =   3255
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 内存中的提示数据库。
Dim Tips As New Collection

' 提示文件名称
Const TIP_FILE = "TIPOFDAY.TXT"

' 当前正在显示的提示集合的索引。
Dim CurrentTip As Long


Private Sub DoNextTip()
    On Error GoTo errHandler

    ' 随机选择一条提示。
    CurrentTip = Int((Tips.Count * Rnd) + 1)
    
    ' 或者，您可以按顺序遍历提示

'    CurrentTip = CurrentTip + 1
'    If Tips.Count < CurrentTip Then
'        CurrentTip = 1
'    End If
    
    ' 显示它。
    frmTip.DisplayCurrentTip
    
    Exit Sub
errHandler:
    sfsub错误处理 "主程序", "frmTip", "DoNextTip", True
End Sub

Function LoadTips(sFile As String) As Boolean
    On Error GoTo errHandler
    Dim NextTip As String   ' 从文件中读出的每条提示。
    Dim InFile As Integer   ' 文件的描述符。
    
    ' 包含下一个自由文件描述符。
    InFile = FreeFile
    
    ' 确定为指定文件。
    If sFile = "" Then
        LoadTips = False
        Exit Function
    End If
    
    ' 在打开前确保文件存在。
    If Dir(sFile) = "" Then
        LoadTips = False
        Exit Function
    End If
    
    ' 从文本文件中读取集合。
    Open sFile For Input As InFile
    While Not EOF(InFile)
        Line Input #InFile, NextTip
        Tips.Add NextTip
    Wend
    Close InFile

    ' 随机显示一条提示。
    DoNextTip
    
    LoadTips = True
    
    Exit Function
errHandler:
    sfsub错误处理 "主程序", "frmTip", "LoadTips", True
End Function

Private Sub chkLoadTipsAtStartup_Click()
    On Error GoTo errHandler
    ' 保存在下次启动时是否显示此窗体
    SaveSetting App.EXEName, "Options", "在启动时显示提示", chkLoadTipsAtStartup.Value
    Exit Sub
errHandler:
    sfsub错误处理 "主程序", "frmTip", "chkLoadTipsAtStartup_Click", False
End Sub

Private Sub cmdNextTip_Click()
    On Error GoTo errHandler
    DoNextTip
    Exit Sub
errHandler:
    sfsub错误处理 "主程序", "frmTip", "cmdNextTip_Click", False
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandler
    Unload Me
    Exit Sub
errHandler:
    sfsub错误处理 "主程序", "frmTip", "cmdOK_Click", False
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    Dim ShowAtStartup As Long
    
    ' 察看在启动时是否将被显示
    ShowAtStartup = GetSetting(App.EXEName, "Options", "在启动时显示提示", 1)
    If ShowAtStartup = 0 Then
        Unload Me
        Exit Sub
    End If
        
    ' 设置复选框，强行将值写回到注册表
    Me.chkLoadTipsAtStartup.Value = vbChecked
    
    ' 随机寻找
    Randomize
    
    ' 读取提示文件并且随机显示一条提示。
    If LoadTips(App.Path & "\" & TIP_FILE) = False Then
        lblTipText.Caption = "文件 " & TIP_FILE & " 没有被找到吗? " & vbCrLf & vbCrLf & _
           "创建文本文件名为 " & TIP_FILE & " 使用记事本每行写一条提示。 " & _
           "然后将它存放在应用程序所在的目录 "
    End If

    
    Exit Sub
errHandler:
    sfsub错误处理 "主程序", "frmTip", "Form_Load", False
End Sub

Public Sub DisplayCurrentTip()
    On Error GoTo errHandler
    If Tips.Count > 0 Then
        lblTipText.Caption = Tips.Item(CurrentTip)
    End If
    Exit Sub
errHandler:
    sfsub错误处理 "主程序", "frmTip", "DisplayCurrentTip", True
End Sub
