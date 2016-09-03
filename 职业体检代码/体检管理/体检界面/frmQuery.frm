VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmQuery 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "查询"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   5025
   ClipControls    =   0   'False
   Icon            =   "frmQuery.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Caption         =   "查询条件"
      Height          =   4455
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4815
      Begin VB.TextBox ctxt系统编号 
         Height          =   350
         Left            =   1680
         TabIndex        =   19
         Top             =   3840
         Width           =   2535
      End
      Begin VB.CheckBox cchkType 
         Caption         =   "系统编号(条码号)"
         Height          =   495
         Index           =   6
         Left            =   120
         TabIndex        =   18
         Top             =   3720
         Width           =   1095
      End
      Begin VB.CheckBox cchkType 
         Caption         =   "试管编号："
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   17
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox ctxt试管编号 
         Height          =   350
         Left            =   1680
         TabIndex        =   16
         Top             =   3360
         Width           =   2535
      End
      Begin VB.TextBox ctxt体检单号 
         Height          =   350
         Left            =   1680
         TabIndex        =   15
         Top             =   2880
         Width           =   2535
      End
      Begin VB.CheckBox cchkType 
         Caption         =   "体检单号："
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   14
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox ctxt姓名 
         Height          =   350
         Left            =   1680
         TabIndex        =   13
         Top             =   2300
         Width           =   2535
      End
      Begin VB.CheckBox cchkType 
         Caption         =   "姓名："
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   12
         Top             =   2400
         Width           =   975
      End
      Begin VB.CheckBox cchkType 
         Caption         =   "体检日期从："
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   1455
      End
      Begin VB.CheckBox cchkType 
         Caption         =   "体检表名称："
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   1455
      End
      Begin VB.ComboBox ccmbSheet 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1680
         TabIndex        =   6
         Top             =   1320
         Width           =   2535
      End
      Begin VB.CheckBox cchkType 
         Caption         =   "单位名称："
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox ctxtUnit 
         Enabled         =   0   'False
         Height          =   350
         Left            =   1680
         TabIndex        =   4
         Top             =   1800
         Width           =   2535
      End
      Begin VB.CommandButton ccmd定位 
         Caption         =   "..."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "单位定位"
         Top             =   1800
         Width           =   495
      End
      Begin MSComCtl2.DTPicker cdtp开始日期 
         Height          =   300
         Left            =   1680
         TabIndex        =   9
         Top             =   360
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   25362432
         CurrentDate     =   37056
      End
      Begin MSComCtl2.DTPicker cdtp结束日期 
         Height          =   300
         Left            =   1680
         TabIndex        =   10
         Top             =   840
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   25362432
         CurrentDate     =   37056
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "到："
         Height          =   180
         Index           =   1
         Left            =   1080
         TabIndex        =   11
         Top             =   960
         Width           =   360
      End
   End
   Begin VB.CommandButton ccmdCancel 
      Caption         =   "取消(&C)"
      Height          =   375
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton ccmdOk 
      Caption         =   "确定(&O)"
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4680
      Width           =   1215
   End
End
Attribute VB_Name = "frmQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'查询条件
Public pstr开始日期 As String
Public pstr截止日期 As String
Public pstr体检表名称 As String
Public pstr单位名称 As String
Public pstr姓名 As String
Public pstr体检单号 As String
Public pstr试管编号 As String
Public pstr系统编号 As String

Public pblnOk As Boolean

Private Sub cchkType_Click(Index As Integer)
    On Error Resume Next
    If cchkType(0).Value = 1 Then
        cdtp开始日期.Enabled = True
        cdtp结束日期.Enabled = True
        cdtp开始日期.SetFocus
    Else
        cdtp开始日期.Enabled = False
        cdtp结束日期.Enabled = False
    End If
    If cchkType(1).Value = 1 Then
        ccmbSheet.Enabled = True
        ccmbSheet.SetFocus
    Else
        ccmbSheet.Enabled = False
    End If
    If cchkType(2).Value = 1 Then
        ctxtUnit.Enabled = True
        ccmd定位.Enabled = True
        ctxtUnit.SetFocus
    Else
        ctxtUnit.Enabled = False
        ccmd定位.Enabled = False
    End If
    
    If cchkType(3).Value = 1 Then
        ctxt姓名.Enabled = True
        ctxt姓名.SetFocus
    Else
        ctxt姓名.Enabled = False
    End If
    
    If cchkType(4).Value = 1 Then
        ctxt体检单号.Enabled = True
        ctxt体检单号.SetFocus
    Else
        ctxt体检单号.Enabled = False
    End If
    
    If cchkType(5).Value = 1 Then
        ctxt试管编号.Enabled = True
        ctxt试管编号.SetFocus
    Else
        ctxt试管编号.Enabled = False
    End If
    If cchkType(6).Value = 1 Then
        ctxt系统编号.Enabled = True
        ctxt系统编号.SetFocus
    Else
        ctxt系统编号.Enabled = False
    End If
End Sub

Private Sub ccmdCancel_Click()
    pblnOk = False
    Unload Me

End Sub

Private Sub ccmdOk_Click()
    If cchkType(0).Value = 1 Then
        pstr开始日期 = Format(cdtp开始日期.Value, "yyyy-mm-dd")
        pstr截止日期 = Format(cdtp结束日期.Value, "yyyy-mm-dd")
    Else
        pstr开始日期 = ""
        pstr截止日期 = ""
    End If
    If cchkType(1).Value = 1 Then
        pstr体检表名称 = ccmbSheet.Text
    Else
        pstr体检表名称 = ""
    End If
    If cchkType(2).Value = 1 Then
        pstr单位名称 = ctxtUnit.Text
    Else
        pstr单位名称 = ""
    End If
    
    If cchkType(3).Value = 1 Then
        pstr姓名 = ctxt姓名.Text
    Else
        pstr姓名 = ""
    End If
    
    If cchkType(4).Value = 1 Then
        pstr体检单号 = ctxt体检单号.Text
    Else
        pstr体检单号 = ""
    End If
    
    If cchkType(5).Value = 1 Then
        pstr试管编号 = ctxt试管编号.Text
    Else
        pstr试管编号 = ""
    End If
    If cchkType(6).Value = 1 Then
        pstr系统编号 = ctxt系统编号.Text
    Else
        pstr系统编号 = ""
    End If
    pblnOk = True
    Unload Me

End Sub

Private Sub ccmd定位_Click()
    Dim lobj接口 As Object
    Dim lobjRec As Object
    
    On Error GoTo errHandler
    
    Set lobj接口 = CreateObject("单位档案业务.ClsUnitInterface")
    Set lobjRec = lobj接口.func单位简单定位(Screen.Width / 2, Screen.Height / 2)
    
    If lobjRec Is Nothing Then
        ctxtUnit.SetFocus
        Exit Sub
    End If
    
    ctxtUnit = lobjRec!单位名称
    ctxtUnit.SetFocus
    Exit Sub
errHandler:
    sfsub错误处理 "体检界面", "frmQuery", "ccmd定位_Click", Err.Number, Err.Description, False
    
End Sub

Private Sub ctxt系统编号_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And ctxt系统编号 <> "" Then
        ccmdOk_Click
    End If
End Sub

Private Sub Form_Load()
    Dim i As Long
    
    If pstr开始日期 = "" Then
        cdtp开始日期.Value = Format(DateAdd("d", -7, Date), "yyyy/mm/dd") '开始日期为当前日期的前7天
        cchkType(0).Value = 0
    Else
        cdtp开始日期.Value = Format(pstr开始日期, "yyyy/mm/dd")
        cchkType(0).Value = 1
    End If
    If pstr截止日期 = "" Then
        cdtp结束日期.Value = Date
    Else
        cdtp结束日期.Value = Format(pstr截止日期, "yyyy/mm/dd")
    End If
    
    
    '获取所有体检表名称。
    Dim lobj体检表模板集 As Object
    Dim lcolInfo As Collection
    Set lobj体检表模板集 = CreateObject("体检对象.ClsMedicalExamTemplateSet")
    Set lcolInfo = lobj体检表模板集.元素集
    
    ccmbSheet.Clear
    If lcolInfo.Count > 0 Then
        ccmbSheet.AddItem ""
    End If
    For i = 1 To lcolInfo.Count
        ccmbSheet.AddItem lcolInfo(i)
    Next i
    ccmbSheet.Text = pstr体检表名称
    If pstr体检表名称 = "" Then
        cchkType(1).Value = 0
    Else
        cchkType(1).Value = 1
    End If
    
    ctxtUnit = pstr单位名称
    If pstr单位名称 = "" Then
        cchkType(2).Value = 0
    Else
        cchkType(2).Value = 1
    End If
    
    ctxt姓名 = pstr姓名
    If pstr姓名 = "" Then
        cchkType(3).Value = 0
    Else
        cchkType(3).Value = 1
    End If
    
    ctxt体检单号 = pstr体检单号
    If pstr体检单号 = "" Then
        cchkType(4).Value = 0
    Else
        cchkType(4).Value = 1
    End If

    ctxt试管编号 = pstr试管编号
    If pstr试管编号 = "" Then
        cchkType(5).Value = 0
    Else
        cchkType(5).Value = 1
    End If

End Sub
