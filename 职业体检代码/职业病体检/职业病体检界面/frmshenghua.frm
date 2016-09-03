VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmshenghua 
   Caption         =   "生化结果录入"
   ClientHeight    =   5340
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8940
   LinkTopic       =   "Form1"
   ScaleHeight     =   5340
   ScaleWidth      =   8940
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox TextPassword 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   15
      Top             =   3600
      Width           =   2655
   End
   Begin VB.TextBox Textname 
      Height          =   375
      Left            =   1560
      TabIndex        =   14
      Top             =   3000
      Width           =   2655
   End
   Begin VB.CheckBox Check是否 
      Caption         =   "是否更改服务器"
      Height          =   255
      Left            =   6960
      TabIndex        =   11
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton command连接 
      Caption         =   "连接"
      Height          =   495
      Left            =   7560
      TabIndex        =   8
      Top             =   4800
      Width           =   1215
   End
   Begin VB.TextBox TextBase 
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   2400
      Width           =   2655
   End
   Begin VB.TextBox TextIP 
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   1800
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "退出"
      Height          =   615
      Left            =   5280
      TabIndex        =   5
      Top             =   2520
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "读取"
      Height          =   615
      Left            =   5280
      TabIndex        =   0
      Top             =   960
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker DTP开始 
      Height          =   420
      Left            =   1080
      TabIndex        =   3
      Top             =   120
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   741
      _Version        =   393216
      CustomFormat    =   "yyyy年MM月dd日"
      Format          =   59965443
      CurrentDate     =   42370
   End
   Begin MSComCtl2.DTPicker DTP截止 
      Height          =   420
      Left            =   1080
      TabIndex        =   4
      Top             =   840
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   741
      _Version        =   393216
      CustomFormat    =   "yyyy年MM月dd日"
      Format          =   59965443
      CurrentDate     =   42370
   End
   Begin VB.Label Label6 
      Caption         =   "密码"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "登录名"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "目标数据库名"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "目标服务器IP"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "结束时间"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "起始时间"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmshenghua"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'链接数据库的字符串
'Private Const Conn As String = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;PWD=123456;Initial Catalog=KHB2LIS;Data Source=192.168.0.155"
'hr是数据库名称  Catalong=数据库名称
Private sqlreuslt As String
Private res As Object
Private IsConnect As Boolean             '判读数据库是否链接
Private cnn As ADODB.Connection       ' 链接数据库的connection对象
Private rs As ADODB.Recordset           '保存结果集的recordset对象
 
Private Sub Check是否_Click()
If Check是否.Value = 1 Then
TextIP.Enabled = True
TextBase.Enabled = True
Textname.Enabled = True
TextPassword.Enabled = True
Else
TextIP.Enabled = False
TextBase.Enabled = False
Textname.Enabled = False
TextPassword.Enabled = False
End If
End Sub

'查询sql2005数据表里的数据并放入sql2000生化数据表里  2016-1-21 by 牟俊
Private Sub Command1_Click()
    Dim xmID As String
    Dim lobject As Object
    Dim ID As String
    Dim lob As Object
     Dim dtpTimeTo As Date
    Dim sql As String
    '根据结果生化表来重新调整导入的信息（导入到重新建立的表中）  2016-2-1 by 牟俊↓
    Dim SysNo As String
    command连接_Click
    '判断是否连接成功，如果不成功直接退出
    If cnn.State <> adStateOpen Then
    Exit Sub
    End If
    sql = " (TestTime between '" & Format(DTP开始.Value, "yyyy-mm-dd" & " 00:00:00") & "' and '" & Format(DTP截止.Value, "yyyy-mm-dd" & " 23:59:59") & "')"
    Set lob = cnn.Execute("select BarCode from labLISItemResult  where Barcode<>'' and  " & sql & " group by BarCode")
    If lob.RecordCount > 0 Then
    Dim k As Integer
    For k = 1 To lob.RecordCount
    SysNo = lob("BarCode")
    Set res = cnn.Execute("select * from labLISItemResult where Barcode='" & SysNo & "' and " & sql & "")
         dafuncGetData ("delete from 职业病体检_结果信息_生化科导入信息表 where BarCode='" & SysNo & "'and " & sql & "") '删除原有记录，以最后导入的为准
        If res.RecordCount > 0 Then
        Dim i As Long
        For i = 1 To res.RecordCount
        dafuncGetData ("insert into 职业病体检_结果信息_生化科导入信息表 values('" & res("SeqNo") & "','" & res("SampleID") & "','" & res("BarCode") & "','" & res("ItemCode") & "','" & res("TestValue") & "','" & res("TestTime") & "','" & res("SampleNo") & "','" & res("Maked") & "')")
        res.MoveNext
        Next i
        End If
    lob.MoveNext
    Next k
    MsgBox "数据已读取完成"
    Unload Me
    Else
    MsgBox "数据不存在，请确定是否发送过LIS"
    End If
    '2016-2-1 by 牟俊 ↑
End Sub


Private Sub Command2_Click()
Unload Me
End Sub


Private Sub command连接_Click()
On Error GoTo erhander
Set cnn = New Connection
Dim strConnection As String
Dim sever As String
Dim base As String
Dim pawd As String
Dim username As String
sever = TextIP.Text
base = TextBase.Text
pawd = TextPassword.Text
username = Textname.Text
If cnn.State = adStateOpen Then cnn.Close
'strConnection = "Provider=SQLOLEDB.1;Password=123456;Persist Security Info=True;User ID=sa;Initial Catalog=" & base & ";Data Source=" & sever & ""
strConnection = "Provider=SQLOLEDB.1;Password=" & pawd & ";Persist Security Info=True;User ID=" & username & ";Initial Catalog=" & base & ";Data Source=" & sever & ""
cnn.ConnectionString = strConnection
cnn.CursorLocation = adUseClient
cnn.Open strConnection
Exit Sub
erhander:
Select Case Err.Number
    Case -2147467259
    MsgBox "找不到指定sql数据库", vbExclamation, "数据库错误"
    Case -2147217843
    MsgBox "指定的sql server数据库用户不存在或口令错误", vbExclamation, "数据库错误"
    Case Else
    MsgBox "数据环境连接失败，请找系统管理员进行检查", vbExclamation, "数据库错误"
End Select
End Sub

Private Sub Form_Load()
TextIP.Enabled = False
TextBase.Enabled = False
Textname.Enabled = False
TextPassword.Enabled = False
command连接.Visible = False
'默认一个服务器IP和数据库  2016-1-29 by 牟俊
TextIP.Text = "192.168.0.164"
'TextBase.Text = "LabConsole"
TextBase.Text = "KHB2LIS"
Textname.Text = "sa"
TextPassword.Text = "123456"
End Sub


