VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmfresult 
   Caption         =   "血常规结果导入"
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9825
   LinkTopic       =   "Form1"
   ScaleHeight     =   4500
   ScaleWidth      =   9825
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame frmuploadFresult 
      Caption         =   "血常规结果录入"
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   8175
      Begin VB.CommandButton Comm选择 
         Caption         =   "选择文件"
         Height          =   375
         Left            =   4320
         TabIndex        =   9
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   960
         Width           =   3975
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   1560
         TabIndex        =   5
         Top             =   2160
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton Command1 
         Caption         =   "开始导入"
         Height          =   375
         Left            =   5520
         TabIndex        =   1
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label医师编号 
         Caption         =   "医师编号"
         Height          =   255
         Left            =   1320
         TabIndex        =   8
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "体检医师："
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   480
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Label3 
         Caption         =   "上传进度："
         Height          =   255
         Left            =   720
         TabIndex        =   4
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Labelstate 
         Caption         =   "还未进行上传操作。"
         Height          =   255
         Left            =   1560
         TabIndex        =   3
         Top             =   1680
         Width           =   3855
      End
      Begin VB.Label Label1 
         Caption         =   "导入文件状态："
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   1680
         Width           =   1335
      End
   End
   Begin MSComDlg.CommonDialog ccdg 
      Left            =   8520
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmfresult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub Check1_Click()
'If Check1.Value = 0 Then
'' Label4.Visible = True
' Text1.Visible = True
' Comm选择.Visible = True
' Else
'' Label4.Visible = False
' Text1.Visible = False
' Comm选择.Visible = False
' End If
'End Sub

Private Sub Command1_Click()
 Dim lstr单项结论 As String
 
Dim name As String
name = Text1.Text

Dim oname As String
Dim nname As String
oname = "d:\fresult\" + name + ".txt"
nname = "d:\fresult\" + name + ".xls"
'重命名将txt改成xls
'Name oname As nname
'用Dir函数来判断文件是否存在
If Dir(nname) = "" Then
    If Dir(oname) = "" Then  '判断text文件是否存在   2015-9-17
    MsgBox "文件不存在或者文件选择错误", , "信息提示"
    Exit Sub
    Else
    Name oname As nname
    End If
End If


Dim xlsApp As Excel.Application     '声明对象变量
Set xlsApp = New Excel.Application        '实例化对象
xlsApp.Visible = False      '使Excel隐藏不可见
xlsApp.Workbooks.Open (nname)   '打开EXCEL文件

'Dim mg As Range
Labelstate.Caption = "正在检查数据结构。"
Dim totals As Long
'Sheets1.name = Replace(ThisWorkbook.name, ".xls", "")
totals = xlsApp.ActiveWorkbook.Sheets(name).UsedRange.rows.Count
Dim prmax As Long
prmax = 0
Dim ii As Long
 For ii = 1 To totals
    If xlsApp.ActiveWorkbook.Sheets(name).Cells(ii, 1) <> "" And xlsApp.ActiveWorkbook.Sheets(name).Cells(ii + 25, 2) = "大型血小板比率|P-LCR" Then
    prmax = prmax + 1
    End If
    Labelstate.Caption = "正在检查数据结构,进度" + Str(ii) + "/" + Str(totals)
Next ii

'初始化进度条
ProgressBar1.Min = 0
ProgressBar1.Max = prmax
 Labelstate.Caption = "开始上传..."
 
 Dim 体检项目编号(2) As String
Dim i As Long
 For i = 1 To totals
    If xlsApp.ActiveWorkbook.Sheets(name).Cells(i, 1) <> "" And xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 25, 2) = "大型血小板比率|P-LCR" Then
'    If xlsApp.ActiveWorkbook.Sheets(name).Cells(i, 1) <> "" And xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 25, 2) = "血小板分布宽度|PDW" Then
    Dim SysNo As String
    '提示：保证用户的原始数据一定是文本文档并保证系统编号与我们一样或者是系统编号是十六进制的excel文件 2015-9-22 by 牟俊
    
'    SysNo = "00" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i, 6).Value  '原来的txt文档转换成的表格，系统编号在第六列
'    SysNo = "00" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i, 12).Value    '现在的txt文档转换成的表格，系统编号在第12列     2016-1-6 by 牟俊
    SysNo = "0" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i, 12).Value    '现在的txt文档转换成的表格，系统编号在第12列并且系统编号前面少了个0（2016-1-4后面导入的才少）   2016-1-6 by 牟俊
    
'   lstr单项结论 = pobj业务对象.func获取单项结论(cgrdInput.TextMatrix(Row, 0), cgrdInput.TextMatrix(Row, 2))
    dafuncGetData ("update 职业病体检_结果信息_血常规化验科 set 体检结果='" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 2, 3) & "', 体检医师='" & Label医师编号.Caption & "',  填写时间='" & Now & "',单项结论='" & pobj业务对象.func获取单项结论("04021", xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 2, 3)) & "' where 系统编号='" & SysNo & "' and 体检项目='04021' ")
    dafuncGetData ("update 职业病体检_结果信息_血常规化验科 set 体检结果='" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 3, 3) & "', 体检医师='" & Label医师编号.Caption & "',  填写时间='" & Now & "',单项结论='" & pobj业务对象.func获取单项结论("04022", xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 3, 3)) & "' where 系统编号='" & SysNo & "' and 体检项目='04022' ")
    dafuncGetData ("update 职业病体检_结果信息_血常规化验科 set 体检结果='" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 4, 3) & "', 体检医师='" & Label医师编号.Caption & "',  填写时间='" & Now & "',单项结论='" & pobj业务对象.func获取单项结论("04023", xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 4, 3)) & "' where 系统编号='" & SysNo & "' and 体检项目='04023' ")
    dafuncGetData ("update 职业病体检_结果信息_血常规化验科 set 体检结果='" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 5, 3) & "', 体检医师='" & Label医师编号.Caption & "',  填写时间='" & Now & "',单项结论='" & pobj业务对象.func获取单项结论("04001", xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 5, 3)) & "' where 系统编号='" & SysNo & "' and 体检项目='04001' ")
    dafuncGetData ("update 职业病体检_结果信息_血常规化验科 set 体检结果='" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 6, 3) & "', 体检医师='" & Label医师编号.Caption & "',  填写时间='" & Now & "',单项结论='" & pobj业务对象.func获取单项结论("04024", xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 6, 3)) & "' where 系统编号='" & SysNo & "' and 体检项目='04024' ")
    
    dafuncGetData ("update 职业病体检_结果信息_血常规化验科 set 体检结果='" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 7, 3) & "', 体检医师='" & Label医师编号.Caption & "',  填写时间='" & Now & "',单项结论='" & pobj业务对象.func获取单项结论("04002", xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 7, 3)) & "' where 系统编号='" & SysNo & "' and 体检项目='04002' ")
    dafuncGetData ("update 职业病体检_结果信息_血常规化验科 set 体检结果='" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 8, 3) & "', 体检医师='" & Label医师编号.Caption & "',  填写时间='" & Now & "',单项结论='" & pobj业务对象.func获取单项结论("04003", xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 8, 3)) & "' where 系统编号='" & SysNo & "' and 体检项目='04003' ")
    dafuncGetData ("update 职业病体检_结果信息_血常规化验科 set 体检结果='" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 9, 3) & "', 体检医师='" & Label医师编号.Caption & "',  填写时间='" & Now & "',单项结论='" & pobj业务对象.func获取单项结论("04004", xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 9, 3)) & "' where 系统编号='" & SysNo & "' and 体检项目='04004' ")
    dafuncGetData ("update 职业病体检_结果信息_血常规化验科 set 体检结果='" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 10, 3) & "', 体检医师='" & Label医师编号.Caption & "',  填写时间='" & Now & "',单项结论='" & pobj业务对象.func获取单项结论("04005", xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 10, 3)) & "' where 系统编号='" & SysNo & "' and 体检项目='04005' ")
    dafuncGetData ("update 职业病体检_结果信息_血常规化验科 set 体检结果='" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 11, 3) & "', 体检医师='" & Label医师编号.Caption & "',  填写时间='" & Now & "',单项结论='" & pobj业务对象.func获取单项结论("04006", xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 11, 3)) & "' where 系统编号='" & SysNo & "' and 体检项目='04006' ")
    
    dafuncGetData ("update 职业病体检_结果信息_血常规化验科 set 体检结果='" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 12, 3) & "', 体检医师='" & Label医师编号.Caption & "',  填写时间='" & Now & "',单项结论='" & pobj业务对象.func获取单项结论("04007", xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 12, 3)) & "' where 系统编号='" & SysNo & "' and 体检项目='04007' ")
    dafuncGetData ("update 职业病体检_结果信息_血常规化验科 set 体检结果='" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 13, 3) & "', 体检医师='" & Label医师编号.Caption & "',  填写时间='" & Now & "',单项结论='" & pobj业务对象.func获取单项结论("04008", xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 13, 3)) & "' where 系统编号='" & SysNo & "' and 体检项目='04008' ")
    dafuncGetData ("update 职业病体检_结果信息_血常规化验科 set 体检结果='" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 14, 3) & "', 体检医师='" & Label医师编号.Caption & "',  填写时间='" & Now & "',单项结论='" & pobj业务对象.func获取单项结论("04009", xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 14, 3)) & "' where 系统编号='" & SysNo & "' and 体检项目='04009' ")
    dafuncGetData ("update 职业病体检_结果信息_血常规化验科 set 体检结果='" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 15, 3) & "', 体检医师='" & Label医师编号.Caption & "',  填写时间='" & Now & "',单项结论='" & pobj业务对象.func获取单项结论("04010", xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 15, 3)) & "' where 系统编号='" & SysNo & "' and 体检项目='04010' ")
    dafuncGetData ("update 职业病体检_结果信息_血常规化验科 set 体检结果='" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 16, 3) & "', 体检医师='" & Label医师编号.Caption & "',  填写时间='" & Now & "',单项结论='" & pobj业务对象.func获取单项结论("04011", xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 16, 3)) & "' where 系统编号='" & SysNo & "' and 体检项目='04011' ")
    
    dafuncGetData ("update 职业病体检_结果信息_血常规化验科 set 体检结果='" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 17, 3) & "', 体检医师='" & Label医师编号.Caption & "',  填写时间='" & Now & "',单项结论='" & pobj业务对象.func获取单项结论("04012", xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 17, 3)) & "' where 系统编号='" & SysNo & "' and 体检项目='04012' ")
    dafuncGetData ("update 职业病体检_结果信息_血常规化验科 set 体检结果='" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 18, 3) & "', 体检医师='" & Label医师编号.Caption & "',  填写时间='" & Now & "',单项结论='" & pobj业务对象.func获取单项结论("04013", xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 18, 3)) & "' where 系统编号='" & SysNo & "' and 体检项目='04013' ")
    dafuncGetData ("update 职业病体检_结果信息_血常规化验科 set 体检结果='" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 19, 3) & "', 体检医师='" & Label医师编号.Caption & "',  填写时间='" & Now & "',单项结论='" & pobj业务对象.func获取单项结论("04014", xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 19, 3)) & "' where 系统编号='" & SysNo & "' and 体检项目='04014' ")
    dafuncGetData ("update 职业病体检_结果信息_血常规化验科 set 体检结果='" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 20, 3) & "', 体检医师='" & Label医师编号.Caption & "',  填写时间='" & Now & "',单项结论='" & pobj业务对象.func获取单项结论("04015", xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 20, 3)) & "' where 系统编号='" & SysNo & "' and 体检项目='04015' ")
    dafuncGetData ("update 职业病体检_结果信息_血常规化验科 set 体检结果='" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 21, 3) & "', 体检医师='" & Label医师编号.Caption & "',  填写时间='" & Now & "',单项结论='" & pobj业务对象.func获取单项结论("04016", xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 21, 3)) & "' where 系统编号='" & SysNo & "' and 体检项目='04016' ")
    
    dafuncGetData ("update 职业病体检_结果信息_血常规化验科 set 体检结果='" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 22, 3) & "', 体检医师='" & Label医师编号.Caption & "',  填写时间='" & Now & "',单项结论='" & pobj业务对象.func获取单项结论("04017", xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 22, 3)) & "' where 系统编号='" & SysNo & "' and 体检项目='04017' ")
    dafuncGetData ("update 职业病体检_结果信息_血常规化验科 set 体检结果='" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 23, 3) & "', 体检医师='" & Label医师编号.Caption & "',  填写时间='" & Now & "',单项结论='" & pobj业务对象.func获取单项结论("04018", xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 23, 3)) & "' where 系统编号='" & SysNo & "' and 体检项目='04018' ")
    dafuncGetData ("update 职业病体检_结果信息_血常规化验科 set 体检结果='" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 24, 3) & "', 体检医师='" & Label医师编号.Caption & "',  填写时间='" & Now & "',单项结论='" & pobj业务对象.func获取单项结论("04019", xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 24, 3)) & "' where 系统编号='" & SysNo & "' and 体检项目='04019' ")
    dafuncGetData ("update 职业病体检_结果信息_血常规化验科 set 体检结果='" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 25, 3) & "', 体检医师='" & Label医师编号.Caption & "',  填写时间='" & Now & "',单项结论='" & pobj业务对象.func获取单项结论("04020", xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 25, 3)) & "' where 系统编号='" & SysNo & "' and 体检项目='04020' ")
    
    '血常规基本信息导入新加的职业病体检_结果信息_血常规基本信息表 2016-1-13 by 牟俊
    dafuncGetData ("delete from 职业病体检_结果信息_血常规基本信息表 where 系统编号='" & SysNo & "'")   '删除原有记录，以最后导入的为准
    dafuncGetData ("insert into 职业病体检_结果信息_血常规基本信息表(系统编号,姓名,性别,年龄,病人类型,科室,标本号,标本类型,送检医生,检验者,检验日期) values('" & SysNo & "','" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i, 5) & "','" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i, 8) & "','" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i, 9) & "','" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i, 7) & "','" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i, 11) & "','" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i, 3) & "','" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i, 17) & "','" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i, 13) & "','" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i, 15) & "','" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i, 14) & "')")
    
    ProgressBar1.Value = ProgressBar1.Value + 1
   End If
Next i
 Labelstate.Caption = "上传完成!"
  xlsApp.ActiveWorkbook.Close savechanges:=False    '关闭活动工作簿同时不保存对它的更改。
  xlsApp.Quit                                      '关闭EXCELL
  Set xlsApp = Nothing    '释放资源
  MsgBox ("导入成功！"), vbInformation, "系统提示"
  Unload Me
End Sub
'选择文件   为了可以自主选择要导入的文件增加的 2015-9-17 by 牟俊
Private Sub Comm选择_Click()
    Dim i As Integer
    Dim lstrTmp As String
'    CommonDialog1.ShowOpen
ccdg.Filter = "All Files (*.*)|*.*|Excel file" & _
            "(*.xls)|*.xls|Batch Files (*.bat)|*.bat"
    ccdg.ShowOpen
'    ccdg.FileName = ""
    Text1.Text = CreateObject("Scripting.FileSystemObject").GetBaseName(ccdg.FileName)  '只要文件名，不要路径和后缀名 2015-9-18
End Sub


