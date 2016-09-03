VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTransfer 
   Caption         =   "传输数据到服务器"
   ClientHeight    =   3075
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3075
   ScaleWidth      =   4680
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton ccmdCancel 
      Caption         =   "返  回"
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton ccmdOK 
      Caption         =   "确  定"
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   2400
      Width           =   855
   End
   Begin MSComctlLib.ProgressBar cprgStatus 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComCtl2.DTPicker cdtpDate 
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   52101121
      CurrentDate     =   39914
   End
   Begin VB.TextBox ctxtServerIP 
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "体检日期："
      Height          =   180
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "服务器IP："
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   900
   End
End
Attribute VB_Name = "frmTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mobj记忆  As cls用户操作记忆

Private Sub ccmdCancel_Click()
    Unload Me
End Sub

Private Sub ccmdOk_Click()
    If ctxtServerIP = "" Then
        MsgBox "必须输入服务器的IP！", vbInformation, "系统提示"
        ctxtServerIP.SetFocus
        Exit Sub
    End If
    
    If ctxtServerIP = sffuncGetSetting("系统管理", "数据库配置", "服务器名") Then
        MsgBox "不能向自己传输数据！", vbInformation, "系统提示"
        ctxtServerIP.SetFocus
        Exit Sub
    End If

    If MsgBox("确定要将数据传输到服务器上吗？", vbQuestion + vbYesNo, "系统询问") = vbNo Then Exit Sub
    
    Dim lobjConn As New ADODB.Connection
    Dim lobjRec As Recordset, lstrSql As String
    Dim lstrDate As String, i As Integer
    
    lstrDate = Format(cdtpDate.Value, "yyyy-mm-dd")
    
    On Error GoTo errHandle
    
    cprgStatus.Visible = True
    
    lobjConn.ConnectionString = "Provider=MSDataShape.1;Data Provider=SQLOLEDB.1;Password=welcome58%*;Persist Security Info=True;User ID=jk_user;Initial Catalog=jk2006;Data Source=" & ctxtServerIP
    lobjConn.Open
    '判断服务器数据库中的服务器代号是否与本机相同
    Set lobjRec = lobjConn.Execute("select 防疫站编号,服务器代号 from 系统管理_系统基本配置表")
    If lobjRec(0) = um防疫站编号 And lobjRec(1) = um服务器代号 Then
        MsgBox "服务器与本机的本单位编号、服务器代号完全相同，不能传输数据！", vbInformation, "系统提示"
        ctxtServerIP.SetFocus
        Exit Sub
    End If
    
    '先删除服务器上来源于本电脑的该日期的体检数据
    Dim lstrNoPre As String         '本电脑的编号前缀
    
    lstrNoPre = um防疫站编号 + um服务器代号
'    lobjConn.Execute "delete 体检管理_体检访问标志表 where 系统编号 in (select 系统编号 from 体检管理_体检基本信息表 where 体检日期='" & lstrDate & "')"
'    lobjConn.Execute "delete 系统管理_系统图片管理表 where 子系统名='体检管理' and 图片编号 in (select 系统编号 from 体检管理_体检基本信息表 where 体检日期='" & lstrDate & "')"
'    lobjConn.Execute "delete 体检管理_体检基本信息表 where 体检日期='" & lstrDate & "'"
'    lobjConn.Execute "delete 体检管理_体检人员基本信息表 where 建档日期='" & lstrDate & "'"
    lobjConn.Execute "delete 体检管理_体检访问标志表 where 系统编号 in (select 系统编号 from 体检管理_体检基本信息表 where 体检日期='" & lstrDate & "' and 系统编号 like '" + lstrNoPre + "%')"
    lobjConn.Execute "delete 系统管理_系统图片管理表 where 子系统名='体检管理' and 图片编号 in (select 系统编号 from 体检管理_体检基本信息表 where 体检日期='" & lstrDate & "' and 系统编号 like '" + lstrNoPre + "%')"
    lobjConn.Execute "delete 体检管理_体检基本信息表 where 体检日期='" & lstrDate & "' and 系统编号 like '" + lstrNoPre + "%'"
    lobjConn.Execute "delete 体检管理_体检人员基本信息表 where 建档日期='" & lstrDate & "' and 健康档案编号 like '" + lstrNoPre + "%'"
    
    Set lobjRec = dafuncGetData("select * from 体检管理_体检人员基本信息表 where 建档日期='" & lstrDate & "'")
    cprgStatus.Max = lobjRec.RecordCount
    cprgStatus.Value = 0
    For i = 1 To lobjRec.RecordCount
        lobjConn.Execute "insert into 体检管理_体检人员基本信息表(健康档案编号,公民身份号码,姓名,性别,年龄,出生日期,单位申请编号,单位名称,建档日期,卫生种类,片区,行业类别)" & _
            "values('" & lobjRec("健康档案编号") & "','" & lobjRec("公民身份号码") & "','" & lobjRec("姓名") & "','" & lobjRec("性别") & "','" & lobjRec("年龄") & "','" & lobjRec("出生日期") & "','" & lobjRec("单位申请编号") & "','" & lobjRec("单位名称") & "','" & lobjRec("建档日期") & "','" & lobjRec("卫生种类") & "','" & lobjRec("片区") & "','" & lobjRec("行业类别") & "')"
        lobjRec.MoveNext
        cprgStatus.Value = i
        DoEvents
    Next
    Set lobjRec = dafuncGetData("select * from 体检管理_体检基本信息表 where 体检日期='" & lstrDate & "'")
    cprgStatus.Max = lobjRec.RecordCount
    cprgStatus.Value = 0
    For i = 1 To lobjRec.RecordCount
        lobjConn.Execute "insert into 体检管理_体检基本信息表(系统编号,健康档案编号,体检单号,试管编号,体检表名称,体检类别,体检日期,收费批号,体检状态) values('" & _
            lobjRec("系统编号") & "','" & lobjRec("健康档案编号") & "','" & lobjRec("体检单号") & "','" & lobjRec("试管编号") & "','" & lobjRec("体检表名称") & "','" & lobjRec("体检类别") & "','" & lobjRec("体检日期") & "','" & lobjRec("收费批号") & "','" & lobjRec("体检状态") & "')"
        lobjConn.Execute "insert into 体检管理_体检访问标志表(系统编号) values('" & lobjRec("系统编号") & "')"
        lobjRec.MoveNext
        cprgStatus.Value = i
        DoEvents
    Next
    Set lobjRec = dafuncGetData("select 图片编号 from 系统管理_系统图片管理表 where 子系统名='体检管理' and 图片编号 in (select 健康档案编号 from 体检管理_体检人员基本信息表 where 建档日期='" & lstrDate & "')")
    cprgStatus.Max = lobjRec.RecordCount
    cprgStatus.Value = 0
    
    Dim lobjPic As StdPicture
    
    For i = 1 To lobjRec.RecordCount
        Set lobjPic = pmfunc获取图片(lobjRec(0), "体检管理")
        pmsub保存图片 lobjConn, lobjPic, lobjRec(0), "体检管理"
        lobjRec.MoveNext
        cprgStatus.Value = i
        DoEvents
    Next
    Set lobjRec = dafuncGetData("select * from 体检管理_体检附加信息表 where 系统编号 in (select 系统编号 from 体检管理_体检基本信息表 where 体检日期='" & lstrDate & "')")
    cprgStatus.Max = lobjRec.RecordCount
    cprgStatus.Value = 0
    For i = 1 To lobjRec.RecordCount
        lobjConn.Execute "insert into 体检管理_体检附加信息表(系统编号,附加项目,项目值,项目值编号) values('" & _
            lobjRec("系统编号") & "','" & lobjRec("附加项目") & "','" & lobjRec("项目值") & "','" & lobjRec("项目值编号") & "')"
        lobjRec.MoveNext
        cprgStatus.Value = i
        DoEvents
    Next
    Set lobjRec = dafuncGetData("select * from 体检管理_体检结果信息表 where 系统编号 in (select 系统编号 from 体检管理_体检基本信息表 where 体检日期='" & lstrDate & "')")
    cprgStatus.Max = lobjRec.RecordCount
    cprgStatus.Value = 0
    For i = 1 To lobjRec.RecordCount
        lobjConn.Execute "insert into 体检管理_体检结果信息表(系统编号,体检项目,体检结果,体检医师,填写日期) values('" & _
            lobjRec("系统编号") & "','" & lobjRec("体检项目") & "','" & lobjRec("体检结果") & "'," & IIf(IsNull(lobjRec("体检医师")), "null", "'" & lobjRec("体检医师") & "'") & ",'" & lobjRec("填写日期") & "')"
        lobjRec.MoveNext
        cprgStatus.Value = i
        DoEvents
    Next
    lobjConn.Close
    MsgBox "传输成功！", vbInformation, "系统提示"
    mobj记忆.sub覆盖记忆值 "服务器IP", ctxtServerIP
    cprgStatus.Visible = False
    Exit Sub
errHandle:
    sfsub错误处理 "体检界面部件", "FrmTransfer", "ccmdOK_Click", 6666, Error, False
    lobjConn.Close
    Exit Sub
    Resume
End Sub
' 功能：    保存图片。
' 输入：    paraPicture：需要保存的图片
'           para标识号：保存图片的唯一标识号。
'           para子系统名：保存此图片的子系统名。
' 输出：    无
' 返回：    无
' 注意事项：如果该标识号在系统中已对应一张图片，将替换原有的图片。
' 作者：    罗庆
' 创建时间：2001-3-5
' 修改说明：需求发生变更，要求图片编号由各子系统产生，并保存子系统名称。
' 修改人：  罗庆
' 修改时间：2001-3-9
Public Sub pmsub保存图片(paraConn As Connection, ParaPicture As StdPicture, paraID As String, para子系统名 As String)
    On Error GoTo errHandler
    Dim lstrSql As String              'SQL语句
    Dim lrecPicture As ADODB.Recordset           '根据语句返回图片信息的RecordSet
    Dim lprbPicture As New PropertyBag '将图片信息进行序列化的属性包
    '将图片写入属性包进行序列化。
    lprbPicture.WriteProperty "Picture", ParaPicture
    '根据标识号取出相应的图片。
    lstrSql = "select * from 系统管理_系统图片管理表 where 图片编号='" & paraID & "' and 子系统名='" & para子系统名 & "'"
    Set lrecPicture = New ADODB.Recordset
    lrecPicture.Open lstrSql, paraConn, adOpenKeyset, adLockOptimistic
    'Set lrecPicture = paraConn.Execute(lstrSql)
    '如果返回空记录集，则新增一条记录。
    If lrecPicture.RecordCount = 0 Then
        lrecPicture.AddNew
    End If
    '将图片信息写入记录集中。
    lrecPicture("图片").AppendChunk lprbPicture.Contents
    lrecPicture("图片编号") = paraID
    lrecPicture("子系统名") = para子系统名
    '保存记录集更新。
    lrecPicture.Update
    lrecPicture.Close
errHandler:
    Set lrecPicture = Nothing
    Set lprbPicture = Nothing
    Set ParaPicture = Nothing
    If Err.Number = 0 Then Exit Sub
    Err.Raise Err.Number, , Err.Description
End Sub

Private Sub Form_Load()
    Set mobj记忆 = New cls用户操作记忆
    mobj记忆.用户编号 = "*"
    mobj记忆.业务名 = "体检管理"
    If mobj记忆.记忆项值("服务器IP") <> "" Then ctxtServerIP = mobj记忆.记忆项值("服务器IP")
    cdtpDate.Value = Date
End Sub
