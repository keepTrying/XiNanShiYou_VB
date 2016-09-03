VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls体检"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3CA962C70141"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


Private mstr系统编号 As String

Public 体检号 As String

Public 姓名 As String
Public 性别 As String
Public 年龄 As String
Public 种类 As String
Public 职业 As String
Public 民族 As String
Public 行业类别 As String

Public 申请编号 As String
Public 单位名称 As String

Public 体检日期 As String
Public 体检结论 As String
Public 检出病种 As String
Public 培训结论 As String
Public 培训日期 As String
Public 处置 As String

Public 健康证号 As String
Public 发证日期 As String
Public 有效期至 As String
Public 发证单位 As String
Public 身份证号 As String   '食品卫生证的证号，按年重编

Public 调离编号 As String
Public 调离期限 As Integer
Public 调离日期 As String
Public 备注 As String

Public 状态 As String

Public 体检系统编号 As String
Public 二代身份证编号 As String
Public 照片 As StdPicture

Public Property Get 系统编号() As String
    系统编号 = mstr系统编号
End Property

Public Property Let 系统编号(ByVal vNewValue As String)
    Dim lobjRec As Object
    
    On Error GoTo errhandler
    mstr系统编号 = vNewValue
    Me.体检号 = ""
    Me.姓名 = ""
    Me.性别 = ""
    Me.年龄 = 0
    Me.种类 = ""
    Me.申请编号 = ""
    Me.单位名称 = ""
    Me.检出病种 = ""
    Me.处置 = ""
    Me.备注 = ""
    Me.民族 = ""
    Me.职业 = ""
    Me.行业类别 = ""
    Me.身份证号 = ""
    
    Me.健康证号 = ""
    Me.发证日期 = ""
    Me.有效期至 = ""
    Me.发证单位 = ""
    体检系统编号 = ""
    
    Me.调离编号 = ""
    Me.调离期限 = 0
    Me.调离日期 = ""
    
    '获取体检信息。
    If vNewValue <> "" Then
        Set lobjRec = dafuncGetData("select  * from 健康证管理_办证申请信息表 where 系统编号='" & mstr系统编号 & "'")
        If lobjRec.RecordCount > 0 Then
            Me.体检号 = IIf(IsNull(lobjRec!体检号), "", lobjRec!体检号)
            Me.姓名 = lobjRec!姓名
            Me.性别 = lobjRec!性别
            Me.年龄 = IIf(IsNull(lobjRec!年龄), "", lobjRec!年龄)
            Me.种类 = IIf(IsNull(lobjRec!种类), "", lobjRec!种类)
            Me.职业 = IIf(IsNull(lobjRec!职业), "", lobjRec!职业)
            Me.民族 = IIf(IsNull(lobjRec!民族), "", lobjRec!民族)
            
            Me.申请编号 = IIf(IsNull(lobjRec!申请编号), "", lobjRec!申请编号)
            Me.单位名称 = IIf(IsNull(lobjRec!单位名称), "", lobjRec!单位名称)
            Me.体检日期 = IIf(IsNull(lobjRec!体检日期), "", lobjRec!体检日期)
            Me.体检结论 = IIf(IsNull(lobjRec!体检结论), "合格", lobjRec!体检结论)
            Me.培训结论 = IIf(IsNull(lobjRec!培训结论), "合格", lobjRec!培训结论)
            Me.培训日期 = IIf(IsNull(lobjRec!培训日期), "", lobjRec!培训日期)
            Me.检出病种 = IIf(IsNull(lobjRec!检出病种), "合格", lobjRec!检出病种)
            Me.处置 = IIf(IsNull(lobjRec!处置), "合格", lobjRec!处置)
            Me.备注 = IIf(IsNull(lobjRec!备注), "合格", lobjRec!备注)
            Me.身份证号 = IIf(IsNull(lobjRec!身份证号), "", lobjRec!身份证号)
            
            Me.健康证号 = IIf(IsNull(lobjRec!健康证号), "", lobjRec!健康证号)
            Me.发证日期 = IIf(IsNull(lobjRec!发证日期), "", lobjRec!发证日期)
            Me.有效期至 = IIf(IsNull(lobjRec!有效期至), "", lobjRec!有效期至)
            Me.发证单位 = IIf(IsNull(lobjRec!发证单位), "", lobjRec!发证单位)
            
            Me.调离编号 = IIf(IsNull(lobjRec!调离编号), "", lobjRec!调离编号)
            Me.调离期限 = IIf(IsNull(lobjRec!调离期限), 0, lobjRec!调离期限)
            Me.调离日期 = IIf(IsNull(lobjRec!调离日期), "", lobjRec!调离日期)
            体检系统编号 = IIf(IsNull(lobjRec!体检系统编号), "", lobjRec!体检系统编号)
            
            行业类别 = IIf(IsNull(lobjRec!行业类别), "", lobjRec!行业类别)
            Me.二代身份证编号 = IIf(IsNull(lobjRec!二代身份证编号), "", lobjRec!二代身份证编号)
            Me.状态 = lobjRec!状态
            
            On Error Resume Next
'            Set 照片 = LoadPicture(App.Path & "\Photo\" & mstr系统编号 & ".bmp")
            If 体检系统编号 <> "" Then
                Set lobjRec = dafuncGetData("select 健康档案编号 from 体检管理_体检基本信息表 where 系统编号='" & 体检系统编号 & "'")
                If lobjRec.RecordCount > 0 Then
                    Set 照片 = pmfunc获取图片(lobjRec(0), "体检管理")
                End If
            Else
                Set 照片 = pmfunc获取图片(mstr系统编号, "健康证管理")
            End If
        End If
    End If
    
    Exit Property
errhandler:
    sfsub错误处理 "健康证管理部件", "cls体检", "Property Let 系统编号", Err.Number, Err.Description, True
End Property

Public Sub sub保存()
    On Error GoTo errhandler
    
    dasubBeginTran
    If mstr系统编号 = "" Then
        '生成系统编号。
        mstr系统编号 = pobj体检管理.func生成系统编号(Me.种类)
        
        状态 = "未打印"
        
        '插入记录。
        dafuncGetData "insert into 健康证管理_办证申请信息表(系统编号,体检号,姓名,性别,年龄,种类,职业,民族,申请编号,单位名称,体检日期,体检结论,培训结论,检出病种,处置,状态,培训日期,行业类别,身份证号)" _
                    & " values('" & mstr系统编号 & "','" & 体检号 & "','" & Me.姓名 & "','" & 性别 & "','" & 年龄 & "','" & 种类 & "','" & 职业 & "','" & 民族 & "','" & 申请编号 & "','" & 单位名称 & "','" & 体检日期 & "','" & 体检结论 & "','" & 培训结论 & "','" & 检出病种 & "','" & 处置 & "','" & 状态 & "','" & 培训日期 & "','" & 行业类别 & "','" & 身份证号 & ")"
    Else
        '修改记录。
        dafuncGetData "update 健康证管理_办证申请信息表 set 体检号='" & 体检号 & "',姓名='" & 姓名 & "',性别='" & 性别 & "',年龄='" & 年龄 & "',种类='" & 种类 & "',职业='" & 职业 & "',民族='" & 民族 & "',申请编号='" & 申请编号 & "',单位名称='" & 单位名称 & "',体检日期='" & 体检日期 & "',体检结论='" & 体检结论 & "',培训结论='" & 培训结论 & "',检出病种='" & 检出病种 & "',处置='" & 处置 & "',培训日期='" & 培训日期 & "',行业类别='" & 行业类别 & "',身份证号='" & 身份证号 & "' where 系统编号='" & mstr系统编号 & "'"
    End If
    
    dafuncGetData "update 健康证管理_办证申请信息表 set 健康证号='" & 健康证号 & "' ,发证日期=" & IIf(发证日期 = "", "null", "'" & 发证日期 & "'") & ",有效期至=" & IIf(有效期至 = "", "null", "'" & 有效期至 & "'") & ",发证单位='" & 发证单位 & "',调离编号='" & 调离编号 & "',调离日期=" & IIf(调离日期 = "", "null", "'" & 调离日期 & "'") & ",调离期限='" & 调离期限 & "',备注='" & 备注 & "' where 系统编号='" & mstr系统编号 & "'"
    
    dasubCommitTran
    
    '保存照片为文件。
    If Not 照片 Is Nothing Then
'        If Dir(App.Path & "\Photo", vbDirectory) = "" Then
'            MkDir App.Path & "\Photo"
'        End If
'        SavePicture 照片, App.Path & "\Photo\" & mstr系统编号 & ".bmp"
        
        pmsub保存图片 照片, mstr系统编号, "健康证管理"
    End If
    
    Exit Sub
errhandler:
    sfsub错误处理 "健康证管理部件", "cls体检", "sub保存", Err.Number, Err.Description, True
End Sub

Public Sub sub删除()
    dafuncGetData "delete 健康证管理_办证申请信息表 where 系统编号='" & mstr系统编号 & "'"
    
'    '恢复体检访问标志。
'    If 体检系统编号 <> "" Then
'        dafuncGetData "update 体检管理_体检访问标志表 set 健康证='1' where 健康证='2' and 系统编号='" & 体检系统编号 & "'"
'    End If
End Sub


