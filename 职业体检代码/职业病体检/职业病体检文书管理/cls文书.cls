VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls文书"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'2012-05-29 陶露
'导出PDF格式的文件所需要的函数
Private Const HKEY_CURRENT_USER As Long = &H80000001
Private Const ERROR_SUCCESS As Long = 0&
Private Const REG_SZ As Long = 1
Private sPath As String
Private Const APPREGPATH As String = "Software\Default Printer Select Tool"
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
'2012-05-29

Private mblnPrint As Boolean
'作者：杨春。


'2012-02-15 于登淼
'修改内容：添加函数。职业病体检最开始，增加打印条码功能
Public Sub sub打印条码(ByVal para文书名称 As String, ByVal para体检条码号 As String, ByVal para是否预览 As Boolean)
    Sub打印文书 para文书名称, Nothing, para是否预览, False, para体检条码号
End Sub

'2012-04-09 陶露
'打印多个条码
Public Sub Sub打印多个文书(ByVal para文书名称 As String, _
                    ByVal para文书内容 As Object, _
                    ByVal para是否预览 As Boolean, _
                    Optional para是否只预览 As Boolean = False, _
                    Optional ByRef para体检条码号 As Collection)
    Dim lcol文书内容 As Collection
    Dim lcolItem As Collection
    Dim lobjRec As Object
    Dim para体检条码号组 As Collection
    Set para体检条码号组 = para体检条码号
    Dim str As String
    
    On Error GoTo errhandler
    
    'If para体检条码号组.Count - LBound(para体检条码号组) = 0 And (para文书内容 Is Nothing) Then
    '    Err.Raise 6666, , "文书内容记录为空，不能打印文书。"
    'End If
    'If para体检条码号组.Count - LBound(para体检条码号组) = 0 Then
    '    If para文书内容.recordcount = 0 Then Err.Raise 6666, , "文书内容记录为空，不能打印文书。"
    'End If
    
    Select Case para文书名称
    Case "仅打印条码"
    Dim i As Integer
    Dim j As Integer
    Dim n As Integer
    Dim page As Integer
    
    '使用数组时 打印的人数不能超过15个
    'For i = 0 To UBound(para体检条码号组) - 1
            'str = para体检条码号组(i)
            'If Len(str) <> 0 Then
                'frmOnlyBarCode.cbccBarCode(i).Value = para体检条码号组(i)
                'frmOnlyBarCode.cbccBarCode(i).Visible = True
                'If para是否只预览 Or para是否预览 Then
                    'frmOnlyBarCode.Show 1
                'End If
            'End If
        'Next i
    'frmOnlyBarCode.PrintForm
    
    '判断输出的页数
    If para体检条码号组.Count Mod 15 > 0 Then
        page = Int(para体检条码号组.Count / 15) + 1
    Else
        page = Int(para体检条码号组.Count / 15)
    End If
    
    n = 1 '记录集合里的个数
    For j = 1 To page
        For i = 0 To 14
            If n <= para体检条码号组.Count Then
                str = para体检条码号组.Item(n)
                If Len(str) <> 0 Then
                    frmOnlyBarCode.cbccBarCode(i).Value = str '输出条码
                    frmOnlyBarCode.cbccBarCode(i).Visible = True '条码可见
                    If para是否只预览 Or para是否预览 Then
                        frmOnlyBarCode.Show 1
                    End If
                End If
                n = n + 1
            Else
                Exit For
            End If
        Next i
        frmOnlyBarCode.PrintForm
        Unload frmOnlyBarCode '卸载对象
    Next j
    End Select
    Set frmOnlyBarCode = Nothing
Exit Sub

errhandler:
    sfsub错误处理 "职业病文书管理", "cls文书", "Sub打印文书", Err.Number, Err.Description, True
    Exit Sub
    Resume
End Sub
'2012-04-09 陶露


'功能：打印体检管理的当前系统编号的指定文书。
'输入：para文书名称
'      paraRec as Recordset 文书内容。
'2012-05-30 陶露
'添加了一个PDF的参数
Public Sub Sub打印文书(ByVal para文书名称 As String, _
                    ByVal para文书内容 As Object, _
                    ByVal para是否预览 As Boolean, _
                    Optional para是否只预览 As Boolean = False, _
                    Optional ByVal para体检条码号 As String = "", _
                    Optional ByVal para是否为PDF As Boolean = False)

    Dim lcol文书内容 As Collection
    Dim lcolItem As Collection
    Dim i As Long
    Dim lobjRec As Object, lobj数据对象 As Object
    Dim lobjSys As Object
    Dim lpicPhoto As StdPicture
    On Error GoTo errhandler
    
    If Len(para体检条码号) = 0 And (para文书内容 Is Nothing) Then
        Err.Raise 6666, , "文书内容记录为空，不能打印文书。"
    End If
    On Error Resume Next
    Kill "c:\职业健康体检_*.bmp"
    On Error GoTo errhandler
'    If Len(para体检条码号) = 0 Then
'        If para文书内容.RecordCount = 0 Then Err.Raise 6666, , "文书内容记录为空，不能打印文书。"
'    End If
    
    Select Case para文书名称
    
    '2012-02-16 于登淼 ↓
    '修改内容：仅打印传入的条码号
    'Case "仅打印条码"
        'With frmOnlyBarCode
            '.cbccBarCode.Value = para体检条码号
            'If para是否只预览 Or para是否预览 Then
                '.Show 1
            'End If
            '.PrintForm
        'End With
    '2012-02-16 于登淼 ↑
'    Case "体检登记表"
'        '套打。
'        With frm打印体检登记表
'            Set .pobj文书内容 = para文书内容
'            If para是否只预览 Or para是否预览 Then
'                .Show 1
'            End If
'            If Not para是否只预览 Then
'                .PrintForm
'            End If
'            Set .pobj文书内容 = Nothing
'        End With
        
'    Case "体检表" '修改：2003-4-16（杨春）增加体检表。
'        With frm打印体检表
'            Set .pobj文书内容 = para文书内容
'            If para是否只预览 Or para是否预览 Then
'                .Show 1
'            End If
'            If Not para是否只预览 Then
'                .PrintForm
'            End If
'        End With
'    Case "阳性体检单"
'        Set lcol文书内容 = New Collection
'        Set lcolItem = New Collection
'        lcolItem.Add para文书内容.Fields(i).Name, "名称"
'        lcolItem.Add IIf(IsNull(para文书内容!系统编号), "", para文书内容.Fields("系统编号").Value), "值"
'        lcol文书内容.Add lcolItem, lcolItem("名称")
'
'        '获取体检单耗。
'        Set lobjRec = dafuncGetData("select 项目值 from 体检管理_体检附加信息表 where 系统编号='" & para文书内容!系统编号 & "' and 附加项目='体检表号'")
'        Set lcolItem = New Collection
'        lcolItem.Add "体检表号", "名称"
'        If lobjRec.RecordCount > 0 Then
'            lcolItem.Add IIf(IsNull(lobjRec(0)), "", lobjRec(0).Value), "值"
'        Else
'            lcolItem.Add "", "值"
'        End If
'        lcol文书内容.Add lcolItem, lcolItem("名称")
'
'        Set lcolItem = New Collection
'        lcolItem.Add "体检单位", "名称"
'        lcolItem.Add um防疫站名, "值"
'        lcol文书内容.Add lcolItem, lcolItem("名称")
'
'        '获取照片。
'        Dim lpicPhoto As StdPicture
'        Dim lobjSys As Object
'        Set lobjSys = CreateObject("Scripting.FileSystemObject")
'        '先拷贝空白签名的图片。
'        lobjSys.CopyFile App.Path & "\空白照片.bmp", "c:\体检照片.bmp"
'
'        Set lpicPhoto = pmfunc获取图片(para文书内容!健康档案编号, "体检管理")
'        '2012-04-14 于登淼 ↓
'        '数据库里没有图片时，lpicPhoto返回值为0，而不是null
'        If lpicPhoto <> 0 Then
'            SavePicture lpicPhoto, "c:\体检照片.bmp"
'        End If
'        '2012-04-14 于登淼 ↑
'
'        funcPrintReport para文书名称, lcol文书内容, para是否预览 Or para是否只预览
'
'    Case "体检结果单"
'        Set lcol文书内容 = New Collection
'        For i = 0 To para文书内容.Fields.Count - 1
'            Set lcolItem = New Collection
'            lcolItem.Add para文书内容.Fields(i).Name, "名称"
'            lcolItem.Add IIf(IsNull(para文书内容.Fields(i).Value), "", para文书内容.Fields(i).Value), "值"
'            lcol文书内容.Add lcolItem
'        Next
'
'        Set lcolItem = New Collection
'        lcolItem.Add "体检单位", "名称"
'        lcolItem.Add um防疫站名, "值"
'        lcol文书内容.Add lcolItem, lcolItem("名称")
'
'        funcPrintReport para文书名称, lcol文书内容, para是否预览 Or para是否只预览
'
'    Case "体检结果通知单"
'        '把文书内容按水晶报表打印的要求存放在collection中。
'        Set lcol文书内容 = New Collection
'        For i = 0 To para文书内容.Fields.Count - 1
'            Set lcolItem = New Collection
'            lcolItem.Add para文书内容.Fields(i).Name, "名称"
'            lcolItem.Add IIf(IsNull(para文书内容.Fields(i).Value), "", para文书内容.Fields(i).Value), "值"
'            lcol文书内容.Add lcolItem
'        Next
'
'        Set lcolItem = New Collection
'        lcolItem.Add "体检单位", "名称"
'        lcolItem.Add um防疫站名, "值"
'        lcol文书内容.Add lcolItem, "tjdw"
'
'        funcPrintReport para文书名称, lcol文书内容, para是否预览 Or para是否只预览
'
'    Case "B超体检结果单"
'        Set lcol文书内容 = New Collection
'        Set lcolItem = New Collection
'        lcolItem.Add "系统编号", "名称"
'        lcolItem.Add para文书内容("系统编号").Value, "值"
'        funcPrintReport para文书名称, lcol文书内容, para是否预览 Or para是否只预览
'
''    Case "化验单"
''        Dim lobj化验单 As Object
''        '创建化验单对象。
''        Set lobj化验单 = CreateObject("化验单对象部件.clsBigCheckSheet")
''        lobj化验单.化验单编号 = para文书内容!化验单编号
''
''        '启动化验单打印界面。
''        Set frm打印化验单.pobj化验单对象 = lobj化验单
''
''        '先预览。
''        If para是否预览 Or para是否只预览 Then
''            frm打印化验单.Move 1330, 400
''            frm打印化验单.Show 1
''        End If
''        If Not para是否只预览 Then
''            frm打印化验单.c化验单.subPrint
''        End If
''        Set frm打印化验单.pobj化验单对象 = Nothing
'    Case "放射工作体检登记单", "职业健康体检登记单", "涉核部队体检登记单"
'
'        Set lcol文书内容 = New Collection
'        Set lcolItem = New Collection
'        lcolItem.Add para文书内容.Fields(i).Name, "名称"
'        lcolItem.Add IIf(IsNull(para文书内容!系统编号), "", para文书内容.Fields("系统编号").Value), "值"
'        lcol文书内容.Add lcolItem, lcolItem("名称")
'
'        Set lcolItem = New Collection
'        lcolItem.Add "体检单位", "名称"
'        lcolItem.Add um防疫站名, "值"
'        lcol文书内容.Add lcolItem, lcolItem("名称")
'
'        '获取照片。
'        'Dim lpicPhoto As StdPicture
'        'Dim lobjSys As Object
'        Set lobjSys = CreateObject("Scripting.FileSystemObject")
'        '先拷贝空白签名的图片。
'        lobjSys.CopyFile App.Path & "\空白照片.bmp", "c:\体检照片.bmp"
'
'        Set lpicPhoto = pmfunc获取图片(para文书内容!系统编号, "职业病体检")
'        '2012-04-14 于登淼 ↓
'        '数据库里没有图片时，lpicPhoto返回值为0，而不是null
'        If lpicPhoto <> 0 Then
'            SavePicture lpicPhoto, "c:\体检照片.bmp"
'        End If
'        '2012-04-14 于登淼 ↑
'
'        funcPrintReport para文书名称, lcol文书内容, para是否预览 Or para是否只预览
    
    '2012-04-13 于登淼 ↓
    '增加职业健康体检_涉核部队YK  2015-11-24 by 牟俊
    Case "职业病体检_最终结论", "职业健康体检_普通体检", "职业健康体检_放射健康", "职业健康体检_涉核部队", "职业健康体检_8023", "职业健康体检_8023部队", "职业健康体检_职业健康" _
            , "职业健康体检_普通体检F", "职业健康体检_放射健康F", "职业健康体检_涉核部队F", "职业健康体检_8023F", "职业健康体检_职业健康F", "职业健康体检_涉核部队YK"
'    Case "职业病体检_最终结论", "职业健康体检_普通体检", "职业健康体检_放射健康", "职业健康体检_涉核部队", "职业健康体检_8023", "职业健康体检_8023部队", "职业健康体检_职业健康" _
'            , "职业健康体检_普通体检F", "职业健康体检_放射健康F", "职业健康体检_涉核部队F", "职业健康体检_8023F", "职业健康体检_职业健康F"
        
        Set lcol文书内容 = New Collection
        Set lcolItem = New Collection
'        lcolItem.Add para文书内容.Fields(i).Name, "名称"
'        lcolItem.Add IIf(IsNull(para文书内容!系统编号), "", para文书内容.Fields("系统编号").Value), "值"
        lcolItem.Add para文书内容(1), "值"
        lcolItem.Add "系统编号", "名称"
        lcol文书内容.Add lcolItem, lcolItem("名称")
        
        '2012-05-31 陶露
        '获取五官科中的眼科照片
        Set lobjSys = CreateObject("Scripting.FileSystemObject")
        '先拷贝空白签名的图片。
        lobjSys.CopyFile App.Path & "\眼科照片.bmp", "c:\职业健康体检_晶状体环面及正面图.bmp"
        
       
        
        Set lpicPhoto = pmfunc获取眼科照片(para文书内容(1))
'        On Error Resume Next
        '数据库里没有五官图片时，lpicPhoto返回值为nothing，而不是null
        If Not (lpicPhoto Is Nothing) Then
            SavePicture lpicPhoto, "c:\职业健康体检_晶状体环面及正面图.bmp"
        End If
        '获取照片。
'        Dim lpicPhoto As StdPicture
'        Dim lobjSys As Object
'        Set lobjSys = CreateObject("Scripting.FileSystemObject")
        '先拷贝空白签名的图片。
        lobjSys.CopyFile App.Path & "\空白照片.bmp", "c:\职业健康体检_照片.bmp"

        Set lpicPhoto = pmfunc获取图片(para文书内容(1), "职业病体检")

        '2012-04-14 于登淼 ↓
        '数据库里没有图片时，lpicPhoto返回值为0，而不是null
        If Not (lpicPhoto Is Nothing Or lpicPhoto = 0) Then
            SavePicture lpicPhoto, "c:\职业健康体检_照片.bmp"
        End If
        
        
                '获取签名。
        'Dim lpicPhoto As StdPicture
        'Dim lobjSys As Object
'        Set lobjSys = CreateObject("Scripting.FileSystemObject")
        '先拷贝空白签名的图片。
        lobjSys.CopyFile App.Path & "\空白签名.bmp", "c:\职业健康体检_B超.bmp"
        lobjSys.CopyFile App.Path & "\空白签名.bmp", "c:\职业健康体检_X光.bmp"
        lobjSys.CopyFile App.Path & "\空白签名.bmp", "c:\职业健康体检_电测听.bmp"
        lobjSys.CopyFile App.Path & "\空白签名.bmp", "c:\职业健康体检_肺功能.bmp"
        lobjSys.CopyFile App.Path & "\空白签名.bmp", "c:\职业健康体检_免疫科.bmp"
        lobjSys.CopyFile App.Path & "\空白签名.bmp", "c:\职业健康体检_内科.bmp"
        lobjSys.CopyFile App.Path & "\空白签名.bmp", "c:\职业健康体检_尿常规.bmp"
        lobjSys.CopyFile App.Path & "\空白签名.bmp", "c:\职业健康体检_染色体.bmp"
        lobjSys.CopyFile App.Path & "\空白签名.bmp", "c:\职业健康体检_生化科.bmp"
        lobjSys.CopyFile App.Path & "\空白签名.bmp", "c:\职业健康体检_外科.bmp"
        lobjSys.CopyFile App.Path & "\空白签名.bmp", "c:\职业健康体检_五官科.bmp"
        lobjSys.CopyFile App.Path & "\空白签名.bmp", "c:\职业健康体检_心电科.bmp"
        lobjSys.CopyFile App.Path & "\空白签名.bmp", "c:\职业健康体检_血常规.bmp"
        lobjSys.CopyFile App.Path & "\空白签名.bmp", "c:\职业健康体检_总检.bmp"
 
        Set lobj数据对象 = CreateObject("职业病对象.clsmanagemedicalexam")
        Set lobjRec = lobj数据对象.func查询医师编号(para文书内容(1))
        On Error Resume Next
        Set lpicPhoto = pmfunc获取图片(lobjRec(0), "系统管理")
        For i = 0 To lobjRec.RecordCount - 1
            Set lpicPhoto = pmfunc获取图片(lobjRec("医生编号"), "系统管理")
            If lobjRec("医生编号") <> "" And lobjRec("科室") = "01" Then
                If Not (lpicPhoto Is Nothing Or lpicPhoto = 0) Then
                    SavePicture lpicPhoto, "c:\职业健康体检_五官科.bmp"
                End If
            End If
            If lobjRec("医生编号") <> "" And lobjRec("科室") = "02" Then
                If Not (lpicPhoto Is Nothing Or lpicPhoto = 0) Then
                    SavePicture lpicPhoto, "c:\职业健康体检_内科.bmp"
                End If
            End If
            If lobjRec("医生编号") <> "" And lobjRec("科室") = "03" Then
                If Not (lpicPhoto Is Nothing Or lpicPhoto = 0) Then
                    SavePicture lpicPhoto, "c:\职业健康体检_外科.bmp"
                End If
            End If
            If lobjRec("医生编号") <> "" And lobjRec("科室") = "04" Then
                If Not (lpicPhoto Is Nothing Or lpicPhoto = 0) Then
                    SavePicture lpicPhoto, "c:\职业健康体检_血常规.bmp"
                End If
            End If
            If lobjRec("医生编号") <> "" And lobjRec("科室") = "05" Then
                If Not (lpicPhoto Is Nothing Or lpicPhoto = 0) Then
                    SavePicture lpicPhoto, "c:\职业健康体检_免疫科.bmp"
                End If
            End If
            If lobjRec("医生编号") <> "" And lobjRec("科室") = "06" Then
                If Not (lpicPhoto Is Nothing Or lpicPhoto = 0) Then
                    SavePicture lpicPhoto, "c:\职业健康体检_尿常规.bmp"
                End If
            End If
            If lobjRec("医生编号") <> "" And lobjRec("科室") = "07" Then
                If Not (lpicPhoto Is Nothing Or lpicPhoto = 0) Then
                    SavePicture lpicPhoto, "c:\职业健康体检_染色体.bmp"
                End If
            End If
            If lobjRec("医生编号") <> "" And lobjRec("科室") = "08" Then
                If Not (lpicPhoto Is Nothing Or lpicPhoto = 0) Then
                    SavePicture lpicPhoto, "c:\职业健康体检_电测听.bmp"
                End If
            End If
            If lobjRec("医生编号") <> "" And lobjRec("科室") = "09" Then
                If Not (lpicPhoto Is Nothing Or lpicPhoto = 0) Then
                    SavePicture lpicPhoto, "c:\职业健康体检_X光.bmp"
                End If
            End If
            If lobjRec("医生编号") <> "" And lobjRec("科室") = "10" Then
                If Not (lpicPhoto Is Nothing Or lpicPhoto = 0) Then
                    SavePicture lpicPhoto, "c:\职业健康体检_心电科.bmp"
                End If
            End If
            If lobjRec("医生编号") <> "" And lobjRec("科室") = "11" Then
                If Not (lpicPhoto Is Nothing Or lpicPhoto = 0) Then
                    SavePicture lpicPhoto, "c:\职业健康体检_B超.bmp"
                End If
            End If
            If lobjRec("医生编号") <> "" And lobjRec("科室") = "12" Then
                If Not (lpicPhoto Is Nothing Or lpicPhoto = 0) Then
                    SavePicture lpicPhoto, "c:\职业健康体检_肺功能.bmp"
                End If
            End If
            If lobjRec("医生编号") <> "" And lobjRec("科室") = "16" Then
                If Not (lpicPhoto Is Nothing Or lpicPhoto = 0) Then
                    SavePicture lpicPhoto, "c:\职业健康体检_总检.bmp"
                End If
            End If
            If lobjRec("医生编号") <> "" And lobjRec("科室") = "17" Then
                If Not (lpicPhoto Is Nothing Or lpicPhoto = 0) Then
                    SavePicture lpicPhoto, "c:\职业健康体检_生化科.bmp"
                End If
            End If
            lobjRec.movenext
        Next
        
        On Error GoTo errhandler
              '2012-04-14 于登淼 ↑
     '   If para是否预览 = False Then ' And para是否为PDF = False Then
     '       funcPrintReport para文书名称, lcol文书内容, False 'para是否只预览   '预览，但有打印按钮
      '  ElseIf para是否只预览 = True Then 'And para是否为PDF = False Then
            funcPrintReport para文书名称, lcol文书内容, para是否预览 'Or para是否只预览   '直接打印，不预览
      '  Else
      '      funcPrintReport para文书名称, lcol文书内容, para是否预览 Or para是否只预览, para是否为PDF  '直接打印，不预览
      '  End If
    '2012-04-13 于登淼 ↑
        On Error Resume Next
        Kill "c:\职业健康体检_*.bmp"
        On Error GoTo errhandler
    '2012-04-28 陶露
    '打印单位报表
    Case "职业病体检_单位报表"
        Set lcol文书内容 = New Collection
        Set lcolItem = New Collection
        lcolItem.Add para文书内容.Fields(i).Name, "名称"
        
        lcolItem.Add IIf(IsNull(para文书内容!系统编号), "", para文书内容.Fields("系统编号").Value), "值"
        lcol文书内容.Add lcolItem, lcolItem("名称")
        
        Set lcolItem = New Collection
        lcolItem.Add "体检单位", "名称"
        lcolItem.Add IIf(IsNull(para文书内容!系统编号), "", para文书内容.Fields("系统编号").Value), "值"
            
        If para是否只预览 = False Then
            funcPrintReport para文书名称, lcol文书内容, False, para是否为PDF 'para是否只预览   '预览，但有打印按钮
        Else
            funcPrintReport para文书名称, lcol文书内容, para是否预览 Or para是否只预览, para是否为PDF  '直接打印，不预览
        End If
    '2012-04-28
    
    '2012-06-06 于登淼 ↓
    '添加多种统计报告的预览和打印
    Case "职业病体检_统计报告(人数)", "职业病体检_统计报告(金额)", "职业病体检_统计报告(合格率)", "职业病体检_统计报告(按体检情况)"
        If para是否只预览 = False Then
            funcPrintReport para文书名称, lcol文书内容, False, para是否为PDF 'para是否只预览   '预览，但有打印按钮
        Else
            funcPrintReport para文书名称, lcol文书内容, para是否预览 Or para是否只预览, para是否为PDF  '直接打印，不预览
        End If
    '2012-06-06 于登淼 ↑
    Case Else
        '体检单，要打印条形码和照片，所以采用窗体打印。
        With frm打印体检单
            '设置那些条码需要打印。
            Select Case para文书名称
            Case "尿常规体检单"
                .pbln是否有尿常规 = True
                .pbln是否有大便培养 = False
                .pbln是否有验血 = False
                
            Case "大便培养体检单"
                .pbln是否有尿常规 = False
                .pbln是否有大便培养 = True
                .pbln是否有验血 = False
            
            Case "验血体检单"
                .pbln是否有尿常规 = False
                .pbln是否有大便培养 = False
                .pbln是否有验血 = True
            
            Case "尿常规大便培养体检单"
                .pbln是否有尿常规 = True
                .pbln是否有大便培养 = True
                .pbln是否有验血 = False
                
            Case "尿常规验血体检单"
                .pbln是否有尿常规 = True
                .pbln是否有大便培养 = True
                .pbln是否有验血 = True
            
            Case "大便培养验血体检单"
                .pbln是否有尿常规 = False
                .pbln是否有大便培养 = True
                .pbln是否有验血 = True
            
            Case Else '"尿常规大便培养验血体检单"
                .pbln是否有尿常规 = True
                .pbln是否有大便培养 = True
                .pbln是否有验血 = True
            End Select
            
            '设置文书内容。
            
            Set .pobj文书内容 = para文书内容
            
            If para是否只预览 Or para是否预览 Then
                .Show 1
            End If
            
            If Not para是否只预览 Then
                '打印窗体。
                .PrintForm
            End If
            
            Set .pobj文书内容 = Nothing
            
        End With
    End Select
    Exit Sub
errhandler:
    sfsub错误处理 "职业病文书管理", "cls文书", "Sub打印文书", Err.Number, Err.Description, True
    Exit Sub
    Resume
End Sub

'功能：获取体检单类型。
'作者：杨春。
Public Function func获取所有体检单类型() As Collection
    On Error GoTo errhandler
    Set func获取所有体检单类型 = New Collection
    With func获取所有体检单类型
        .Add "尿常规体检单", "尿常规体检单"
        .Add "大便培养体检单", "大便培养体检单"
        .Add "验血体检单", "验血体检单"
        
        .Add "尿常规大便培养体检单", "尿常规大便培养体检单"
        .Add "尿常规验血体检单", "尿常规验血体检单"
        .Add "大便培养验血体检单", "大便培养验血体检单"
        
        .Add "尿常规大便培养验血体检单", "尿常规大便培养验血体检单"
        
    End With

    Exit Function
errhandler:
    sfsub错误处理 "职业病文书管理", "cls文书", "func获取所有体检单类型", Err.Number, Err.Description, True
End Function


Public Function funcPrintReport(ByVal para文书名称 As String, ByVal para文书内容 As Collection, ByVal para是否预览 As Boolean, Optional ByVal para是否为PDF As Boolean = False) As Boolean

    Dim i As Integer
    Dim lstrItem As String
    
    On Error GoTo errhandler

    
    Dim capp As New CRAXDRT.Application
    Dim cr As CRAXDRT.Report
    Dim lstrDatabase As String
    Dim lstrServer As String
'    lstrDatabase = sffuncGetSetting("系统管理", "数据库配置", "数据库名")
    If InStr(dapstrDSN, "Initial Catalog=") > 0 Then
        lstrDatabase = Right(dapstrDSN, Len(dapstrDSN) - InStr(dapstrDSN, "Initial Catalog=") - Len("Initial Catalog=") + 1)
        If InStr(lstrDatabase, ";") > 0 Then
            lstrDatabase = Left(lstrDatabase, InStr(lstrDatabase, ";") - 1)
        End If
    ElseIf InStr(dapstrDSN, "Database=") > 0 Then
        lstrDatabase = Right(dapstrDSN, Len(dapstrDSN) - InStr(dapstrDSN, "Database=") - Len("Database=") + 1)
        If InStr(lstrDatabase, ";") > 0 Then
            lstrDatabase = Left(lstrDatabase, InStr(lstrDatabase, ";") - 1)
        End If
    Else
        lstrDatabase = "防疫2001"
    End If
    If InStr(dapstrDSN, "Data Source=") > 0 Then
        lstrServer = Right(dapstrDSN, Len(dapstrDSN) - InStr(dapstrDSN, "Data Source=") - Len("Data Source=") + 1)
        If InStr(lstrServer, ";") > 0 Then
            lstrServer = Left(lstrServer, InStr(lstrServer, ";") - 1)
        End If
    ElseIf InStr(dapstrDSN, "Server=") > 0 Then
        lstrServer = Right(dapstrDSN, Len(dapstrDSN) - InStr(dapstrDSN, "Server=") - Len("Server=") + 1)
        If InStr(lstrServer, ";") > 0 Then
            lstrServer = Left(lstrServer, InStr(lstrServer, ";") - 1)
        End If
    Else
        lstrServer = "local"
    End If
    
    
    '设置报表文件。
    Set cr = capp.OpenReport(App.Path & "\" & para文书名称 & ".rpt")
    '连接到数据库。
    On Error Resume Next
'    For i = 1 To cr.Database.Tables.Count
'        cr.Database.Tables.Item(i).ConnectionProperties("User ID").Value = "user26"
'        cr.Database.Tables.Item(i).ConnectionProperties("Password").Value = "welcome"
'        cr.Database.Tables.Item(i).ConnectionProperties("Database").Value = lstrDatabase
'        If InStr(cr.Database.Tables.Item(i).Location, ".dbo.") > 0 Then
'            cr.Database.Tables.Item(i).Location = lstrDatabase & ".dbo." & Right(cr.Database.Tables.Item(i).Location, Len(cr.Database.Tables.Item(i).Location) - InStr(cr.Database.Tables.Item(i).Location, ".dbo.") + 1)
'        Else
'            cr.Database.Tables.Item(i).Location = lstrDatabase & ".dbo." & cr.Database.Tables.Item(i).Location
'        End If
'        If Err.Number <> 0 Then Exit For
'    Next
    
    
    '使用oledb连接
'    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo errhandler
        For i = 1 To cr.Database.Tables.Count
            cr.Database.Tables(i).SetLogOnInfo lstrServer, lstrDatabase, "user26", "welcome"
        Next
        cr.Database.LogOnServerEx "p2soledb.dll", lstrServer, "", "user26", "welcome", "OLE DB", "Provider=SQLOLEDB.1;Password=welcome;Persist Security Info=True;User ID=user26;Initial Catalog=" & lstrDatabase & ";Data Source=" & lstrServer

        For i = 1 To cr.Database.Tables.Count
            If InStr(cr.Database.Tables.Item(i).Location, ".dbo.") > 0 Then
                cr.Database.Tables.Item(i).Location = lstrDatabase & ".dbo." & Right(cr.Database.Tables.Item(i).Location, Len(cr.Database.Tables.Item(i).Location) - InStr(cr.Database.Tables.Item(i).Location, ".dbo.") + 1)
            Else
                cr.Database.Tables.Item(i).Location = lstrDatabase & ".dbo." & cr.Database.Tables.Item(i).Location
            End If
        Next
'    End If
    
'    If Err.Number <> 0 Then
'        Err.Clear
        
'        For i = 1 To cr.Database.Tables.Count
'            Dim dbTable As Object
'            Dim cpProperties As Object
'            Dim lstrLocation As String
'            Set dbTable = cr.Database.Tables(i)
'            lstrLocation = dbTable.Location
'            Set cpProperties = dbTable.ConnectionProperties
'            dbTable.DllName = "crdb_ado.dll"
'            cpProperties.DeleteAll
'            cpProperties.Add "Provider", "SQLOLEDB"
'            cpProperties.Add "Data Source", lstrServer
'            cpProperties.Add "Initial Catalog", lstrDatabase
'            cpProperties.Add "User ID", "user26"
'            cpProperties.Add "Password", "welcome"
'            dbTable.Location = lstrLocation
'        Next
'    End If
    
    '设置打印参数
    Dim Item As CRAXDRT.FormulaFieldDefinition
    For i = 1 To para文书内容.Count
        lstrItem = Trim(para文书内容(i)("值"))
        '把回车符替换掉。
        If InStr(1, lstrItem, Chr(13) & Chr(10)) > 0 Then
            lstrItem = Replace(lstrItem, Chr(13) & Chr(10), " ")
        End If
                
        On Error Resume Next

        Set Item = cr.FormulaFields.GetItemByName(para文书内容(i)("名称"))

        Item.Text = "'" & lstrItem & "'"
        On Error GoTo errhandler
    Next i
    
    cr.DiscardSavedData
       
'    '2012-05-29 陶露
'    '判断是否需要打印出PDF格式的文件
'    If para是否为PDF = True Then
'
'        '修改为需要的打印机
'        Select Case WinVer
'            Case "NT", "00", "XP"
'            sPath = "Software\Microsoft\Windows NT\CurrentVersion\Windows"
'            Case "95", "98", "ME"
'            sPath = "Software\Microsoft\Windows\CurrentVersion\Windows"
'        End Select
'
'            '获取原有的默认打印机的名称
'            Dim sOrigPrinter As String
'            'Remember the current default printer to allow for restoring after a crash
'            sOrigPrinter = GetString(HKEY_CURRENT_USER, sPath, "Device")
'            Call SaveString(HKEY_CURRENT_USER, APPREGPATH, "Printer Name", sOrigPrinter)
'            Call SaveString(HKEY_CURRENT_USER, APPREGPATH, "Shutdown Properly", "0")
'            sOrigPrinter = GetString(HKEY_CURRENT_USER, sPath, "Device")
'
'            '传递需要用的PDF打印机的名称
'            Call SetDefaultPrinter("PDFCreator,winspool,Ne01:")
'            cr.PrintOut False
'
'            Call SetDefaultPrinter(sOrigPrinter)
'
'            '恢复打印机为修改前的默认打印机
'            'Restore original default printer if so desired
'            Call SaveString(HKEY_CURRENT_USER, sPath, "Device", sOrigPrinter)
'            'Indicate in the registry that we didn't crash and the original printer name is not needed
'            Call SaveString(HKEY_CURRENT_USER, APPREGPATH, "Shutdown Properly", "1")
'            'Remember whether to confirm or not
'            Call SaveString(HKEY_CURRENT_USER, APPREGPATH, "Confirm", 1)
'
'            'Call SetDefaultPrinter(sOrigPrinter)
'            Exit Function
'    End If
'    '2012-05-29
    
    '2012-06-01 陶露
    '判断是否需要打印出PDF格式的文件
    If para是否为PDF = True Then
        Dim objNetwork  As Object
        Set objNetwork = CreateObject("WScript.Network")                '建立WScript.Network对象
        objNetwork.SetDefaultPrinter "PDFCreator"                       '设置打印PDF格式需要的打印机
        Set objNetwork = Nothing
        cr.PrintOut False
            
        Set objNetwork = CreateObject("WScript.Network")                '建立WScript.Network对象
        objNetwork.SetDefaultPrinter "发送至 OneNote 2007"              '把打印机设置回以前的默认打印机
        Set objNetwork = Nothing
    Else
    '2012-06-01
    
        '显示报表
        Dim lobjFrm As New frmPrint
        With lobjFrm.cRepPrint
            If para是否预览 = True Then
                If InStr(para文书名称, "_") > 0 Then
                    lobjFrm.Caption = Right(para文书名称, Len(para文书名称) - InStr(para文书名称, "_"))
                Else
                    lobjFrm.Caption = para文书名称
                End If
                .ReportSource = cr
                '2012-04-13 于登淼 ↓
                '如果预览，则隐藏打印按钮(预览为false时，直接打印，不预览)
                .EnablePrintButton = False
                '2012-04-13 于登淼 ↑
                .ViewReport
                lobjFrm.Show
            Else
'                cr.PrintOut False
                cr.printerduplex() = crPRDPVertical  '设置打印机双面打印  2016-4-21 by 牟俊
                cr.PrintOut False
'                cr.PrintOutEx False
               
            End If
        End With
    End If
    
    funcPrintReport = True
    Exit Function
    
errhandler:
    Dim llngErr As Long
    Dim lstrError As String
    llngErr = Err.Number
    lstrError = Err.Description
    If llngErr = 20599 Then
        '连接数据库失败，重试。
        i = i + 1
        If i < 3 Then
            Resume
        Else
            lstrError = "连接数据库失败，可能是网络阻塞引起。请找系统管理员解决。"
        End If
    ElseIf llngErr = -2147206461 Then
        lstrError = "文书模板“" & App.Path & "\" & para文书名称 & ".rpt" & "”不存在！" & Chr(13) & Chr(10) & "请找到这个文件并拷贝到" & App.Path & "下面后再打印。"
    Else
        lstrError = func错误处理(llngErr, lstrError)
    End If
    lstrError = "打印水晶文书“" & para文书名称 & "”时失败。原因是：" & lstrError
    
    sfsub错误处理 "文书管理部件", "cls文书集", "funcPrintReport", 6666, lstrError, True
    funcPrintReport = False
    Exit Function
    Resume
End Function


Private Function func错误处理(ByVal paraErrNumber As Long, ByVal paraErrDes As String) As String
    Select Case paraErrNumber
    Case -2147217833
        func错误处理 = "输入数据过长（或过大），已超过系统规定长度（或大小）。"
    Case 6
        func错误处理 = "输入数据过大，已超过系统规定大小。"
    Case -2147217913
        func错误处理 = "日期格式非法！"
    Case 94 '无效使用Null。
        func错误处理 = "使用的字典项被人通过字典管理操作删除了，系统无法再继续正常处理。请找系统管理员恢复字典内容。请注意，不要随便删除字典项！"
    Case 336, 337, 338, 429, 430
        func错误处理 = "系统部件已损坏（或已丢失），系统无法再正常运行。请退出系统，并重新安装系统。"
    Case 440 '外部对象错误：类自动错误。
        func错误处理 = "系统部件不正常终止运行。请退出系统，再重新启动系统。"
    Case 91 '对象没有初始化成功。
        func错误处理 = "因为网络阻塞，系统启动功能时无法完成正常的初始化。请退出功能界面，再重新进入功能界面。"
    Case 20504, 20507
        func错误处理 = "在系统路径下未找到要打印的文件模板文件。解决办法：" & Chr(13) & Chr(10) & Chr(13) & Chr(10) _
                     & " (1) 请检查“文书管理部件.dll”系统文件所在路径下是否存在该水晶文书模板。" & Chr(13) & Chr(10) _
                     & " (2) 请进入“业务设置”的“文书格式设置”，在文书的“说明”信息中准确设置要打印的文书模板文件名（不带后缀）。"
    Case 20510
        func错误处理 = "水晶报表公式与获取文书内容的存储过程不一致（存储过程返回字段在水晶报表上不存在）。"
    Case 20526, 20545
        func错误处理 = "存在妨碍打印的问题。此错误产生的原因及解决方法如下： " & Chr(13) & Chr(10) & Chr(13) & Chr(10) _
                    & "  (1) 没有从 Windows 控制面板中安装打印机？" & Chr(13) & Chr(10) _
                    & "      解决：打开控制面板，双击“打印机”图标，选择“添加打印机”以装入打印机。" & Chr(13) & Chr(10) & Chr(13) & Chr(10) _
                    & "  (2) 打印机没在线？" & Chr(13) & Chr(10) _
                    & "      解决：检查打印机与计算机的连接是否正常。" & Chr(13) & Chr(10) & Chr(13) & Chr(10) _
                    & "  (3)  打印机阻塞或缺纸？" & Chr(13) & Chr(10) _
                    & "      解决：解决这些问题。" & Chr(13) & Chr(10) & Chr(13) & Chr(10) _
                    & "  (4) 试图在只能接受文本的打印机上打印窗体？" & Chr(13) & Chr(10) _
                    & "      解决：切换到一台能打印图形的打印机。"
        
    Case Else
        func错误处理 = paraErrDes
    End Select
End Function


'数据库连接，打印参数设置，打印
Public Function funcCHPrintReport(ByVal para文书名称 As String, ByVal para文书内容 As Collection, ByVal paraPath As String, ByVal para是否预览 As Boolean) As Boolean
    Dim i As Integer
    Dim lstrItem As String
    
    On Error GoTo errhandler
    
    Dim capp As New CRAXDRT.Application
    Dim cr As CRAXDRT.Report
    Dim lstrDatabase As String
    Dim lstrServer As String
    
'    lstrDatabase = sffuncGetSetting("系统管理", "数据库配置", "数据库名")
    If InStr(dapstrDSN, "Initial Catalog=") > 0 Then
        lstrDatabase = Right(dapstrDSN, Len(dapstrDSN) - InStr(dapstrDSN, "Initial Catalog=") - Len("Initial Catalog=") + 1)
        If InStr(lstrDatabase, ";") > 0 Then
            lstrDatabase = Left(lstrDatabase, InStr(lstrDatabase, ";") - 1)
        End If
    ElseIf InStr(dapstrDSN, "Database=") > 0 Then
        lstrDatabase = Right(dapstrDSN, Len(dapstrDSN) - InStr(dapstrDSN, "Database=") - Len("Database=") + 1)
        If InStr(lstrDatabase, ";") > 0 Then
            lstrDatabase = Left(lstrDatabase, InStr(lstrDatabase, ";") - 1)
        End If
    ElseIf InStr(dapstrDSN, "/") Then   'Internet运行模式
        lstrDatabase = paraPath & "\xkz.mdb"
    Else
        lstrDatabase = "防疫2001"
    End If
    If InStr(dapstrDSN, "Data Source=") > 0 Then
        lstrServer = Right(dapstrDSN, Len(dapstrDSN) - InStr(dapstrDSN, "Data Source=") - Len("Data Source=") + 1)
        If InStr(lstrServer, ";") > 0 Then
            lstrServer = Left(lstrServer, InStr(lstrServer, ";") - 1)
        End If
    ElseIf InStr(dapstrDSN, "Server=") > 0 Then
        lstrServer = Right(dapstrDSN, Len(dapstrDSN) - InStr(dapstrDSN, "Server=") - Len("Server=") + 1)
        If InStr(lstrServer, ";") > 0 Then
            lstrServer = Left(lstrServer, InStr(lstrServer, ";") - 1)
        End If
    Else
        lstrServer = "local"
    End If
    
    
    '设置报表文件。
    If paraPath = "" Then paraPath = App.Path
    If Right(paraPath, 1) <> "\" Then paraPath = paraPath & "\"
    Set cr = capp.OpenReport(paraPath & para文书名称 & ".rpt")
    
    On Error Resume Next

    
    '改用oledb连接到数据库。
    Dim CPProperties As CRAXDRT.ConnectionProperties
    Dim DBTable As CRAXDRT.DatabaseTable
    For i = 1 To cr.Database.Tables.Count
        Set DBTable = cr.Database.Tables.Item(i)
        Set CPProperties = DBTable.ConnectionProperties
        DBTable.DllName = "crdb_ado.dll"
        CPProperties.DeleteAll
        If InStr(dapstrDSN, "/") > 0 Then
            CPProperties.Add "Provider", "Microsoft.Jet.OLEDB.4.0"
            CPProperties.Add "Data Source", lstrDatabase
        Else
            CPProperties.Add "Provider", "SQLOLEDB"
            CPProperties.Add "Data Source", lstrServer
            CPProperties.Add "Initial Catalog", lstrDatabase
            CPProperties.Add "User ID", "user26"
            CPProperties.Add "Password", "welcome"
            'DBTable.Location = ""
            If InStr(DBTable.Location, ".dbo.") > 0 Then
                DBTable.Location = lstrDatabase & ".dbo." & Right(DBTable.Location, Len(DBTable.Location) - InStr(DBTable.Location, ".dbo.") + 1)
            Else
                DBTable.Location = lstrDatabase & ".dbo." & DBTable.Location
            End If
        End If
    Next
    
    '设置打印参数
    Dim Item As CRAXDRT.FormulaFieldDefinition
    For i = 1 To para文书内容.Count
        lstrItem = Trim(para文书内容(i)("值"))
        '把回车符替换掉。
        If InStr(1, lstrItem, Chr(13) & Chr(10)) > 0 Then
            lstrItem = gffuncStrReplace(lstrItem, Chr(13) & Chr(10), " ")
        End If
                
        On Error Resume Next

        Set Item = cr.FormulaFields.GetItemByName(para文书内容(i)("名称"))
        If InStr(lstrItem, "'") > 0 Then
            Item.Text = lstrItem
        Else
            Item.Text = "'" & lstrItem & "'"
        End If
        On Error GoTo errhandler
    Next i
    
    cr.DiscardSavedData
    '2012-01-17  刘云乐
    '因为点打印后都不预览，在此实现打印两次
    '显示报表
    'Dim lobjFrm As New frm打印水晶报表
    'With lobjFrm.cRepPrint
    '    If para是否预览 = True Then
    '        If InStr(para文书名称, "_") > 0 Then
     '           lobjFrm.Caption = Right(para文书名称, Len(para文书名称) - InStr(para文书名称, "_"))
     '       Else
      '          lobjFrm.Caption = para文书名称
      '      End If
      ''      .ReportSource = cr
      '      .ViewReport
      '      lobjFrm.Show
      '  Else
            cr.PrintOut False

       ' End If
    'End With
    '2012-01-17
    '函数返回
    funcCHPrintReport = True
    Exit Function
    
errhandler:
    Dim llngErr As Long
    Dim lstrError As String
    llngErr = Err.Number
    lstrError = Err.Description
    If llngErr = 20599 Then
        '连接数据库失败，重试。
        i = i + 1
        If i < 3 Then
            Resume
        Else
            lstrError = "连接数据库失败，可能是网络阻塞引起。请找系统管理员解决。"
        End If
    ElseIf llngErr = -2147206461 Then
        lstrError = "文书模板“" & paraPath & para文书名称 & ".rpt" & "”不存在！" & Chr(13) & Chr(10) & "请找到这个文件并拷贝到" & paraPath & "下面后再打印。"
    Else
        lstrError = func错误处理(llngErr, lstrError)
    End If
    lstrError = "打印水晶文书“" & para文书名称 & "”时失败。原因是：" & lstrError
    
    sfsub错误处理 "文书管理部件", "cls文书集", "funcPrintReport", 6666, lstrError, True
    funcCHPrintReport = False
    Exit Function
    Resume
End Function
'目的：子串替换。
'输入：paraStr      ：被替换的串；
'      paraSubStr   ：被替换的子串；
'      paraNewSubStr：替换为的子串。
'返回：被替换后的串。
Private Function gffuncStrReplace(ByVal paraStr As String, _
                                ByVal paraSubStr As String, _
                                ByVal paraNewSubStr As String) As String
    Dim i As Integer
    Dim intSubStrLen As Integer
    Dim intNewSubStrLen As Integer
    Dim strNew As String
    
    On Error GoTo errhandler
    intSubStrLen = Len(paraSubStr)
    intNewSubStrLen = Len(paraNewSubStr)
    i = 1
    Do While (i + intSubStrLen - 1) <= Len(paraStr)
        If Mid(paraStr, i, intSubStrLen) = paraSubStr Then
            paraStr = Left(paraStr, i - 1) & paraNewSubStr & Right(paraStr, Len(paraStr) - i - intSubStrLen + 1)
            i = i + intNewSubStrLen
        Else
            i = i + 1
        End If
    Loop
    gffuncStrReplace = paraStr
    
    Exit Function
errhandler:
    
End Function

'2012-05-29 陶露
'Private Sub SetDefaultPrinter(ByVal sPrinter As String)
'    Call SaveString(HKEY_CURRENT_USER, sPath, "Device", sPrinter)
'End Sub
'
'Public Function GetString(ByVal lKey As Long, ByVal sPath As String, ByVal sValue As String) As String
'    Dim lKeyHand As Long, lValueType As Long, lResult As Long
'    Dim sBuff As String
'    Dim lDataBufSize As Long
'    Dim iZeroPos As Integer
'
'    Call RegOpenKey(lKey, sPath, lKeyHand)
'    Call RegQueryValueEx(lKeyHand, sValue, 0, lValueType, ByVal 0, lDataBufSize)
'
'    If lValueType = REG_SZ Then
'      sBuff = String$(lDataBufSize, " ")
'      lResult = RegQueryValueEx(lKeyHand, sValue, 0, 0, ByVal sBuff, lDataBufSize)
'
'      If lResult = ERROR_SUCCESS Then
'        iZeroPos = InStr(sBuff, vbNullChar)
'        If iZeroPos > 0 Then
'          GetString = Left$(sBuff, iZeroPos - 1)
'        Else
'          GetString = sBuff
'        End If
'      End If
'    End If
'
'    Call RegCloseKey(lKey)
'End Function
'
'Public Sub SaveString(ByVal lKey As Long, ByVal sPath As String, ByVal sValue As String, ByVal sData As String)
'    Dim lKeyHand As Long
'
'    Call RegCreateKey(lKey, sPath, lKeyHand)
'    Call RegSetValueEx(lKeyHand, sValue, 0, REG_SZ, ByVal sData, Len(sData))
'    Call RegCloseKey(lKeyHand)
'End Sub
'2012-05-29


'2012-05-31 陶露
'根据paraID取回眼科照片
Public Function pmfunc获取眼科照片(paraID As String) As StdPicture
    On Error GoTo errhandler
    Dim lstrSql As String               'SQL语句
    Dim lrecPicture As Object           '根据语句返回眼科照片信息的RecordSet
    Dim lprbPicture As New PropertyBag  '将眼科照片信息进行序列化的属性包
    Dim lvarPicture() As Byte           '眼科照片转换成的Byte数组
    '根据标识号返回相应的眼科照片
    lstrSql = "select * from 职业病体检_体检人员图片结果表 where 系统编号='" & paraID & "'"
    Set lrecPicture = dafuncGetData(lstrSql)
    '从返回的记录集中取出图片信息。
    If lrecPicture.RecordCount = 0 Then
        Err.Raise 7777, , "系统中不存在序号为" & paraID & "的图片！"
    End If
    lvarPicture = lrecPicture("图片").GetChunk(lrecPicture("图片").ActualSize)
    '将图片信息由属性包中读出。
    lprbPicture.Contents = lvarPicture
    '返回根据唯一标识号所取得的图片。
    Set pmfunc获取眼科照片 = lprbPicture.ReadProperty("Picture")
errhandler:
    Set lrecPicture = Nothing
    Set lprbPicture = Nothing
    If Err.Number = 0 Then Exit Function
End Function
'2012-05-31

'2012-06-20 于登淼
'预览体检清单(用于测试程序中)
Public Function func预览体检清单()
    frm打印体检清单简化版.Show 1
End Function

'2012-06-20 于登淼  更改：2012-10-17 罗李奎
'打印体检清单
Public Function func打印体检清单(ByVal paraSysNo As Collection)
    Dim i As Integer
    For i = 1 To paraSysNo.Count
      frm打印体检清单简化版.sysNo = paraSysNo.Item(i)
      frm打印体检清单简化版.subPrint
    Next i
End Function

'2012-06-20 于登淼  更改：2012-10-17 罗李奎
'打印体检清单
Public Function func打印单个体检清单(ByVal paraSysNo As String)
    Dim i As Integer
      frm打印体检清单简化版.sysNo = paraSysNo
      frm打印体检清单简化版.subPrint
End Function

'2012-06-20 于登淼
'打印试管标签
Public Function func打印试管标签(ByVal paraSysNo As String, ByVal paraName As String, ByVal paraSex As String, ByVal paraAge As String, Optional ByVal paraDeptName As String = "")
'Public Function func打印试管标签(ByVal paraSysNo As String, ByVal paraName As String, Optional ByVal paraDeptName As String = "")
    frm打印试管标签.sysNo = paraSysNo
    frm打印试管标签.paraName = paraName
    '加入性别和年龄 2015-12-25 by 牟俊
    frm打印试管标签.paraSex = paraSex
    frm打印试管标签.paraAge = paraAge
    frm打印试管标签.paraDeptName = paraDeptName
    frm打印试管标签.PrintLabel
'    frm打印试管标签.Show
End Function
