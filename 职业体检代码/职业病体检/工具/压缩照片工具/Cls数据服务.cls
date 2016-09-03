VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cls数据服务"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3CF3389101DA"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Option Explicit


'功能：将照片信息生成照片文件
Public Sub sub服务进程()
    On Error GoTo errHandler
    Dim lobjRec As Object       '定义临时记录集
    Dim i As Integer        '定义循环变量
    Dim lstr照片名称 As String  '定义名称记录照片名称
    Dim lstrFullPath As String  '照片全路径。
    
    Dim lpic数据 As IPictureDisp
    
    Dim lobjFileSysobject As Scripting.FileSystemObject
    Dim lobjFile As File
    Dim lobjfolder As Folder
    Dim llng状态 As Long '健康证状态InnerID
    Dim lstrSys As String
    Dim lstdPicture As StdPicture
    
    Dim lobjTransfer As Object
    
    On Error GoTo errHandler
    Set lobjTransfer = CreateObject("BmpToJGP.clsBmpToJPG")
    
    '获取体检记录集信息（已下结论，并且日期在指定期限之前的）。
    lstrSys = "体检管理"
    Set lobjRec = dafuncGetData("select 系统编号 from 体检管理_体检基本信息表 a,系统管理_系统图片管理表 b where a.系统编号=b.图片编号 and b.子系统名='" & lstrSys & "'")
    If lobjRec.RecordCount > 0 Then
        lstrFullPath = App.path & "\tjtmp.bmp"
        
        lobjRec.MoveFirst
        For i = 0 To lobjRec.RecordCount - 1
            Set lpic数据 = pmfunc获取图片(lobjRec!系统编号, lstrSys)
            '生成照片临时文件(.bmp)
            RidFile lstrFullPath
            SavePicture lpic数据, lstrFullPath
            
            '压缩图片。
            lobjTransfer.subSetBMPFile lstrFullPath
            Set lstdPicture = lobjTransfer.funcSaveToJPG()
            
            '保存回库。
            pmsub删除图片 lobjRec!系统编号, lstrSys
            pmsub保存图片 lstdPicture, lobjRec!系统编号, lstrSys
            
            lobjRec.MoveNext
            DoEvents
        Next i
    End If
    
    Set lobjRec = dafuncGetData("select * from 健康证_健康证状态字典视图 where 名称='已发放'")
    llng状态 = lobjRec("InnerID")
    
    '获取健康证记录集信息（已发放，并且领取日期在指定期限之前的）。
    lstrSys = "健康证管理"
    Set lobjRec = dafuncGetData("select 系统编号 from 健康证_健康证信息表 a,系统管理_系统图片管理表 b where a.系统编号=b.图片编号 and b.子系统名='" & lstrSys & "'")
    If lobjRec.RecordCount > 0 Then
        lstrFullPath = App.path & "\jkztmp.bmp"
        
        lobjRec.MoveFirst
        For i = 0 To lobjRec.RecordCount - 1
            Set lpic数据 = pmfunc获取图片(lobjRec!系统编号, lstrSys)
            '生成照片文件(.bmp)
            RidFile lstrFullPath
            SavePicture lpic数据, lstrFullPath
                        
            '压缩图片。
            lobjTransfer.subSetBMPFile lstrFullPath
            Set lstdPicture = lobjTransfer.funcSaveToJPG()
            
            '保存回库。
            pmsub删除图片 lobjRec!系统编号, lstrSys
            pmsub保存图片 lstdPicture, lobjRec!系统编号, lstrSys
            
            lobjRec.MoveNext
            DoEvents
        Next i
    End If

    Exit Sub
errHandler:
    Err.Raise Err.Number, , Err.Description
    Exit Sub
    Resume
End Sub




Public Function FileExists(FileName As String) As Boolean
    If Len(FileName) > 0 Then FileExists = (Len(Dir(FileName, vbNormal Or vbReadOnly Or vbHidden Or vbSystem Or vbArchive)) > 0)
End Function
Public Sub RidFile(FileName As String)
    On Error Resume Next
    If FileExists(FileName) Then
        SetAttr FileName, vbNormal
        Kill FileName
    End If
End Sub

