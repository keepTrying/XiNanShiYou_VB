VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBmpToJPG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_Image     As New cImage
Private m_Jpeg      As cJpeg


'设置要转换的bmp文件名。
Public Sub subSetBMPFile(ByVal paraFile As String)
    Dim lstdPic As StdPicture
    Set lstdPic = LoadPicture(paraFile)
    Set m_Image = New cImage
    m_Image.CopyStdPicture lstdPic
    
End Sub

Public Sub subSetBMP(ByVal paraPicture As StdPicture)
    Set m_Image = New cImage
    m_Image.CopyStdPicture paraPicture
End Sub

'把bmp图片转换为jgp文件。
Public Sub subSaveToJPGFile(ByVal paraFileName As String, Optional paraQuality As Long = 75)
    Set m_Jpeg = New cJpeg
    m_Jpeg.SetSamplingFrequencies 2, 2, 1, 1, 1, 1
    m_Jpeg.SampleHDC m_Image.hDC, m_Image.Width, m_Image.Height
    m_Jpeg.Quality = 75
    If paraQuality <> 75 Then
        m_Jpeg.Quality = paraQuality
    End If
    
    'Delete file if it exists
    RidFile paraFileName

    'Save the JPG file
    m_Jpeg.SaveFile paraFileName

End Sub


'把bmp图片转换为jgp文件。
Public Function funcSaveToJPG(Optional paraQuality As Long = 75) As StdPicture
    Dim lstrTmpFileName As String
    Set m_Jpeg = New cJpeg
    m_Jpeg.SetSamplingFrequencies 2, 2, 1, 1, 1, 1
    m_Jpeg.SampleHDC m_Image.hDC, m_Image.Width, m_Image.Height
    m_Jpeg.Quality = 75
    If paraQuality <> 75 Then
        m_Jpeg.Quality = paraQuality 'paraQuality
    End If
    lstrTmpFileName = App.Path & "\tmpbmptojgp.jpg"
    
    'Delete file if it exists
    RidFile lstrTmpFileName

    'Save the JPG file
    m_Jpeg.SaveFile lstrTmpFileName
    
    Set funcSaveToJPG = LoadPicture(lstrTmpFileName)
    
End Function

