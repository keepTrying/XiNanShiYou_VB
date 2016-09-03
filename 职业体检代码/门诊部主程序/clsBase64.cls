VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBase64"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TempBucket
    nData(0 To 3) As Byte
    nSize As Byte
End Type

Private aPower2(0 To 31) As Long

Private m_nDBufLen As Long
Private m_nEBufLen As Long
Private m_nDDataLen As Long
Private m_nEDataLen As Long

Private m_DBuffer() As Byte
Private m_EBuffer() As Byte

Private Const Base64Digits = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
Private Const PageSize = 1

Private m_Init As Boolean
Private m_DecodeTable(0 To 255) As String

Private Function RoundToPage(a As Long)
    If PageSize > 1 Then
        RoundToPage = (((a \ PageSize) + 1) * PageSize)
    Else
        RoundToPage = a
    End If
End Function

Public Property Get DecodedMessage() As String
    DecodedMessage = StrConv(m_DBuffer, vbUnicode)
End Property
Public Property Get DecodedMessageB() As Byte()
    DecodedMessageB = m_DBuffer
End Property

Public Property Get EncodedMessage() As String
    EncodedMessage = StrConv(m_EBuffer, vbUnicode)
End Property

Private Sub AllocEncode(nSize As Long)
    If m_nEBufLen < nSize Then
        m_nEBufLen = RoundToPage(nSize)
        ReDim m_EBuffer(m_nEBufLen) As Byte
    End If
    m_nEDataLen = 0
End Sub
Private Sub AllocDecode(nSize As Long)
    If m_nDBufLen < nSize Then
        m_nDBufLen = RoundToPage(nSize)
        ReDim m_DBuffer(m_nDBufLen) As Byte
    End If
    m_nDDataLen = 0
End Sub

Private Function IsBadMimeChar(nData As Byte) As Boolean
    Select Case nData
    '\r, \n, \t, \b, \a, \r, \f, \v
        Case Is <= 32: IsBadMimeChar = True
        Case Else: IsBadMimeChar = False
    End Select
End Function

Private Sub SetEncodeBuffer(pBuffer() As Byte, nBufLen As Long)
    Dim i As Long
    
    AllocEncode nBufLen
    Do While i < nBufLen
        If Not IsBadMimeChar(pBuffer(i)) Then
            m_EBuffer(m_nEDataLen) = pBuffer(i)
            m_nEDataLen = m_nEDataLen + 1
        End If
        i = i + 1
    Loop
End Sub
Private Sub SetDecodeBuffer(pBuffer() As Byte, nBufLen As Long)
    Dim i As Long
    
    AllocDecode nBufLen
    Do While i < nBufLen
        m_DBuffer(i) = pBuffer(i)
        i = i + 1
    Loop
    m_nDDataLen = nBufLen
End Sub

Private Sub ClearBucket(lBucket As TempBucket)
    Dim i
    For i = 0 To 3
        lBucket.nData(i) = 0
    Next i
    lBucket.nSize = 0
End Sub
Private Sub SetBucket(lBucket As TempBucket, pBuffer() As Byte, lIndex As Long, Optional lMax As Byte = 4)
    Dim i As Long
    For i = 0 To lMax - 1
        lBucket.nData(i) = pBuffer(lIndex + i)
    Next i
    lBucket.nSize = lMax
End Sub

Private Sub doEncode(pBuffer() As Byte, nBufLen As Long)
    Dim Raw As TempBucket, nIndex As Long
    
    SetDecodeBuffer pBuffer, nBufLen
    AllocEncode nBufLen * 2
    
    Do Until ((nIndex + 3) > nBufLen)
        ClearBucket Raw
        SetBucket Raw, m_DBuffer, nIndex, 3
        EncodeToBuffer Raw, m_EBuffer, m_nEDataLen
        nIndex = nIndex + 3
        m_nEDataLen = m_nEDataLen + 4
    Loop
    If nBufLen > nIndex Then
        ClearBucket Raw
        SetBucket Raw, m_DBuffer, nIndex, (nBufLen - nIndex)
        EncodeToBuffer Raw, m_EBuffer, m_nEDataLen
        m_nEDataLen = m_nEDataLen + 4
    End If
End Sub
Public Function EncodeB(pBuffer() As Byte) As String
    Dim leng As Long
    
    leng = UBound(pBuffer) - 0 + 1
    If leng > 0 Then
        doEncode pBuffer, leng
        EncodeB = EncodedMessage
    Else
        EncodeB = ""
    End If
End Function
Public Function Encode(strMessage As String) As String
    Dim tmpbuff() As Byte
    Dim strtmp As String
    
    If strMessage = "" Then
        Encode = ""
    Else
        strtmp = StrConv(strMessage, vbFromUnicode)
        'tmpBuff = StrConv(strMessage, vbFromUnicode)
        tmpbuff = strtmp
        doEncode tmpbuff, LenB(strtmp)
        Encode = Trim(EncodedMessage)
    End If
End Function

Private Sub doDecode(pBuffer() As Byte, dwBufLen As Long)
    Dim Raw As TempBucket, nIndex As Long, i As Long
    
    If Not m_Init Then Init
    SetEncodeBuffer pBuffer, dwBufLen
    AllocDecode dwBufLen
    
    Do Until ((nIndex + 4) > m_nEDataLen)
        ClearBucket Raw
        For i = 0 To 3
            Raw.nData(i) = m_DecodeTable(m_EBuffer(nIndex + i))
        Next i
        If Raw.nData(2) = 255 Then Raw.nData(2) = 0
        If Raw.nData(3) = 255 Then Raw.nData(3) = 0
        Raw.nSize = 4
        DecodeToBuffer Raw, m_DBuffer, m_nDDataLen
        nIndex = nIndex + 4
        m_nDDataLen = m_nDDataLen + 3
    Loop
    '// If nIndex < m_nEDataLen, then we got a decode message without padding.
    '// We may want to throw some kind of warning here, but we are still required
    '// to handle the decoding as if it was properly padded.
    
    If nIndex < m_nEDataLen Then
        ClearBucket Raw
        For i = nIndex To m_nEDataLen - 1
            Raw.nData(i - nIndex) = m_DecodeTable(m_EBuffer(i))
            Raw.nSize = Raw.nSize + 1
            If Raw.nData(i - nIndex) = 255 Then Raw.nData(i - nIndex) = 0
        Next i
        DecodeToBuffer Raw, m_DBuffer, m_nDDataLen
        m_nDDataLen = m_nDDataLen + m_nEDataLen - nIndex
    End If
End Sub
Public Function DecodeB(strMessage As String) As Byte()
    Dim tmpbuff() As Byte
    
    tmpbuff = strMessage
    If strMessage = "" Then
        DecodeB = ""
    Else
    'strtmp = StrConv(strMessage, vbFromUnicode)
    'tmpbuff = strtmp
    'tmpBuff = Trim(strMessage)
        doDecode tmpbuff, LenB(strMessage)
        DecodeB = DecodedMessageB
    End If
End Function
Public Function Decode(ByVal strMessage As String) As String
    Dim tmpbuff() As Byte
    Dim strtmp As String
    
    If strMessage = "" Then
        Decode = ""
    Else
        strtmp = StrConv(strMessage, vbFromUnicode)
        tmpbuff = strtmp
        'tmpBuff = Trim(strMessage)
        doDecode tmpbuff, LenB(strtmp)
        Decode = DecodedMessage
    End If
End Function

Private Function DecodeToBuffer(Decode As TempBucket, pBuffer() As Byte, nIndex As Long) As Long
    Dim Data As TempBucket, nCount As Long, i As Long
    
    DecodeRaw Data, Decode
    For i = 0 To 2
        pBuffer(nIndex + i) = Data.nData(i)
        If Not (pBuffer(nIndex + i) = 255) Then nCount = nCount + 1
    Next
    DecodeToBuffer = nCount
End Function
Private Sub EncodeToBuffer(Decode As TempBucket, pBuffer() As Byte, nIndex As Long)
    Dim Data As TempBucket, i As Long
    
    EncodeRaw Data, Decode
    For i = 0 To 3
        pBuffer(nIndex + i) = Asc(Mid(Base64Digits, Data.nData(i) + 1, 1))
    Next i
    If Decode.nSize = 1 Then
        pBuffer(nIndex + 2) = 61 '"="
        pBuffer(nIndex + 3) = 61 '"="
    ElseIf Decode.nSize = 2 Then
        pBuffer(nIndex + 3) = 61 '"="
    End If
End Sub

Private Function LShiftByte(ByVal b As Byte, ByVal c As Integer) As Byte
    Dim w As Integer
    
    w = b * Power2(c)
    If w And &H80& Then
        LShiftByte = CByte(w And &H7F&) Or &H80
    Else
        LShiftByte = w And &HFF&
    End If
End Function

Function RShiftByte(ByVal b As Byte, ByVal c As Integer) As Byte
    Dim w As Integer
    
    If c = 0 Then
        RShiftByte = b
    Else
        w = b And &HFF&
        w = w \ Power2(c)
        RShiftByte = w And &HFF&
    End If
End Function

Property Get Power2(ByVal i As Integer) As Long
    If aPower2(0) = 0 Then
        aPower2(0) = &H1&
        aPower2(1) = &H2&
        aPower2(2) = &H4&
        aPower2(3) = &H8&
        aPower2(4) = &H10&
        aPower2(5) = &H20&
        aPower2(6) = &H40&
        aPower2(7) = &H80&
        aPower2(8) = &H100&
        aPower2(9) = &H200&
        aPower2(10) = &H400&
        aPower2(11) = &H800&
        aPower2(12) = &H1000&
        aPower2(13) = &H2000&
        aPower2(14) = &H4000&
        aPower2(15) = &H8000&
        aPower2(16) = &H10000
        aPower2(17) = &H20000
        aPower2(18) = &H40000
        aPower2(19) = &H80000
        aPower2(20) = &H100000
        aPower2(21) = &H200000
        aPower2(22) = &H400000
        aPower2(23) = &H800000
        aPower2(24) = &H1000000
        aPower2(25) = &H2000000
        aPower2(26) = &H4000000
        aPower2(27) = &H8000000
        aPower2(28) = &H10000000
        aPower2(29) = &H20000000
        aPower2(30) = &H40000000
        aPower2(31) = &H80000000
    End If
    Power2 = aPower2(i)
End Property

Private Sub DecodeRaw(Data As TempBucket, Decode As TempBucket)
    Dim nTemp As Byte
    Data.nData(0) = Decode.nData(0)
    Data.nData(0) = LShiftByte(Data.nData(0), 2)
    
    nTemp = Decode.nData(1)
    nTemp = RShiftByte(nTemp, 4) And &H3
    Data.nData(0) = Data.nData(0) Or nTemp
    
    Data.nData(1) = Decode.nData(1)
    Data.nData(1) = LShiftByte(Data.nData(1), 4)
    
    nTemp = Decode.nData(2)
    nTemp = RShiftByte(nTemp, 2) And &HF
    Data.nData(1) = Data.nData(1) Or nTemp
    
    Data.nData(2) = Decode.nData(2)
    Data.nData(2) = LShiftByte(Data.nData(2), 6)
    
    nTemp = Decode.nData(3) And &H3F
    Data.nData(2) = Data.nData(2) Or nTemp
End Sub

Private Sub EncodeRaw(Data As TempBucket, Decode As TempBucket)
    Dim nTemp As Byte
    
    Data.nData(0) = Decode.nData(0)
    Data.nData(0) = RShiftByte(Data.nData(0), 2)
    
    Data.nData(1) = Decode.nData(0)
    Data.nData(1) = LShiftByte(Data.nData(1), 4)
    nTemp = Decode.nData(1)
    nTemp = RShiftByte(nTemp, 4)
    Data.nData(1) = Data.nData(1) Or nTemp
    Data.nData(1) = Data.nData(1) And &H3F
    
    Data.nData(2) = Decode.nData(1)
    Data.nData(2) = LShiftByte(Data.nData(2), 2)
    
    nTemp = Decode.nData(2)
    nTemp = RShiftByte(nTemp, 6)
    
    Data.nData(2) = Data.nData(2) Or nTemp
    Data.nData(2) = Data.nData(2) And &H3F
    
    Data.nData(3) = Decode.nData(2)
    Data.nData(3) = Data.nData(3) And &H3F
End Sub

Private Sub Init()
    Dim i As Long
    
    For i = 0 To 255
        m_DecodeTable(i) = &HFE '-2
    Next i
    For i = 0 To 63
        m_DecodeTable(Asc(Mid(Base64Digits, i + 1, 1))) = i
        m_DecodeTable(Asc(Mid(Base64Digits, i + 1, 1)) Or &H80) = i
    Next i
    m_DecodeTable(61) = &HFF '61 = "="
    m_DecodeTable(189) = &HFF '189 = 61 Or &H80
    
    m_Init = True
End Sub


