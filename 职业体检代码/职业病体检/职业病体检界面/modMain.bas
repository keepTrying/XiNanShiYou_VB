Attribute VB_Name = "modMain"

'���ƣ�ְҵ��ʷְҵ����������
'������
'���ܣ�ˢ�������֤��������
'      ������
'���ߣ�Yunle Liu
'ʱ�䣺2012.03
Public mstrϵͳ��� As String  '2015-10-22
Public InputFlag As String
Public InputFlagNo As String
Option Explicit
'*********************************************************************
'�������֤������ ��������
'Public Declare Function CVR_InitComm Lib "termb.dll" (ByVal Port As Long) As Integer
'Public Declare Function CVR_CloseComm Lib "termb.dll" () As Integer
'Public Declare Function CVR_Authenticate Lib "termb.dll" () As Integer
'Public Declare Function CVR_Read_Content Lib "termb.dll" (ByVal Active As Long) As Integer
'Public Declare Function CVR_Ant Lib "termb.dll" (ByVal mode As Long) As Integer
'
'Public Declare Function GetPeopleName Lib "termb.dll" (ByVal lpReturnedString As String, ByRef nReturnLen As Integer) As Integer
'Public Declare Function GetPeopleAddress Lib "termb.dll" (ByVal lpReturnedString As String, ByRef nReturnLen As Integer) As Integer
'Public Declare Function GetPeopleIDCode Lib "termb.dll" (ByVal lpReturnedString As String, ByRef nReturnLen As Integer) As Integer
'Public Declare Function GetPeopleNation Lib "termb.dll" (ByVal lpReturnedString As String, ByRef nReturnLen As Integer) As Integer
Declare Function InitComm Lib "Sdtapi.dll" (ByVal iPort As Integer) As Integer

Declare Function Authenticate Lib "Sdtapi.dll" () As Integer

Declare Function ReadBaseInfos Lib "Sdtapi.dll" (ByVal iname As String, ByVal isex As String, ByVal folk As String, ByVal birthday As String, ByVal code As String, ByVal addr As String, ByVal agency As String, ByVal startdate As String, ByVal enddate As String) As Integer

Declare Function CloseComm Lib "Sdtapi.dll" () As Integer

Declare Function ReadBaseMsgW Lib "Sdtapi.dll" (ByVal pMsg As String, ByRef LenT As Integer) As Integer


Declare Function ReadBaseMsg Lib "Sdtapi.dll" (ByVal pMsg As String, ByRef LenT As Integer) As Integer
Declare Function ReadIINSNDN Lib "Sdtapi.dll" (ByVal pIINSNDN As String) As Integer
Declare Function GetSAMIDToStr Lib "Sdtapi.dll" (ByVal pcSAMID As String) As Integer
Global Comm As Boolean
Public Declare Function SendMessage Lib "user32" _
            Alias "SendMessageA" (ByVal hwnd As Long, _
            ByVal wMsg As Long, ByVal wParam As Long, _
            lParam As Any) As Long
'**********************************************************************

Public pobjDict As Object
Public pstrFilename As String
Public pstrWordname As String
Public pobjFileToDatabase As Object
Public pstr����վ���� As String
Public pstrPhoto As String

Public ���ʱ�־ As Integer
Public pobjҵ����� As Object '������ҵ�����clsManageMedicalExam��

Public mstrQuery As String




Public Sub Main()
    On Error Resume Next
     '�����ֵ����
    Set pobjDict = CreateObject("�ֵ����.clsDictionary")
    Err.Clear
    '����ҵ�����
    Set pobjҵ����� = CreateObject("ְҵ������.clsManageMedicalExam")
    If Err <> 0 Then
        On Error GoTo errHandler
        Err.Raise 6666, , "�޷�����ְҵ��ʷ¼����档������ע�ᡰְҵ��ʷ¼�����.dll����"
    End If
   
    
    
    Dim lstrServer As String
    Dim lstrData As String
    lstrServer = sffuncGetSetting("ϵͳ����", "���ݿ�����", "��������")
    lstrData = sffuncGetSetting("ϵͳ����", "���ݿ�����", "���ݿ���")
'    lstrServer = "KAMA-AA251EA62C"
'    lstrData = "BJB-SJK2012"
    
     '����д�ļ�����
    Set pobjFileToDatabase = CreateObject("FileToDatabase.clsFileToDatabase")
    '��⽨�����ӡ�
    With pobjFileToDatabase
        .pstrConnectString = "Provider=MSDataShape.1;Data Provider=SQLOLEDB.1;Password=welcome;Persist Security Info=True;User ID=user26;Initial Catalog=" & lstrData & ";Data Source=" & lstrServer
        .subConnect
    End With
  Exit Sub
errHandler:
    sfsub������ "ְҵ������", "modmain", "Main", Err.Number, Err.Description, False
End Sub

'���ܣ���¼������ݡ�
'���ߣ����
Public Sub sub��¼���ֵ(ByVal para¼��� As Control, _
                        ByVal paraGUI As cls����ͨ�ö���, _
                        ByVal paraInfo As Collection)
    Dim lstrItem As String
    Dim lstrItemText  As String
    Dim i As Integer
    Dim lint�������� As Integer
    Dim lint��ҵ��� As Integer
    Dim j As Integer
    
    On Error GoTo errHandler
    
    
    para¼���.pblnTemp = True
    lint�������� = 0
    
    For i = 1 To para¼���.InfoCollection.Count
        '¼����Ŀ���ơ�
        lstrItem = para¼���.InfoCollection(i).Title
        
        If sffunc�жϼ��ϼ�ֵ�Ƿ����(paraInfo, lstrItem) Then
            '����TrueText��
            para¼���.ItemTrueText(i - 1) = paraInfo(lstrItem)("��Ŀֵ���")
            '����Text��
            para¼���.ItemText(i - 1) = paraInfo(para¼���.InfoCollection(i).Title)("��Ŀֵ")
            
            If lstrItem = "��������" Then
                lint�������� = i
            ElseIf lstrItem = "��ҵ���" Then
                lint��ҵ��� = i
            End If
        Else
            para¼���.ItemTrueText(i - 1) = ""
            para¼���.ItemText(i - 1) = ""
        End If
    Next i
    
    Dim lobjRec As Object
    Dim lstrItemTrueText As String
    '������ҵ���¼�����ֵ����ݵ�������
    If lint�������� > 0 And lint��ҵ��� > 0 Then

        '��ȡ���������š�
        lstrItemTrueText = para¼���.ItemTrueText(lint�������� - 1)

        '������ҵ���¼�����ֵ䡣
        If lstrItemTrueText <> "" And Not para¼���.InfoCollection(lint��������).DictRecordSet Is Nothing Then
            Set lobjRec = para¼���.InfoCollection(lint��������).DictRecordSet
            If Not lobjRec.EOF Then
                paraGUI.sub��ʼ���ֵ�� lint��ҵ���, "Parent=" & lobjRec("InnerId")
            End If
        End If
    End If
  
    para¼���.pblnTemp = False
    Exit Sub
errHandler:
    para¼���.pblnTemp = False
    sfsub������ "ְҵ�����沿��", "modMain", "sub��¼���ֵ", Err.Number, Err.Description, True
    Exit Sub
    Resume
End Sub

Public Sub sub��ʾ��λ����(ByVal ciptBase As Control, _
            ByVal para��λ������ As String, _
            ByVal paraGUI As cls����ͨ�ö���)
    Dim i As Long
    Dim lcolInfo As Collection
    
    If para��λ������ <> "" Then
    
        '��ȡ��λ���ԡ�
        On Error Resume Next
        '��ȡ��λ���ԡ�
        Set lcolInfo = pobjҵ�����.func��ȡ��λ����(para��λ������)
        
        ciptBase.pblnTemp = True
        
        ciptBase.Box1("��������").TrueText = ""
        ciptBase.Box1("��ҵ���").TrueText = ""
        ciptBase.Box1("Ƭ��").TrueText = ""
        ciptBase.Box1("��������").TrueText = ""
        
        ciptBase.Box1("��������").TrueText = lcolInfo("��������")
        ciptBase.Box1("��ҵ���").TrueText = lcolInfo("��ҵ���")
        ciptBase.Box1("Ƭ��").TrueText = lcolInfo("Ƭ��")
        ciptBase.Box1("��������").TrueText = lcolInfo("��������")
        
        ciptBase.Box1("��������").Text = lcolInfo("������������")
        ciptBase.Box1("��ҵ���").Text = lcolInfo("��ҵ�������")
        ciptBase.Box1("Ƭ��").Text = lcolInfo("Ƭ������")
        ciptBase.Box1("��������").Text = lcolInfo("������������")
        ciptBase.Box1("��λ��ַ").Text = lcolInfo("��λ��ַ")
        
        
        
        Dim lstrItem As String
        Dim lint�������� As Integer
        Dim lint��ҵ���  As Integer
        
        Err.Clear
        
        '�ж��Ƿ����������ࡣ
        For i = 1 To ciptBase.InfoCollection.Count
            '¼����Ŀ���ơ�
            lstrItem = ciptBase.InfoCollection(i).Title
            
            If lstrItem = "��������" Then
                lint�������� = i
            ElseIf lstrItem = "��ҵ���" Then
                lint��ҵ��� = i
            End If
            If Err <> 0 Then Exit For
        Next i
        
        '������ҵ���¼�����ֵ����ݵ�������
        Dim lstrItemTrueText As String
        Dim lobjRec As Object
        If lint�������� > 0 And lint��ҵ��� > 0 Then
            '��ȡ���������š�
            lstrItemTrueText = ciptBase.ItemTrueText(lint�������� - 1)
            '������ҵ���¼�����ֵ䡣
            If lstrItemTrueText <> "" And Not ciptBase.InfoCollection(lint��������).DictRecordSet Is Nothing Then
                Set lobjRec = ciptBase.InfoCollection(lint��������).DictRecordSet
                If Not lobjRec.EOF Then
                    paraGUI.sub��ʼ���ֵ�� lint��ҵ���, "Parent=" & lobjRec("InnerId")
                End If
            End If
        End If
        
        ciptBase.pblnTemp = False
    End If

End Sub


'------��ӡ���鲿����������Ҫ��������������
'------(ϵͳԭ��Ҫ���ǣ�û�и�����Ϣ��������ӡ��������ֻ�Ǵ�����)
'------����ӡ������ʲô����?����������ڻ���֪��ʡ�����Ǳ�ϣ����ӡ��ɶ��~
Public Function sub��ӡ������������(ByVal para�������� As String)
    Dim lobjTmp As Object
    Set lobjTmp = CreateObject("ְҵ������.cls����")
    lobjTmp.sub��ӡ���� "����ӡ����", para��������, False
End Function

'2012-04-05 ��¶
'��ӡ�����������
Public Function sub��ӡ�����������(ByRef para�������� As Collection)
    Dim lobjTmp As Object
    Set lobjTmp = CreateObject("ְҵ������.cls����")
    lobjTmp.Sub��ӡ������� "����ӡ����", Nothing, False, False, para��������
End Function
'2012-04-05 ��¶

'2012-08-20 �ڵ��
'�༭word�ĵ�������
Public Function sub�༭word�ĵ�(paraParent As Object, ByVal paraSysNo As String, ByVal mstr��������, ByVal paraReadOnly As Boolean)
    Dim objWord As Object                      'Word.Application
    Dim objWordDocument As Object       'Word.Document
    Dim lobjRec As Object, lobjRec1 As Object
    Dim lstrType As String, lstrTypeNo As String
    Dim lintRepID As Integer                    '����Word�����ID
    Dim lstr������� As String
    
    On Error GoTo errHandler
    
    '����word
    On Error Resume Next
    MkDir App.Path & "\temp"
    Kill App.Path & "\temp\*.*"
    
    Const CLASSOBJECT = "Word.Application"
    Set objWord = GetObject(, CLASSOBJECT)
    
    Err.Clear
    If (objWord Is Nothing) Then
        Set objWord = CreateObject(CLASSOBJECT)
        If Err.Number <> 0 Then
            Err.Raise 6666, , "��û�а�װWord���޷��༭���棡���Ȱ�װMS Office 2000���ϰ汾��"
             Unload frmPrintPaper
        End If
    End If
    
    On Error GoTo errHandler
    
    Dim lstrNewDoc As String
    Dim lstrDotFile As String           'ģ���ļ�
    Dim i As Integer, j As Integer
    
    
    '������顣
    'ȡ��ģ���ļ�
    '�޸��ˣ������ ʱ�䣺2013-1-8 ��
    '˵��������������ȡ����Ҫ��wordģ��
'    If mstr�������� = "��챨��" Then
'         '�ж������Ƿ��Ѵ��ڡ�
'    Set lobjRec = dafuncGetData("select ���,�ļ����� from ְҵ�����_��챨����Ϣ�� where ������='" & paraSysNo & "'")
'    If lobjRec.RecordCount = 0 Then
''        If paraReadOnly Then        '�������޸ģ�����Ϊ�鿴������������������
'            Err.Raise 6666, , "����Ʒû��¼��Word���棬����ʱ������Ϊ�����Word���棡", "ϵͳ��ʾ"
''        End If
'    End If
'         pstrWordname = para��������
'         mstr�������� = ""
'    Else
'        pstrWordname = frmFinalConclusion.cgrdInfo.TextMatrix(frmFinalConclusion.cgrdInfo.SelectedRow(0), 7)
'    End If
'     pstrWordname = frmFinalConclusion.cgrdInfo.TextMatrix(frmFinalConclusion.cgrdInfo.SelectedRow(0), 7)
        pstrWordname = mstr��������

    If pstrWordname = "" Then
        Err.Raise 6666, , "�������Ϊ��"
         Unload frmPrintPaper
    End If
    subȡ��wordģ��
    '�޸��ˣ������ ʱ�䣺2013-1-8 ��
    
'     frmѡ��Wordģ��.pstrWordname = frmFinalConclusion.cgrdInfo.TextMatrix(frmFinalConclusion.cgrdInfo.SelectedRow(0), 6)
'    frmѡ��Wordģ��.Show 1
'    lstrDotFile = frmѡ��Wordģ��.pstrFilename

    lstrDotFile = pstrFilename
    If lstrDotFile = "" Then Exit Function
    lstrNewDoc = App.Path & "\temp\" & paraSysNo & "_" & Format(Now, "yyyy-mm-dd") & ".doc"
    '��ģ�壬�������ĵ�����ʱΪ��ʱ�ļ���
    Set objWordDocument = objWord.Documents.Open(FileName:=App.Path & "\" & lstrDotFile, ReadOnly:=False)
    objWordDocument.ActiveWindow.Caption = lstrNewDoc
    
    
    '��ȡ��ǩ���ݡ�
    If Right(lstrDotFile, 4) = ".dot" Then
        lstrDotFile = Left(lstrDotFile, Len(lstrDotFile) - 4)
    ElseIf Right(lstrDotFile, 4) = "dotx" Then
        lstrDotFile = Left(lstrDotFile, Len(lstrDotFile) - 5)
    End If
    'Set lobjrec1 = dafuncGetData("exec ְҵ�����_��ȡword����_�հױ��� '" & paraSysNo & "','" & um�û���� & "','" & lstrDotFile & "'")
    dasubSetQueryTimeout 600
    Set lobjRec1 = dafuncGetData("select * from  ְҵ�����_�����Ա������Ϣ�� where ϵͳ���='" & paraSysNo & "'")

    'objWord.Visible = True
    
    '�����ǩ����
    Dim lstrValue As String
    Dim myRange As Object, myTable As Object

    If lobjRec1.RecordCount > 0 Then
        Set myRange = objWordDocument.Content

        '����������������2012-10-25 �����
        If objWordDocument.Tables.Count > 1 Then
            If InStr(lstrDotFile, "���乤����Ա") > 0 Then                           '���乤����
                sub���ط��乤����Ա����word��Ϣ objWordDocument, myRange, paraSysNo
            ElseIf InStr(lstrDotFile, "����Թ�����Ա") > 0 Then                       '��˹�����
                sub������˹�����Աְҵ����word��Ϣ objWordDocument, myRange, paraSysNo
            ElseIf InStr(lstrDotFile, "ְҵ����") > 0 Then         'ְҵ����
                sub����ְҵ����word��Ϣ objWordDocument, myRange, paraSysNo
            ElseIf InStr(lstrDotFile, "��ͨ�Թ�����Ա") > 0 Then          '����ͨ����
                sub������ͨ�Թ�����Ա����word��Ϣ objWordDocument, myRange, paraSysNo
            ElseIf InStr(lstrDotFile, "8023") Or InStr(lstrDotFile, "�����Թ�����Ա") > 0 Then                       '8023�ͷ����Թ�����
                sub����8023�ͷ����Թ�����Աword��Ϣ objWordDocument, myRange, paraSysNo
            End If
            
        End If

        '�����ĵ����ģ�������ҳü��ҳ�ţ��е��������Ҫ��ҳ�롢ҳ��
        For i = 1 To objWordDocument.Range.Fields.Count
            objWordDocument.Range.Fields(i).Update
        Next
        '�����ļ�
        objWordDocument.SaveAs lstrNewDoc
        objWordDocument.Saved = False
    End If

    With objWord.ActiveWindow.View
        .ShowRevisionsAndComments = False
        .RevisionsView = 0          'wdRevisionsViewFinal
    End With

    objWordDocument.ActiveWindow.ActivePane.View.SeekView = 9
    objWordDocument.ActiveWindow.ActivePane.View.SeekView = 0

    objWord.Visible = True

    On Error GoTo errHandler

    objWordDocument.Activate
    objWord.Activate
     
    '�����  2013��1��10��   ��
        If Not paraReadOnly Then
            If lintRepID = 0 Then
                 objWord.Run "subStart", paraParent, -1, paraSysNo
            Else
                objWord.Run "subStart", paraParent, lintRepID, paraSysNo
            End If
          End If
     '�����  2013��1��10��   ��
     
    If paraReadOnly Then
        If objWordDocument.Range.Fields.Count = 0 Then objWordDocument.Protect 3, , "sccdc789"
        objWordDocument.Saved = True
    End If
    Exit Function
    
errHandler:
    If Err = 3001 Then
        MsgBox "û�������ݿ����ҵ���Word����ľ����ļ��������Ǳ���ñ��浽ϵͳ��ʱ���緢�����ϣ�����ϵͳ��ɾ���ñ�����Ϣ������¼��ñ��档", vbInformation, "ϵͳ��ʾ"
    Else
        sfsub������ "���沿��", "mod�������", "sub�༭word�ĵ�", Err.Number, Err.Description, True
    End If
    Exit Function
    Resume
End Function

'2012-09-14 ����
'��ӡ��λ����
Sub sub�༭��λ����(paralobjRec As Object, paralcol As Collection)
    Dim objWord As Object                      'Word.Application
    Dim objWordDocument As Object       'Word.Document
    Dim lobjRec As Object, lobjRec1 As Object
    Dim lstrType As String, lstrTypeNo As String
    Dim lintRepID As Integer                    '����Word�����ID
    '�����ǩ����
    Dim lstrValue As String
    Dim myRange As Object, myTable As Object
    Dim lstrNewDoc As String
    Dim lstrDotFile As String           'ģ���ļ�
    Dim i As Integer, j As Integer
    
    On Error GoTo errHandler
    
    '����word
    On Error Resume Next
    MkDir App.Path & "\temp"
    Kill App.Path & "\temp\*.*"
    
    Const CLASSOBJECT = "Word.Application"
    Set objWord = GetObject(, CLASSOBJECT)
    
    Err.Clear
    If (objWord Is Nothing) Then
        Set objWord = CreateObject(CLASSOBJECT)
        If Err.Number <> 0 Then
            Err.Raise 6666, , "��û�а�װWord���޷��༭���棡���Ȱ�װMS Office 2000���ϰ汾��"
        End If
    End If

    '������顣
    'ȡ��ģ���ļ�
    lstrDotFile = "ְҵ�����_��λ����.dot"
    If lstrDotFile = "" Then Exit Sub
    lstrNewDoc = App.Path & "\temp\" & paralcol("�������") & "_" & Format(Now, "yyyy-mm-dd") & ".doc"
    '��ģ�壬�������ĵ�����ʱΪ��ʱ�ļ���
    Set objWordDocument = objWord.Documents.Open(FileName:=App.Path & "\" & lstrDotFile, ReadOnly:=False)
    objWordDocument.ActiveWindow.Caption = lstrNewDoc
    
    '��ȡ��ǩ���ݡ�
    If Right(lstrDotFile, 4) = ".dot" Then
        lstrDotFile = Left(lstrDotFile, Len(lstrDotFile) - 4)
    ElseIf Right(lstrDotFile, 4) = "dotx" Then
        lstrDotFile = Left(lstrDotFile, Len(lstrDotFile) - 5)
    End If
    
    If paralobjRec.RecordCount > 0 Then
        Set myRange = objWordDocument.Content
        
            '�����ͷ����һ�����ݡ�
            objWordDocument.Sections(1).Range.Find.Execute FindText:="����λ���ơ�", ReplaceWith:=paralcol("��λ����"), Replace:=2
            objWordDocument.Sections(1).Range.Find.Execute FindText:="����λ��ַ��", ReplaceWith:=paralcol("��λ��ַ"), Replace:=2
            objWordDocument.Sections(1).Range.Find.Execute FindText:="���ļ���š�", ReplaceWith:=paralcol("�������") & Format(Now, "yyyymmdd"), Replace:=2

            '�������ģ��ڶ�������
            Set myRange = objWordDocument.Range(myRange.Start, myRange.StoryLength)
            Set myTable = myRange.Tables(2)
            If myTable.rows.Count < paralobjRec.RecordCount + 1 Then
                j = myTable.rows.Count
                myTable.rows(j).Select
                objWordDocument.ActiveWindow.Selection.InsertRows paralobjRec.RecordCount - j + 1
                For i = 1 To paralobjRec.RecordCount - myTable.rows.Count + 1
                    myTable.rows.Add (myTable.rows(j))
                Next
            End If
            For i = 1 To paralobjRec.RecordCount
                For j = 1 To paralobjRec.Fields.Count
                    If j = paralobjRec.Fields.Count Then
                        myTable.rows(i + 1).Cells(j).Range.Text = IIf(IsNull(paralobjRec(j - 1)), "", Format(paralobjRec(j - 1), "yyyy-mm-dd"))
                    Else
                        myTable.rows(i + 1).Cells(j).Range.Text = IIf(IsNull(paralobjRec(j - 1)), "", paralobjRec(j - 1))
                    End If
                    
                Next
                paralobjRec.MoveNext
            Next
            
        '�����ĵ����ģ�������ҳü��ҳ�ţ��е��������Ҫ��ҳ�롢ҳ��
        For i = 1 To objWordDocument.Range.Fields.Count
            objWordDocument.Range.Fields(i).Update
        Next
        '�����ļ�
        objWordDocument.SaveAs lstrNewDoc
        objWordDocument.Saved = False
    End If

    With objWord.ActiveWindow.View
        .ShowRevisionsAndComments = False
        .RevisionsView = 0          'wdRevisionsViewFinal
    End With

    objWordDocument.ActiveWindow.ActivePane.View.SeekView = 9
    objWordDocument.ActiveWindow.ActivePane.View.SeekView = 0

    objWord.Visible = True

    On Error GoTo errHandler

    objWordDocument.Activate
    objWord.Activate

    Exit Sub
    
errHandler:
    If Err = 3001 Then
        MsgBox "û�������ݿ����ҵ���Word����ľ����ļ��������Ǳ���ñ��浽ϵͳ��ʱ���緢�����ϣ�����ϵͳ��ɾ���ñ�����Ϣ������¼��ñ��档", vbInformation, "ϵͳ��ʾ"
    Else
        sfsub������ "���沿��", "mod�������", "sub�༭��λ����", Err.Number, Err.Description, True
    End If
    Exit Sub
    Resume
End Sub

'2012-09-22 ����
'2012-09-22 ����
'���ܣ������λ�ܼ챨��
Sub sub�༭�ܼ챨��(lcolFactor As Collection)
    Dim objWord As Object                      'Word.Application
    Dim objWordDocument As Object       'Word.Document
    Dim lobjRec As Object, lobjRec1 As Object
    Dim lstrType As String, lstrTypeNo As String
    Dim lintRepID As Integer                    '����Word�����ID
    '�����ǩ����
    Dim lstrValue As String
    Dim myRange As Object, myTable As Object
    Dim lstrNewDoc As String
    Dim lstrDotFile As String           'ģ���ļ�
    Dim i As Integer, j As Integer
    Dim lcolInfo As Collection, lcolInfo2 As Collection
    On Error GoTo errHandler
    
    '����word
    On Error Resume Next
    MkDir App.Path & "\temp"
    Kill App.Path & "\temp\*.*"
    
    Const CLASSOBJECT = "Word.Application"
    Set objWord = GetObject(, CLASSOBJECT)
    
    Err.Clear
    If (objWord Is Nothing) Then
        Set objWord = CreateObject(CLASSOBJECT)
        If Err.Number <> 0 Then
            Err.Raise 6666, , "��û�а�װWord���޷��༭���棡���Ȱ�װMS Office 2000���ϰ汾��"
        End If
    End If

    '������顣
    'ȡ��ģ���ļ�
    lstrDotFile = "��λ��˾ְҵ������챨��.dot"    '2015-10-28
    
'    lstrDotFile = "��˾ְҵ������챨��.dot"
    If lstrDotFile = "" Then Exit Sub
    lstrNewDoc = App.Path & "\temp\" & Format(Now, "yyyymmddss") & ".doc"
    '��ģ�壬�������ĵ�����ʱΪ��ʱ�ļ���
    Set objWordDocument = objWord.Documents.Open(FileName:=App.Path & "\" & lstrDotFile, ReadOnly:=False)
    objWordDocument.ActiveWindow.Caption = lstrNewDoc
    
    '��ȡ��ǩ���ݡ�
    If Right(lstrDotFile, 4) = ".dot" Then
        lstrDotFile = Left(lstrDotFile, Len(lstrDotFile) - 4)
    ElseIf Right(lstrDotFile, 4) = "dotx" Then
        lstrDotFile = Left(lstrDotFile, Len(lstrDotFile) - 5)
    End If
    
'    If paralobjRec.RecordCount > 0 Then
        Set myRange = objWordDocument.Content
'        Set lcolInfo = lcolFactor("Σ������")
'        Set lcolInfo2 = lcolFactor("Σ�����")
        
            '�����������ڸ��ڼ�ȣ�  2015-11-2��
            Dim TlobjRec As Object
            Dim Tlstr As String
            Dim Testtype As String
            Set TlobjRec = dafuncGetData("select ������ from ְҵ�����_���������ݿ� where  ��λ���� = '" & lcolFactor("��λ����") & "' and (������� >= '" & lcolFactor("��ʼ����") & "' and ������� <= '" & lcolFactor("��ֹ����") & "') group by ������")
'            Testtype = ""
            For i = 1 To TlobjRec.RecordCount
            Testtype = TlobjRec("������")
             lcolFactor.Add Testtype, "������" & i
            Testtype = Testtype & "��"
'            Testtype = Testtype & lcolFactor("������" & i) & "��"
            TlobjRec.MoveNext

            Next
            Testtype = Left(Testtype, Len(Testtype) - 1)
            objWordDocument.Sections(1).Range.Find.Execute FindText:="���������", ReplaceWith:=Testtype, Replace:=2
            '2015-11-2��
        
        
        
            '�����ͷ����һ�����ݡ�
            objWordDocument.Sections(1).Range.Find.Execute FindText:="����λ���ơ�", ReplaceWith:=lcolFactor("��λ����"), Replace:=2
            objWordDocument.Sections(1).Range.Find.Execute FindText:="��������ڡ�", ReplaceWith:=lcolFactor("�������"), Replace:=2
            objWordDocument.Sections(1).Range.Find.Execute FindText:="������ʱ�䡿", ReplaceWith:=Format(Now, "yyyy��mm��dd��"), Replace:=2
            objWordDocument.Sections(1).Range.Find.Execute FindText:="��Ӧ��������", ReplaceWith:=lcolFactor("Ӧ������"), Replace:=2       '����Ӧ������  2015-10-29
            objWordDocument.Sections(1).Range.Find.Execute FindText:="��ʵ��������", ReplaceWith:=lcolFactor("ʵ������"), Replace:=2
            
            Dim lcolItem As Collection
            Dim lstr As String, ltemp As String, lstr2 As String, lsql As String
            Set lcolItem = lcolFactor("�����Ŀ")
            For i = 1 To lcolItem.Count
                lstr = lstr & "��" & lcolItem(i)
            Next

            objWordDocument.Sections(1).Range.Find.Execute FindText:="�������", ReplaceWith:=lstr, Replace:=2
            lstr2 = "����"
            dasubSetQueryTimeout 600
  
                'ȷ��Σ�������м���  2015-10-29
            Dim KlobjRec As Object
            Dim lstr3 As String
            Dim lstr4 As String
            Set KlobjRec = dafuncGetData("select Σ������ from ְҵ�����_���������ݿ� where  ��λ���� = '" & lcolFactor("��λ����") & "' and (������� >= '" & lcolFactor("��ʼ����") & "' and ������� <= '" & lcolFactor("��ֹ����") & "') group by Σ������")

            For i = 1 To KlobjRec.RecordCount
'            For i = 1 To 3
'                    FrmQueryCompany.BigNum (i)    '��Сд���������ָĳɴ�д  2015-11-4)
                '����ÿ��������صĲ��ϸ���Ա
                    Set lobjRec = dafuncGetData("select b.����,count(*) ���� from dbo.ְҵ�����_�������ͼ a,ְҵ�����_�����Ŀ���ñ� b where a.�����Ŀ=b.���� and ϵͳ��� in(" _
                        & " select ϵͳ��� from ְҵ�����_���������ݿ� where Σ������ = '" & lcolFactor("Σ������" & i) & "'and ��λ���� = '" & lcolFactor("��λ����") & "'" _
                        & " and (������� >= '" & lcolFactor("��ʼ����") & "' and ������� <= '" & lcolFactor("��ֹ����") & "') and ���״̬ in( 7,5)" _
                        & ") and ������� = '���ϸ�' group by b.����")
'                    If lobjRec Is Nothing Then Exit Sub
                    If lcolFactor("Σ������" & i) <> "" Then
                        If Not lobjRec Is Nothing Then
                        While Not lobjRec.EOF
                            ltemp = ltemp & lobjRec("����") & "���ϸ�" & lobjRec("����") & "�ˣ�"
                            lobjRec.MoveNext
                        Wend
                        End If

                        ltemp = Left(ltemp, Len(ltemp) - 1)
                        lstr2 = lstr2 & "�Ӵ�" & lcolFactor("Σ������" & i) & "��ҵ��Ա" & lcolFactor("����" & i) & "�ˣ�"
                        If lstr3 = "" Then
                        lstr3 = lstr3 & "��" & i & "���Ӵ�" & lcolFactor("Σ������" & i) & "��Ա��" & Chr(13) & Chr(10) & IIf(ltemp = "", "�������δ�����ְҵ��ص��쳣�ı䡣", ltemp)
                        Else
                        lstr3 = lstr3 & Chr(13) & Chr(10) & "��" & i & "���Ӵ�" & lcolFactor("Σ������" & i) & "��Ա��" & Chr(13) & Chr(10) & IIf(ltemp = "", "�������δ�����ְҵ��ص��쳣�ı䡣", ltemp)
                        End If
                        If lstr4 = "" Then
                        lstr4 = lstr4 & "��" & i & "���Ӵ�" & lcolFactor("Σ������" & i) & "��Ա��" & Chr(13) & Chr(10) & IIf(ltemp = "", "�������δ���ְҵ����֢����ְҵ��صĽ����𺦡�", ltemp)
                        Else
                        lstr4 = lstr4 & Chr(13) & Chr(10) & "��" & i & "���Ӵ�" & lcolFactor("Σ������" & i) & "��Ա��" & Chr(13) & Chr(10) & IIf(ltemp = "", "�������δ���ְҵ����֢����ְҵ��صĽ����𺦡�", ltemp)
                        End If
'                        objWordDocument.Sections(1).Range.Find.Execute FindText:="��Σ�����ؼ������", ReplaceWith:="�Ӵ�" & lcolFactor("Σ������" & i) & "��Ա��" & IIf(ltemp = "", "�������δ�����ְҵ��ص��쳣�ı䡣", ltemp), Replace:=2
'                        objWordDocument.Sections(1).Range.Find.Execute FindText:="��Σ������������", ReplaceWith:=IIf(ltemp = "", "�������δ�����ְҵ��ص��쳣�ı䡣", ltemp), Replace:=2
'                        objWordDocument.Sections(1).Range.Find.Execute FindText:="��Σ������" & i & "��", ReplaceWith:="�Ӵ�" & lcolFactor("Σ������" & i) & "��Ա��", Replace:=2
'                        objWordDocument.Sections(1).Range.Find.Execute FindText:="��Σ�����" & i & "��", ReplaceWith:=IIf(ltemp = "", "�������δ�����ְҵ��ص��쳣�ı䡣", ltemp), Replace:=2

                    End If

                KlobjRec.MoveNext
                ltemp = ""
            Next
            lstr3 = lstr3 & Chr(13) & Chr(10) & "��" & i & "����ϸ���������ְҵ���������һ�����͸�����챨��"
            lstr4 = lstr4 & Chr(13) & Chr(10) & "��" & i & "������������������Ŀ����쳣�ߣ��ܼ��߿ɸ��������ٴ�֢״�򸴲����Ϊ�쳣����ߵ�ҽԺ��ؿ������Ρ���ϸ������ɼ����˱���"
            lstr2 = Left(lstr2, Len(lstr2) - 1)
'            lstr3 = Left(lstr3, Len(lstr3) - 1)
            
            objWordDocument.Sections(1).Range.Find.Execute FindText:="��Σ�����ء�", ReplaceWith:=lstr2, Replace:=2
            objWordDocument.Sections(1).Range.Find.Execute FindText:="��Σ�����ؼ������", ReplaceWith:=lstr3, Replace:=2
            objWordDocument.Sections(1).Range.Find.Execute FindText:="��Σ�����ؼ����ۡ�", ReplaceWith:=lstr4, Replace:=2
            
'            For i = 1 To lcolInfo.Count
'                lstr = "select ϵͳ���,����,Σ������,(������ + ��ϼ��������) from ְҵ�����_���������ݿ� where " _
'                        & "��λ���� = '" & lcolFactor("��λ����") & "' and Σ������ ='" & lcolInfo(i) & "' and (������� >= '" & lcolFactor("��ʼ����") & "' and ������� <= '" & lcolFactor("��ֹ����") & "')"
'                Set lobjRec = dafuncGetData(lstr)
'            Next i

'
'            '�������ģ��ڶ�������
            
'            Set lobjRec = dafuncGetData("select distinct a.ϵͳ���,a.����,a.Σ������,b.�����Ŀ as " _
'                & "��������,a.������+a.��Ϻʹ������ as ������� from ְҵ�����_���������ݿ� a,(select ϵͳ���,�����Ŀ=dbo.z_fc(ϵͳ���) from " _
'                & "ְҵ�����_�������ͼ where ������� = '���ϸ�') b where a.ϵͳ��� = b.ϵͳ��� and ��λ���� = '" & lcolFactor("��λ����") & "'" _
'                & " and (������� >= '" & lcolFactor("��ʼ����") & " 00:00:00' and ������� <='" & lcolFactor("��ֹ����") & " 23:59:59') and ���״̬ in( 7,6,5)")
'            Set lobjRec = dafuncGetData("select distinct a.ϵͳ���,a.����,a.Σ������,a.��Ϻʹ������ from ְҵ�����_���������ݿ� a where ��λ���� = '" & lcolFactor("��λ����") & "'" _
                & " and (������� >= '" & lcolFactor("��ʼ����") & " 00:00:00' and ������� <='" & lcolFactor("��ֹ����") & " 23:59:59') and ���״̬ in( 7,6,5)")
            
            
          '  ��ѯ��Ա��Ϣ������  2015-10-30  by Ĳ��
             Set lobjRec = dafuncGetData("select convert(varchar(100),�������,111) as �������,ϵͳ���,����,�Ա�,����,����,�ֹ���,Σ������,������,��Ϻʹ������ from ְҵ�����_���������ݿ�  where ��λ���� = '" & lcolFactor("��λ����") & "'" _
                & " and (������� >= '" & lcolFactor("��ʼ����") & " 00:00:00' and ������� <='" & lcolFactor("��ֹ����") & " 23:59:59') and ���״̬ >=1 ")

'             Set lobjRec = dafuncGetData("select �������,ϵͳ���,����,�Ա�,����,����,�ֹ���,Σ������,������,��Ϻʹ������ from ְҵ�����_���������ݿ�  where ��λ���� = '" & lcolFactor("��λ����") & "'" _
'                & " and (������� >= '" & Format(" & lcolFactor("��ʼ����") & ", "yyyy-mm-dd") & "' and ������� <='" & Format(" & lcolFactor("��ֹ����") & ", "yyyy-mm-dd") & "') and ���״̬ >=1 ")


'            Set lobjRec = dafuncGetData("select distinct b.�������,a.ϵͳ���,a.����,a.�Ա�,a.����,a.����,a.�ֹ���,a.Σ������,b.������,b.��Ϻʹ������ from ְҵ�����_�����Ա������Ϣ�� a ,ְҵ�����_��������Ϣ�� b where ��λ���� = '" & lcolFactor("��λ����") & "'" _
'                & " and (������� >= '" & lcolFactor("��ʼ����") & " 00:00:00' and ������� <='" & lcolFactor("��ֹ����") & " 23:59:59') and ���״̬ >=1 ")

            Set myRange = objWordDocument.Range(myRange.Start, myRange.StoryLength)
            Set myTable = myRange.Tables(1)
            If myTable.rows.Count <= lobjRec.RecordCount Then
                j = myTable.rows.Count
                myTable.rows(j).Select
                objWordDocument.ActiveWindow.Selection.InsertRows lobjRec.RecordCount - j + 1
                For i = 1 To lobjRec.RecordCount - myTable.rows.Count
                    myTable.rows.Add (myTable.rows(j))
                Next
            End If

            For i = 1 To lobjRec.RecordCount
                For j = 1 To lobjRec.Fields.Count
                    myTable.rows(i + 1).Cells(j).Range.Text = IIf(IsNull(lobjRec(j - 1)), "", lobjRec(j - 1))
                Next
                lobjRec.MoveNext
            Next
            
            '���Ӹ�����Աһ����������  2015-11-13 by Ĳ�� ��
            Dim lobjRec2 As Object
            Dim yijian As String
            Set lobjRec2 = dafuncGetData("select ϵͳ���,����,Σ������,����ԭ��,��Ϻʹ������ from ְҵ�����_���������ݿ�  where ��λ���� = '" & lcolFactor("��λ����") & "'" _
                & " and (������� >= '" & lcolFactor("��ʼ����") & " 00:00:00' and ������� <='" & lcolFactor("��ֹ����") & " 23:59:59') and ���״̬ >=1 and ��Ϻʹ������ like '%����ְҵ��������%'")
                Set myRange = objWordDocument.Range(myRange.Start, myRange.StoryLength)
                Set myTable = myRange.Tables(2)
                If myTable.rows.Count <= lobjRec2.RecordCount Then
                    j = myTable.rows.Count
                    myTable.rows(j).Select
                    objWordDocument.ActiveWindow.Selection.InsertRows lobjRec2.RecordCount - j + 1
                    For i = 1 To lobjRec2.RecordCount - myTable.rows.Count
                        myTable.rows.Add (myTable.rows(j))
                    Next
                End If

                For i = 1 To lobjRec2.RecordCount
                    For j = 1 To lobjRec2.Fields.Count
                        myTable.rows(i + 1).Cells(j).Range.Text = IIf(IsNull(lobjRec2(j - 1)), "", lobjRec2(j - 1))
                    Next
                    lobjRec2.MoveNext
                Next

           '2015-11-13 by Ĳ�� ��
            

        '�����ĵ����ģ�������ҳü��ҳ�ţ��е��������Ҫ��ҳ�롢ҳ��
        For i = 1 To objWordDocument.Range.Fields.Count
            objWordDocument.Range.Fields(i).Update
        Next
        '�����ļ�
        objWordDocument.SaveAs lstrNewDoc
        objWordDocument.Saved = False

        With objWord.ActiveWindow.View
            .ShowRevisionsAndComments = False
            .RevisionsView = 0          'wdRevisionsViewFinal
        End With
    
        objWordDocument.ActiveWindow.ActivePane.View.SeekView = 9
        objWordDocument.ActiveWindow.ActivePane.View.SeekView = 0
    
        objWord.Visible = True
    
        On Error GoTo errHandler
    
        objWordDocument.Activate
        objWord.Activate
        
        
        
        
'            '���ݶ���������Ա㱣���ļ������ݿ⡣
'        On Error Resume Next
'        If Not paraReadOnly Then
'            If lintRepID = 0 Then
'                objWord.Run "subStart", paraParent, -1, "", ""
'            Else
'                objWord.Run "subStart", paraParent, lintRepID, "", ""
'            End If
'        Else
'            If lintRepID = 0 Then
'                objWord.Run "subStart", Nothing, -1, "", ""
'            Else
'                objWord.Run "subStart", Nothing, lintRepID, "", ""
'            End If
'        End If
'        If Err.Number = 450 Then     '�������ԣ�˵����ģ��û��������Ʒ����Ĳ���
'            If Not paraReadOnly Then
'                If lintRepID = 0 Then
'                    objWord.Run "subStart", paraParent, -1, ""
'                Else
'                    objWord.Run "subStart", paraParent, lintRepID, ""
'                End If
'            Else
'                If lintRepID = 0 Then
'                    objWord.Run "subStart", Nothing, -1, ""
'                Else
'                    objWord.Run "subStart", Nothing, lintRepID, ""
'                End If
'            End If
'        End If
'        If Err.Number = 438 Then
'            MsgBox "�ñ����ģ��û�а��չ涨��д�����subStart���������޷����浽���ݿ��", vbOKOnly + vbCritical, "ϵͳ��ʾ"
'        End If
        Exit Sub
    
errHandler:
    If Err = 3001 Then
        MsgBox "û�������ݿ����ҵ���Word����ľ����ļ��������Ǳ���ñ��浽ϵͳ��ʱ���緢�����ϣ�����ϵͳ��ɾ���ñ�����Ϣ������¼��ñ��档", vbInformation, "ϵͳ��ʾ"
    Else
        sfsub������ "���沿��", "mod�������", "sub�༭�ܼ챨��", Err.Number, Err.Description, True
    End If
    Exit Sub
    Resume
End Sub
Sub sub������ͨ�Թ�����Ա����word��Ϣ(objWordDocument As Object, myRange As Object, paraSysNo As String)
      Dim lobjRec1 As Object, lobjRec2 As Object, lobjrec3 As Object, lobjrec4 As Object, lobjrec0 As Object  'ÿ��һ��object,ģ���ļ���5�ڡ�
    Dim myTable As Object
    Dim i As Integer, j As Integer
    Dim lobjRow As Object
    Dim strSQL As String
    Dim lstr������� As String
    Dim lstr���ս���, lstr��콨�� As String
    Dim lstrTmp
    
    On Error GoTo errHandler
    dasubSetQueryTimeout 600
    '��Ӻ͸����������
    strSQL = "select ������� from ְҵ�����_��������Ϣ�� where ϵͳ���='" & paraSysNo & "'"
    Set lobjrec0 = dafuncGetData(strSQL)
    lstr������� = Format(IIf(IsNull(lobjrec0("�������")), Now, lobjrec0("�������")), "yyyy-mm-dd")
    myRange.Find.Execute FindText:="��������ڡ�", ReplaceWith:=lstr�������, Replace:=2    'wdReplaceAll

    '�����ͷ����һ�����ݡ�
    strSQL = "select ϵͳ���,����,��λ����,סַ,�绰����,�ʱ�,�������� from ְҵ�����_�����Ա������Ϣ�� where ϵͳ���='" & paraSysNo & "'"
    Set lobjRec1 = dafuncGetData(strSQL)
    For i = 0 To lobjRec1.Fields.Count - 1
        myRange.Find.Execute FindText:="��" & lobjRec1.Fields(i).Name & "��", ReplaceWith:=IIf(IsNull(lobjRec1(i)), "", lobjRec1(i)), Replace:=2    'wdReplaceAll
'        For j = 1 To objWordDocument.Sections(1).Headers.Count
'            If lobjrec1.Fields(i).Name = "��������" Then
'                objWordDocument.Sections(1).Headers(j).Range.Find.Execute FindText:="��" & lobjrec1.Fields(i).Name & "��", ReplaceWith:=IIf(IsNull(lobjrec1(i)), "", Format(lobjrec1(i), "yyyy-mm-dd")), Replace:=2   'wdReplaceAll
'            Else
'                objWordDocument.Sections(1).Headers(j).Range.Find.Execute FindText:="��" & lobjrec1.Fields(i).Name & "��", ReplaceWith:=IIf(IsNull(lobjrec1(i)), "", lobjrec1(i)), Replace:=2    'wdReplaceAll
'            End If
'        Next
   
         If lobjRec1.Fields(i).Name = "��������" Then
              objWordDocument.Range.Find.Execute FindText:="��" & lobjRec1.Fields(i).Name & "��", ReplaceWith:=IIf(IsNull(lobjRec1(i)), "", Format(lobjRec1(i), "yyyy-mm-dd")), Replace:=2   'wdReplaceAll
         Else
             objWordDocument.Range.Find.Execute FindText:="��" & lobjRec1.Fields(i).Name & "��", ReplaceWith:=IIf(IsNull(lobjRec1(i)), "", lobjRec1(i)), Replace:=2    'wdReplaceAll
         End If
         
    Next
    '1
    frmProcess.proPercent.Value = frmProcess.proPercent.Value + 1
        DoEvents
    
    '�ڶ������ݣ�������Ϣ��ְҵ��ʷ��Ϣ��
    strSQL = "select �Ա�,����,��������,������,�Ļ��̶� from ְҵ�����_�����Ա������Ϣ�� where ϵͳ���='" & paraSysNo & "'"
    Set lobjRec2 = dafuncGetData(strSQL)
    For i = 0 To lobjRec2.Fields.Count - 1
        myRange.Find.Execute FindText:="��" & lobjRec2.Fields(i).Name & "��", ReplaceWith:=IIf(IsNull(lobjRec2(i)), "", lobjRec2(i)), Replace:=2    'wdReplaceAll
'        For j = 1 To objWordDocument.Sections(2).Headers.Count
'            objWordDocument.Sections(2).Headers(j).Range.Find.Execute FindText:="��" & lobjrec2.Fields(i).Name & "��", ReplaceWith:=IIf(IsNull(lobjrec2(i)), "", lobjrec2(i)), Replace:=2    'wdReplaceAll
'        Next

 
         objWordDocument.Range.Find.Execute FindText:="��" & lobjRec2.Fields(i).Name & "��", ReplaceWith:=IIf(IsNull(lobjRec2(i)), "", lobjRec2(i)), Replace:=2    'wdReplaceAll

     
    Next
    strSQL = "select * from ְҵ�����_��������ʷ�� where ϵͳ���='" & paraSysNo & "'"
    Set lobjRec2 = dafuncGetData(strSQL)
    For i = 0 To lobjRec2.Fields.Count - 1
        If Not lobjRec2.EOF Then
        myRange.Find.Execute FindText:="��" & lobjRec2.Fields(i).Name & "��", ReplaceWith:=IIf(IsNull(lobjRec2(i)), "", lobjRec2(i)), Replace:=2    'wdReplaceAll
'        For j = 1 To objWordDocument.Sections(2).Headers.Count
'            objWordDocument.Sections(2).Headers(j).Range.Find.Execute FindText:="��" & lobjrec2.Fields(i).Name & "��", ReplaceWith:=IIf(IsNull(lobjrec2(i)), "", lobjrec2(i)), Replace:=2    'wdReplaceAll
'        Next
        
        objWordDocument.Range.Find.Execute FindText:="��" & lobjRec2.Fields(i).Name & "��", ReplaceWith:=IIf(IsNull(lobjRec2(i)), "", lobjRec2(i)), Replace:=2    'wdReplaceAll

        Else
            myRange.Find.Execute FindText:="��" & lobjRec2.Fields(i).Name & "��", ReplaceWith:="", Replace:=2    'wdReplaceAll
        End If
    Next
    '2
    frmProcess.proPercent.Value = frmProcess.proPercent.Value + 1
        DoEvents

    '���������ݣ�����������ҽʦ��䡣
'    strSQL = "select * from ְҵ�����_�������ͼ where ϵͳ���='" & paraSysNo & "'"
    strSQL = "select b.���� as �����Ŀ,isnull(�����,'')as �����,ϵͳ��� from ְҵ�����_�������ͼ a right join ְҵ�����_�����Ŀ���ñ� b on a.�����Ŀ=b.���� and ϵͳ���='" & paraSysNo & "'"
    Set lobjrec3 = dafuncGetData(strSQL)
    lobjrec3.MoveFirst
    For i = 0 To lobjrec3.RecordCount - 1
        myRange.Find.Execute FindText:="��" & lobjrec3("�����Ŀ") & "��", ReplaceWith:=IIf(IsNull(lobjrec3("�����")), "", lobjrec3("�����")), Replace:=2    'wdReplaceAll
'        For j = 1 To objWordDocument.Sections(3).Headers.Count
'            objWordDocument.Sections(3).Headers(j).Range.Find.Execute FindText:="��" & lobjrec3("�����Ŀ") & "��", ReplaceWith:=IIf(IsNull(lobjrec3("�����")), "", lobjrec3("�����")), Replace:=2    'wdReplaceAll
'        Next

            objWordDocument.Range.Find.Execute FindText:="��" & lobjrec3("�����Ŀ") & "��", ReplaceWith:=IIf(IsNull(lobjrec3("�����")), "", lobjrec3("�����")), Replace:=2    'wdReplaceAll
        
        lobjrec3.MoveNext
    Next
    
       '3
    frmProcess.proPercent.Value = frmProcess.proPercent.Value + 1
        DoEvents
    
    strSQL = "select a.����,b.���� ҽʦ���� from ְҵ�����_���ҽ��۱� a, ϵͳ����_Ա��������Ϣ�� b where a.ϵͳ���='" & paraSysNo & "' and a.����<>'06' and a.ҽ�����=b.���"
    Set lobjrec3 = dafuncGetData(strSQL)
    lobjrec3.MoveFirst
    For i = 0 To lobjrec3.RecordCount - 1
        myRange.Find.Execute FindText:="��" & lobjrec3("����") & "���ҽʦ��", ReplaceWith:=IIf(IsNull(lobjrec3("����")), "", lobjrec3("ҽʦ����")), Replace:=2   'wdReplaceAll
'        For j = 1 To objWordDocument.Sections(3).Headers.Count
'            objWordDocument.Sections(3).Headers(j).Range.Find.Execute FindText:="��" & lobjrec3("����") & "���ҽʦ��", ReplaceWith:=IIf(IsNull(lobjrec3("����")), "", lobjrec3("ҽʦ����")), Replace:=2    'wdReplaceAll
'        Next

        objWordDocument.Range.Find.Execute FindText:="��" & lobjrec3("����") & "���ҽʦ��", ReplaceWith:=IIf(IsNull(lobjrec3("����")), "", lobjrec3("ҽʦ����")), Replace:=2   'wdReplaceAll
        
        lobjrec3.MoveNext
    Next
    '4
    frmProcess.proPercent.Value = frmProcess.proPercent.Value + 1
        DoEvents
dasubSetQueryTimeout 600
    '���Ľ����ݣ����ս���ͽ��ۡ�������䡣
    strSQL = "select * from ְҵ�����_���ҽ��۱� where ϵͳ���='" & paraSysNo & "' and ����='16'"
    Set lobjrec4 = dafuncGetData(strSQL)
    If lobjrec4.RecordCount > 0 Then
        lstrTmp = Split(lobjrec4("���ֽ���"), "_00_", -1, vbBinaryCompare)
        If UBound(lstrTmp) = 1 Then
            lstr���ս��� = lstrTmp(0)
            lstr��콨�� = lstrTmp(1)
        End If
        myRange.Find.Execute FindText:="�����ս��ۡ�", ReplaceWith:=lstr���ս���, Replace:=2    'wdReplaceAll
        myRange.Find.Execute FindText:="����콨�顿", ReplaceWith:=lstr��콨��, Replace:=2    'wdReplaceAll
'        For j = 1 To objWordDocument.Sections(4).Headers.Count
'            objWordDocument.Sections(4).Headers(j).Range.Find.Execute FindText:="�����ս����", ReplaceWith:=lstr���ս���, Replace:=2    'wdReplaceAll
'            objWordDocument.Sections(4).Headers(j).Range.Find.Execute FindText:="����콨�顿", ReplaceWith:=lstr��콨��, Replace:=2    'wdReplaceAll
'        Next

         objWordDocument.Range.Find.Execute FindText:="�����ս����", ReplaceWith:=lstr���ս���, Replace:=2    'wdReplaceAll
         objWordDocument.Range.Find.Execute FindText:="����콨�顿", ReplaceWith:=lstr��콨��, Replace:=2    'wdReplaceAll
    End If
    '5
    frmProcess.proPercent.Value = frmProcess.proPercent.Value + 1
        DoEvents
    '�滻��ٿ��۾����ͼƬ
    '����c���¸���һ�ݣ������������ν
'    Dim lobjSys As Object
'    Set lobjSys = CreateObject("Scripting.FileSystemObject")
'    lobjSys.copyfile App.Path & "\��״�廷״������ͼ.bmp", "c:\��״�廷״������ͼ.bmp"
'
'    '�����ݿ��и��Ƴ��������Ա��ͼƬ������c����ӦͼƬ
'    Set lobjSys = CreateObject("ְҵ�������¼��.ClsCommon")
'    frmFinalConclusion.libPicture.AutoRedraw = True
'    frmFinalConclusion.libPicture.Picture = lobjSys.func��ȡ���ͼƬ(paraSysNo, "01069", "��״�廷״������ͼ.bmp")
'    SavePicture frmFinalConclusion.libPicture.Picture, "c:\��״�廷״������ͼ.bmp"
 
    
    Set lobjRec1 = Nothing
    Set lobjRec2 = Nothing
    Set lobjrec3 = Nothing
    Set lobjrec4 = Nothing
    Set lobjrec0 = Nothing
    Set frmFinalConclusion.libPicture = Nothing
'    Set lobjSys = Nothing
    Exit Sub
errHandler:
End Sub
'2012-08-20 �ڵ��
'�������word���������ݵĴ��룬��˲��ӡ�8023����
Sub sub���ط��乤����Ա����word��Ϣ(objWordDocument As Object, myRange As Object, paraSysNo As String)
    Dim lobjRec1 As Object, lobjRec2 As Object, lobjrec3 As Object, lobjrec4 As Object, lobjrec0 As Object  'ÿ��һ��object,ģ���ļ���5�ڡ�
    Dim myTable As Object
    Dim i As Integer, j As Integer
    Dim lobjRow As Object
    Dim strSQL As String
    Dim lstr������� As String
    Dim lstr���ս���, lstr��콨�� As String
    Dim lstrTmp
    
    On Error GoTo errHandler
    dasubSetQueryTimeout 600
    '��Ӻ͸����������
    strSQL = "select ������� from ְҵ�����_��������Ϣ�� where ϵͳ���='" & paraSysNo & "'"
    Set lobjrec0 = dafuncGetData(strSQL)
    lstr������� = Format(IIf(IsNull(lobjrec0("�������")), Now, lobjrec0("�������")), "yyyy-mm-dd")
    myRange.Find.Execute FindText:="��������ڡ�", ReplaceWith:=lstr�������, Replace:=2    'wdReplaceAll

    '�����ͷ����һ�����ݡ�
    strSQL = "select ϵͳ���,����,��λ����,סַ,�绰����,�ʱ�,�������� from ְҵ�����_�����Ա������Ϣ�� where ϵͳ���='" & paraSysNo & "'"
    Set lobjRec1 = dafuncGetData(strSQL)
    For i = 0 To lobjRec1.Fields.Count - 1
        myRange.Find.Execute FindText:="��" & lobjRec1.Fields(i).Name & "��", ReplaceWith:=IIf(IsNull(lobjRec1(i)), "", lobjRec1(i)), Replace:=2    'wdReplaceAll
'        For j = 1 To objWordDocument.Sections(1).Headers.Count
'            If lobjrec1.Fields(i).Name = "��������" Then
'                objWordDocument.Sections(1).Headers(j).Range.Find.Execute FindText:="��" & lobjrec1.Fields(i).Name & "��", ReplaceWith:=IIf(IsNull(lobjrec1(i)), "", Format(lobjrec1(i), "yyyy-mm-dd")), Replace:=2   'wdReplaceAll
'            Else
'                objWordDocument.Sections(1).Headers(j).Range.Find.Execute FindText:="��" & lobjrec1.Fields(i).Name & "��", ReplaceWith:=IIf(IsNull(lobjrec1(i)), "", lobjrec1(i)), Replace:=2    'wdReplaceAll
'            End If
'        Next
   
         If lobjRec1.Fields(i).Name = "��������" Then
              objWordDocument.Range.Find.Execute FindText:="��" & lobjRec1.Fields(i).Name & "��", ReplaceWith:=IIf(IsNull(lobjRec1(i)), "", Format(lobjRec1(i), "yyyy-mm-dd")), Replace:=2   'wdReplaceAll
         Else
             objWordDocument.Range.Find.Execute FindText:="��" & lobjRec1.Fields(i).Name & "��", ReplaceWith:=IIf(IsNull(lobjRec1(i)), "", lobjRec1(i)), Replace:=2    'wdReplaceAll
         End If
         
    Next
    
       '1
       frmProcess.proPercent.Value = frmProcess.proPercent.Value + 1
    DoEvents

    '�ڶ������ݣ�������Ϣ��ְҵ��ʷ��Ϣ��
    strSQL = "select �Ա�,����,��������,������,�Ļ��̶� from ְҵ�����_�����Ա������Ϣ�� where ϵͳ���='" & paraSysNo & "'"
    Set lobjRec2 = dafuncGetData(strSQL)
    For i = 0 To lobjRec2.Fields.Count - 1
        myRange.Find.Execute FindText:="��" & lobjRec2.Fields(i).Name & "��", ReplaceWith:=IIf(IsNull(lobjRec2(i)), "", lobjRec2(i)), Replace:=2    'wdReplaceAll
'        For j = 1 To objWordDocument.Sections(2).Headers.Count
'            objWordDocument.Sections(2).Headers(j).Range.Find.Execute FindText:="��" & lobjrec2.Fields(i).Name & "��", ReplaceWith:=IIf(IsNull(lobjrec2(i)), "", lobjrec2(i)), Replace:=2    'wdReplaceAll
'        Next

 
         objWordDocument.Range.Find.Execute FindText:="��" & lobjRec2.Fields(i).Name & "��", ReplaceWith:=IIf(IsNull(lobjRec2(i)), "", lobjRec2(i)), Replace:=2    'wdReplaceAll

     
    Next
    strSQL = "select * from ְҵ�����_��������ʷ�� where ϵͳ���='" & paraSysNo & "'"
    Set lobjRec2 = dafuncGetData(strSQL)
    For i = 0 To lobjRec2.Fields.Count - 1
        If Not lobjRec2.EOF Then
        myRange.Find.Execute FindText:="��" & lobjRec2.Fields(i).Name & "��", ReplaceWith:=IIf(IsNull(lobjRec2(i)), "", lobjRec2(i)), Replace:=2    'wdReplaceAll
'        For j = 1 To objWordDocument.Sections(2).Headers.Count
'            objWordDocument.Sections(2).Headers(j).Range.Find.Execute FindText:="��" & lobjrec2.Fields(i).Name & "��", ReplaceWith:=IIf(IsNull(lobjrec2(i)), "", lobjrec2(i)), Replace:=2    'wdReplaceAll
'        Next
        
        objWordDocument.Range.Find.Execute FindText:="��" & lobjRec2.Fields(i).Name & "��", ReplaceWith:=IIf(IsNull(lobjRec2(i)), "", lobjRec2(i)), Replace:=2    'wdReplaceAll

        Else
            myRange.Find.Execute FindText:="��" & lobjRec2.Fields(i).Name & "��", ReplaceWith:="", Replace:=2    'wdReplaceAll
        End If
    Next

           '2
       frmProcess.proPercent.Value = frmProcess.proPercent.Value + 1
    DoEvents
    
    '���������ݣ�����������ҽʦ��䡣
'    strSQL = "select * from ְҵ�����_�������ͼ where ϵͳ���='" & paraSysNo & "'"
    strSQL = "select b.���� as �����Ŀ,isnull(�����,'')as �����,ϵͳ��� from ְҵ�����_�������ͼ a right join ְҵ�����_�����Ŀ���ñ� b on a.�����Ŀ=b.���� and ϵͳ���='" & paraSysNo & "'"
    Set lobjrec3 = dafuncGetData(strSQL)
    lobjrec3.MoveFirst
    For i = 0 To lobjrec3.RecordCount - 1
        myRange.Find.Execute FindText:="��" & lobjrec3("�����Ŀ") & "��", ReplaceWith:=IIf(IsNull(lobjrec3("�����")), "", lobjrec3("�����")), Replace:=2    'wdReplaceAll
'        For j = 1 To objWordDocument.Sections(3).Headers.Count
'            objWordDocument.Sections(3).Headers(j).Range.Find.Execute FindText:="��" & lobjrec3("�����Ŀ") & "��", ReplaceWith:=IIf(IsNull(lobjrec3("�����")), "", lobjrec3("�����")), Replace:=2    'wdReplaceAll
'        Next

            objWordDocument.Range.Find.Execute FindText:="��" & lobjrec3("�����Ŀ") & "��", ReplaceWith:=IIf(IsNull(lobjrec3("�����")), "", lobjrec3("�����")), Replace:=2    'wdReplaceAll
        
        lobjrec3.MoveNext
    Next
    strSQL = "select a.����,b.���� ҽʦ���� from ְҵ�����_���ҽ��۱� a, ϵͳ����_Ա��������Ϣ�� b where a.ϵͳ���='" & paraSysNo & "' and a.����<>'06' and a.ҽ�����=b.���"
    Set lobjrec3 = dafuncGetData(strSQL)
    lobjrec3.MoveFirst
    For i = 0 To lobjrec3.RecordCount - 1
        myRange.Find.Execute FindText:="��" & lobjrec3("����") & "���ҽʦ��", ReplaceWith:=IIf(IsNull(lobjrec3("����")), "", lobjrec3("ҽʦ����")), Replace:=2   'wdReplaceAll
'        For j = 1 To objWordDocument.Sections(3).Headers.Count
'            objWordDocument.Sections(3).Headers(j).Range.Find.Execute FindText:="��" & lobjrec3("����") & "���ҽʦ��", ReplaceWith:=IIf(IsNull(lobjrec3("����")), "", lobjrec3("ҽʦ����")), Replace:=2    'wdReplaceAll
'        Next

        objWordDocument.Range.Find.Execute FindText:="��" & lobjrec3("����") & "���ҽʦ��", ReplaceWith:=IIf(IsNull(lobjrec3("����")), "", lobjrec3("ҽʦ����")), Replace:=2   'wdReplaceAll
        
        lobjrec3.MoveNext
    Next
           '3
       frmProcess.proPercent.Value = frmProcess.proPercent.Value + 1
    DoEvents
        
        
    '���Ľ����ݣ����ս���ͽ��ۡ�������䡣
    strSQL = "select * from ְҵ�����_���ҽ��۱� where ϵͳ���='" & paraSysNo & "' and ����='16'"
    Set lobjrec4 = dafuncGetData(strSQL)
    If lobjrec4.RecordCount > 0 Then
        lstrTmp = Split(lobjrec4("���ֽ���"), "_00_", -1, vbBinaryCompare)
        If UBound(lstrTmp) = 1 Then
            lstr���ս��� = lstrTmp(0)
            lstr��콨�� = lstrTmp(1)
        End If
        myRange.Find.Execute FindText:="�����ս��ۡ�", ReplaceWith:=lstr���ս���, Replace:=2    'wdReplaceAll
        myRange.Find.Execute FindText:="����콨�顿", ReplaceWith:=lstr��콨��, Replace:=2    'wdReplaceAll
'        For j = 1 To objWordDocument.Sections(4).Headers.Count
'            objWordDocument.Sections(4).Headers(j).Range.Find.Execute FindText:="�����ս����", ReplaceWith:=lstr���ս���, Replace:=2    'wdReplaceAll
'            objWordDocument.Sections(4).Headers(j).Range.Find.Execute FindText:="����콨�顿", ReplaceWith:=lstr��콨��, Replace:=2    'wdReplaceAll
'        Next

         objWordDocument.Range.Find.Execute FindText:="�����ս����", ReplaceWith:=lstr���ս���, Replace:=2    'wdReplaceAll
         objWordDocument.Range.Find.Execute FindText:="����콨�顿", ReplaceWith:=lstr��콨��, Replace:=2    'wdReplaceAll
    End If
    
       '4
       frmProcess.proPercent.Value = frmProcess.proPercent.Value + 1
    DoEvents
    
    '�滻��ٿ��۾����ͼƬ
    '����c���¸���һ�ݣ������������ν
    Dim lobjSys As Object
    Set lobjSys = CreateObject("Scripting.FileSystemObject")
    lobjSys.copyfile App.Path & "\��״�廷״������ͼ.bmp", "c:\��״�廷״������ͼ.bmp"
    
    '�����ݿ��и��Ƴ��������Ա��ͼƬ������c����ӦͼƬ
    Set lobjSys = CreateObject("ְҵ�������¼��.ClsCommon")
    frmFinalConclusion.libPicture.AutoRedraw = True
    frmFinalConclusion.libPicture.Picture = lobjSys.func��ȡ���ͼƬ(paraSysNo, "01069", "��״�廷״������ͼ.bmp")
    SavePicture frmFinalConclusion.libPicture.Picture, "c:\��״�廷״������ͼ.bmp"
    
       '5
       frmProcess.proPercent.Value = frmProcess.proPercent.Value + 1
         DoEvents
    
    Set lobjRec1 = Nothing
    Set lobjRec2 = Nothing
    Set lobjrec3 = Nothing
    Set lobjrec4 = Nothing
    Set lobjrec0 = Nothing
    Set frmFinalConclusion.libPicture = Nothing
    Set lobjSys = Nothing
    Exit Sub
errHandler:
'    MsgBox ("sdgdg")
End Sub
Sub sub����8023�ͷ����Թ�����Աword��Ϣ(objWordDocument As Object, myRange As Object, paraSysNo As String)
     Dim lobjRec1 As Object, lobjRec2 As Object, lobjrec3 As Object, lobjrec4 As Object, lobjrec0 As Object  'ÿ��һ��object,ģ���ļ���5�ڡ�
    Dim myTable As Object
    Dim i As Integer, j As Integer
    Dim lobjRow As Object
    Dim strSQL As String
    Dim lstr������� As String
    Dim lstr���ս���, lstr��콨�� As String
    Dim lstrTmp
    
    On Error GoTo errHandler
    dasubSetQueryTimeout 600
    '��Ӻ͸����������
    strSQL = "select ������� from ְҵ�����_��������Ϣ�� where ϵͳ���='" & paraSysNo & "'"
    Set lobjrec0 = dafuncGetData(strSQL)
    lstr������� = Format(IIf(IsNull(lobjrec0("�������")), Now, lobjrec0("�������")), "yyyy-mm-dd")
    myRange.Find.Execute FindText:="��������ڡ�", ReplaceWith:=lstr�������, Replace:=2    'wdReplaceAll

    '�����ͷ����һ�����ݡ�
    strSQL = "select ϵͳ���,����,��λ����,סַ,�绰����,�ʱ�,�������� from ְҵ�����_�����Ա������Ϣ�� where ϵͳ���='" & paraSysNo & "'"
    Set lobjRec1 = dafuncGetData(strSQL)
    For i = 0 To lobjRec1.Fields.Count - 1
        myRange.Find.Execute FindText:="��" & lobjRec1.Fields(i).Name & "��", ReplaceWith:=IIf(IsNull(lobjRec1(i)), "", lobjRec1(i)), Replace:=2    'wdReplaceAll
        For j = 1 To objWordDocument.Sections(1).Headers.Count
            If lobjRec1.Fields(i).Name = "��������" Then
                objWordDocument.Sections(1).Headers(j).Range.Find.Execute FindText:="��" & lobjRec1.Fields(i).Name & "��", ReplaceWith:=IIf(IsNull(lobjRec1(i)), "", Format(lobjRec1(i), "yyyy-mm-dd")), Replace:=2   'wdReplaceAll
            Else
                objWordDocument.Sections(1).Headers(j).Range.Find.Execute FindText:="��" & lobjRec1.Fields(i).Name & "��", ReplaceWith:=IIf(IsNull(lobjRec1(i)), "", lobjRec1(i)), Replace:=2    'wdReplaceAll
            End If
        Next
    Next
    '1
    frmProcess.proPercent.Value = frmProcess.proPercent.Value + 1
        DoEvents
    '�ڶ������ݣ�������Ϣ��ְҵ��ʷ��Ϣ��
    strSQL = "select �Ա�,����,��������,������,�Ļ��̶� from ְҵ�����_�����Ա������Ϣ�� where ϵͳ���='" & paraSysNo & "'"
    Set lobjRec2 = dafuncGetData(strSQL)
    For i = 0 To lobjRec2.Fields.Count - 1
        myRange.Find.Execute FindText:="��" & lobjRec2.Fields(i).Name & "��", ReplaceWith:=IIf(IsNull(lobjRec2(i)), "", lobjRec2(i)), Replace:=2    'wdReplaceAll
        For j = 1 To objWordDocument.Sections(2).Headers.Count
            objWordDocument.Sections(2).Headers(j).Range.Find.Execute FindText:="��" & lobjRec2.Fields(i).Name & "��", ReplaceWith:=IIf(IsNull(lobjRec2(i)), "", lobjRec2(i)), Replace:=2    'wdReplaceAll
        Next
    Next
    strSQL = "select * from ְҵ�����_��������ʷ�� where ϵͳ���='" & paraSysNo & "'"
    Set lobjRec2 = dafuncGetData(strSQL)
    For i = 0 To lobjRec2.Fields.Count - 1
        If Not lobjRec2.EOF Then
            myRange.Find.Execute FindText:="��" & lobjRec2.Fields(i).Name & "��", ReplaceWith:=IIf(IsNull(lobjRec2(i)), "", lobjRec2(i)), Replace:=2    'wdReplaceAll
            For j = 1 To objWordDocument.Sections(2).Headers.Count
                objWordDocument.Sections(2).Headers(j).Range.Find.Execute FindText:="��" & lobjRec2.Fields(i).Name & "��", ReplaceWith:=IIf(IsNull(lobjRec2(i)), "", lobjRec2(i)), Replace:=2    'wdReplaceAll
            Next
        Else
            myRange.Find.Execute FindText:="��" & lobjRec2.Fields(i).Name & "��", ReplaceWith:="", Replace:=2    'wdReplaceAll
        End If
    Next
    '2
    frmProcess.proPercent.Value = frmProcess.proPercent.Value + 1
        DoEvents
dasubSetQueryTimeout 600
    '���������ݣ�����������ҽʦ��䡣
    strSQL = "select * from ְҵ�����_�������ͼ where ϵͳ���='" & paraSysNo & "'"
    Set lobjrec3 = dafuncGetData(strSQL)
    lobjrec3.MoveFirst
    For i = 0 To lobjrec3.RecordCount - 1
        myRange.Find.Execute FindText:="��" & lobjrec3("�����Ŀ") & "��", ReplaceWith:=IIf(IsNull(lobjrec3("�����")), "", lobjrec3("�����")), Replace:=2    'wdReplaceAll
        For j = 1 To objWordDocument.Sections(3).Headers.Count
            objWordDocument.Sections(3).Headers(j).Range.Find.Execute FindText:="��" & lobjrec3("�����Ŀ") & "��", ReplaceWith:=IIf(IsNull(lobjrec3("�����")), "", lobjrec3("�����")), Replace:=2    'wdReplaceAll
        Next
        lobjrec3.MoveNext
    Next
    '3
    frmProcess.proPercent.Value = frmProcess.proPercent.Value + 1
        DoEvents
        
    strSQL = "select a.����,b.���� ҽʦ���� from ְҵ�����_���ҽ��۱� a, ϵͳ����_Ա��������Ϣ�� b where a.ϵͳ���='" & paraSysNo & "' and a.����<>'06' and a.ҽ�����=b.���"
    Set lobjrec3 = dafuncGetData(strSQL)
    lobjrec3.MoveFirst
    For i = 0 To lobjrec3.RecordCount - 1
        myRange.Find.Execute FindText:="��" & lobjrec3("����") & "���ҽʦ��", ReplaceWith:=IIf(IsNull(lobjrec3("����")), "", lobjrec3("ҽʦ����")), Replace:=2   'wdReplaceAll
        For j = 1 To objWordDocument.Sections(3).Headers.Count
            objWordDocument.Sections(3).Headers(j).Range.Find.Execute FindText:="��" & lobjrec3("����") & "���ҽʦ��", ReplaceWith:=IIf(IsNull(lobjrec3("����")), "", lobjrec3("ҽʦ����")), Replace:=2    'wdReplaceAll
        Next
        lobjrec3.MoveNext
    Next
    '4
    frmProcess.proPercent.Value = frmProcess.proPercent.Value + 1
        DoEvents
        
    '���Ľ����ݣ����ս���ͽ��ۡ�������䡣
    strSQL = "select * from ְҵ�����_���ҽ��۱� where ϵͳ���='" & paraSysNo & "' and ����='16'"
    Set lobjrec4 = dafuncGetData(strSQL)
    If lobjrec4.RecordCount > 0 Then
        lstrTmp = Split(lobjrec4("���ֽ���"), "_00_", -1, vbBinaryCompare)
        If UBound(lstrTmp) = 1 Then
            lstr���ս��� = lstrTmp(0)
            lstr��콨�� = lstrTmp(1)
        End If
        myRange.Find.Execute FindText:="�����ս��ۡ�", ReplaceWith:=lstr���ս���, Replace:=2    'wdReplaceAll
        myRange.Find.Execute FindText:="����콨�顿", ReplaceWith:=lstr��콨��, Replace:=2    'wdReplaceAll
        For j = 1 To objWordDocument.Sections(4).Headers.Count
            objWordDocument.Sections(4).Headers(j).Range.Find.Execute FindText:="�����ս����", ReplaceWith:=lstr���ս���, Replace:=2    'wdReplaceAll
            objWordDocument.Sections(4).Headers(j).Range.Find.Execute FindText:="����콨�顿", ReplaceWith:=lstr��콨��, Replace:=2    'wdReplaceAll
        Next
    End If
    
    '5
    frmProcess.proPercent.Value = frmProcess.proPercent.Value + 1
        DoEvents
        
    '�滻��ٿ��۾����ͼƬ
    '����c���¸���һ�ݣ������������ν
    Dim lobjSys As Object
    Set lobjSys = CreateObject("Scripting.FileSystemObject")
    lobjSys.copyfile App.Path & "\��״�廷״������ͼ.bmp", "c:\��״�廷״������ͼ.bmp"
    
    '�����ݿ��и��Ƴ��������Ա��ͼƬ������c����ӦͼƬ
    Set lobjSys = CreateObject("ְҵ�������¼��.ClsCommon")
    frmFinalConclusion.libPicture.AutoRedraw = True
    frmFinalConclusion.libPicture.Picture = lobjSys.func��ȡ���ͼƬ(paraSysNo, "01069", "��״�廷״������ͼ.bmp")
    SavePicture frmFinalConclusion.libPicture.Picture, "c:\��״�廷״������ͼ.bmp"
    
    '6
    frmProcess.proPercent.Value = frmProcess.proPercent.Value + 1
       DoEvents
        
    Set lobjRec1 = Nothing
    Set lobjRec2 = Nothing
    Set lobjrec3 = Nothing
    Set lobjrec4 = Nothing
    Set lobjrec0 = Nothing
    Set frmFinalConclusion.libPicture = Nothing
    Set lobjSys = Nothing
    Unload frmProcess
    Exit Sub
errHandler:
End Sub
'�������word���������ݵĴ��룬���乤����
Sub sub������˹�����Աְҵ����word��Ϣ(objWordDocument As Object, myRange As Object, paraSysNo As String)
     Dim lobjRec1 As Object, lobjRec2 As Object, lobjrec3 As Object, lobjrec4 As Object, lobjrec0 As Object  'ÿ��һ��object,ģ���ļ���5�ڡ�
    Dim myTable As Object
    Dim i As Integer, j As Integer
    Dim lobjRow As Object
    Dim strSQL As String
    Dim lstr������� As String
    Dim lstr���ս���, lstr��콨�� As String
    Dim lstrTmp
    
    On Error GoTo errHandler
    dasubSetQueryTimeout 600
    '��Ӻ͸����������
    strSQL = "select ������� from ְҵ�����_��������Ϣ�� where ϵͳ���='" & paraSysNo & "'"
    Set lobjrec0 = dafuncGetData(strSQL)
    lstr������� = Format(IIf(IsNull(lobjrec0("�������")), Now, lobjrec0("�������")), "yyyy-mm-dd")
    myRange.Find.Execute FindText:="��������ڡ�", ReplaceWith:=lstr�������, Replace:=2    'wdReplaceAll

    '�����ͷ����һ�����ݡ�
    strSQL = "select ϵͳ���,����,��λ����,סַ,�绰����,�ʱ�,�������� from ְҵ�����_�����Ա������Ϣ�� where ϵͳ���='" & paraSysNo & "'"
    Set lobjRec1 = dafuncGetData(strSQL)
    For i = 0 To lobjRec1.Fields.Count - 1
        myRange.Find.Execute FindText:="��" & lobjRec1.Fields(i).Name & "��", ReplaceWith:=IIf(IsNull(lobjRec1(i)), "", lobjRec1(i)), Replace:=2    'wdReplaceAll
        For j = 1 To objWordDocument.Sections(1).Headers.Count
            If lobjRec1.Fields(i).Name = "��������" Then
                objWordDocument.Sections(1).Headers(j).Range.Find.Execute FindText:="��" & lobjRec1.Fields(i).Name & "��", ReplaceWith:=IIf(IsNull(lobjRec1(i)), "", Format(lobjRec1(i), "yyyy-mm-dd")), Replace:=2   'wdReplaceAll
            Else
                objWordDocument.Sections(1).Headers(j).Range.Find.Execute FindText:="��" & lobjRec1.Fields(i).Name & "��", ReplaceWith:=IIf(IsNull(lobjRec1(i)), "", lobjRec1(i)), Replace:=2    'wdReplaceAll
            End If
        Next
    Next
    '1
       frmProcess.proPercent.Value = frmProcess.proPercent.Value + 1
    DoEvents

    '�ڶ������ݣ�������Ϣ��ְҵ��ʷ��Ϣ��
    strSQL = "select �Ա�,����,��������,������,�Ļ��̶� from ְҵ�����_�����Ա������Ϣ�� where ϵͳ���='" & paraSysNo & "'"
    Set lobjRec2 = dafuncGetData(strSQL)
    For i = 0 To lobjRec2.Fields.Count - 1
        myRange.Find.Execute FindText:="��" & lobjRec2.Fields(i).Name & "��", ReplaceWith:=IIf(IsNull(lobjRec2(i)), "", lobjRec2(i)), Replace:=2    'wdReplaceAll
        For j = 1 To objWordDocument.Sections(2).Headers.Count
            objWordDocument.Sections(2).Headers(j).Range.Find.Execute FindText:="��" & lobjRec2.Fields(i).Name & "��", ReplaceWith:=IIf(IsNull(lobjRec2(i)), "", lobjRec2(i)), Replace:=2    'wdReplaceAll
        Next
    Next
    strSQL = "select * from ְҵ�����_��������ʷ�� where ϵͳ���='" & paraSysNo & "'"
    Set lobjRec2 = dafuncGetData(strSQL)
    For i = 0 To lobjRec2.Fields.Count - 1
        If Not lobjRec2.EOF Then
            myRange.Find.Execute FindText:="��" & lobjRec2.Fields(i).Name & "��", ReplaceWith:=IIf(IsNull(lobjRec2(i)), "", lobjRec2(i)), Replace:=2    'wdReplaceAll
            For j = 1 To objWordDocument.Sections(2).Headers.Count
                objWordDocument.Sections(2).Headers(j).Range.Find.Execute FindText:="��" & lobjRec2.Fields(i).Name & "��", ReplaceWith:=IIf(IsNull(lobjRec2(i)), "", lobjRec2(i)), Replace:=2    'wdReplaceAll
            Next
        Else
            myRange.Find.Execute FindText:="��" & lobjRec2.Fields(i).Name & "��", ReplaceWith:="", Replace:=2    'wdReplaceAll
        End If
    Next
    
       '2
       frmProcess.proPercent.Value = frmProcess.proPercent.Value + 1
    DoEvents

    '���������ݣ�����������ҽʦ��䡣
    strSQL = "select * from ְҵ�����_�������ͼ where ϵͳ���='" & paraSysNo & "'"
    Set lobjrec3 = dafuncGetData(strSQL)
    lobjrec3.MoveFirst
    For i = 0 To lobjrec3.RecordCount - 1
        myRange.Find.Execute FindText:="��" & lobjrec3("�����Ŀ") & "��", ReplaceWith:=IIf(IsNull(lobjrec3("�����")), "", lobjrec3("�����")), Replace:=2    'wdReplaceAll
        For j = 1 To objWordDocument.Sections(3).Headers.Count
            objWordDocument.Sections(3).Headers(j).Range.Find.Execute FindText:="��" & lobjrec3("�����Ŀ") & "��", ReplaceWith:=IIf(IsNull(lobjrec3("�����")), "", lobjrec3("�����")), Replace:=2    'wdReplaceAll
        Next
        lobjrec3.MoveNext
    Next
    strSQL = "select a.����,b.���� ҽʦ���� from ְҵ�����_���ҽ��۱� a, ϵͳ����_Ա��������Ϣ�� b where a.ϵͳ���='" & paraSysNo & "' and a.����<>'06' and a.ҽ�����=b.���"
    Set lobjrec3 = dafuncGetData(strSQL)
    lobjrec3.MoveFirst
    For i = 0 To lobjrec3.RecordCount - 1
        myRange.Find.Execute FindText:="��" & lobjrec3("����") & "���ҽʦ��", ReplaceWith:=IIf(IsNull(lobjrec3("����")), "", lobjrec3("ҽʦ����")), Replace:=2   'wdReplaceAll
        For j = 1 To objWordDocument.Sections(3).Headers.Count
            objWordDocument.Sections(3).Headers(j).Range.Find.Execute FindText:="��" & lobjrec3("����") & "���ҽʦ��", ReplaceWith:=IIf(IsNull(lobjrec3("����")), "", lobjrec3("ҽʦ����")), Replace:=2    'wdReplaceAll
        Next
        lobjrec3.MoveNext
    Next
    
       '3
       frmProcess.proPercent.Value = frmProcess.proPercent.Value + 1
    DoEvents

    '���Ľ����ݣ����ս���ͽ��ۡ�������䡣
    strSQL = "select * from ְҵ�����_���ҽ��۱� where ϵͳ���='" & paraSysNo & "' and ����='16'"
    Set lobjrec4 = dafuncGetData(strSQL)
    If lobjrec4.RecordCount > 0 Then
        lstrTmp = Split(lobjrec4("���ֽ���"), "_00_", -1, vbBinaryCompare)
        If UBound(lstrTmp) = 1 Then
            lstr���ս��� = lstrTmp(0)
            lstr��콨�� = lstrTmp(1)
        End If
        myRange.Find.Execute FindText:="�����ս��ۡ�", ReplaceWith:=lstr���ս���, Replace:=2    'wdReplaceAll
        myRange.Find.Execute FindText:="����콨�顿", ReplaceWith:=lstr��콨��, Replace:=2    'wdReplaceAll
        For j = 1 To objWordDocument.Sections(4).Headers.Count
            objWordDocument.Sections(4).Headers(j).Range.Find.Execute FindText:="�����ս����", ReplaceWith:=lstr���ս���, Replace:=2    'wdReplaceAll
            objWordDocument.Sections(4).Headers(j).Range.Find.Execute FindText:="����콨�顿", ReplaceWith:=lstr��콨��, Replace:=2    'wdReplaceAll
        Next
    End If
    
       '4
       frmProcess.proPercent.Value = frmProcess.proPercent.Value + 1
         DoEvents
    
    '�滻��ٿ��۾����ͼƬ
    '����c���¸���һ�ݣ������������ν
    Dim lobjSys As Object
    Set lobjSys = CreateObject("Scripting.FileSystemObject")
    lobjSys.copyfile App.Path & "\��״�廷״������ͼ.bmp", "c:\��״�廷״������ͼ.bmp"
    
    '�����ݿ��и��Ƴ��������Ա��ͼƬ������c����ӦͼƬ
    Set lobjSys = CreateObject("ְҵ�������¼��.ClsCommon")
    frmFinalConclusion.libPicture.AutoRedraw = True
    frmFinalConclusion.libPicture.Picture = lobjSys.func��ȡ���ͼƬ(paraSysNo, "01069", "��״�廷״������ͼ.bmp")
    SavePicture frmFinalConclusion.libPicture.Picture, "c:\��״�廷״������ͼ.bmp"
    
       '5
       frmProcess.proPercent.Value = frmProcess.proPercent.Value + 1
         DoEvents
    
    Set lobjRec1 = Nothing
    Set lobjRec2 = Nothing
    Set lobjrec3 = Nothing
    Set lobjrec4 = Nothing
    Set lobjrec0 = Nothing
    Set frmFinalConclusion.libPicture = Nothing
    Set lobjSys = Nothing
    Exit Sub
errHandler:
'    MsgBox ("sdgdg")
End Sub

'2012-08-20 �ڵ�� �޸ģ������ ʱ�䣺2013-1-9
'�������word���������ݵĴ��룬ְҵ������
Sub sub����ְҵ����word��Ϣ(objWordDocument As Object, myRange As Object, paraSysNo As String)
    Dim lobjRec1 As Object, lobjRec2 As Object, lobjrec3 As Object, lobjrec4 As Object, lobjrec0 As Object  'ÿ��һ��object,ģ���ļ���5�ڡ�
    Dim myTable As Object
    Dim i As Integer, j As Integer
    Dim lobjRow As Object
    Dim strSQL As String
    Dim lstr������� As String
    Dim lstr���ս���, lstr��콨�� As String
    Dim lstrTmp
    
    On Error GoTo errHandler
    dasubSetQueryTimeout 600
    '��Ӻ͸����������
    strSQL = "select ������� from ְҵ�����_��������Ϣ�� where ϵͳ���='" & paraSysNo & "'"
    Set lobjrec0 = dafuncGetData(strSQL)
    lstr������� = Format(IIf(IsNull(lobjrec0("�������")), Now, lobjrec0("�������")), "yyyy-mm-dd")
    myRange.Find.Execute FindText:="��������ڡ�", ReplaceWith:=lstr�������, Replace:=2    'wdReplaceAll

    '�����ͷ����һ�����ݡ�
    strSQL = "select ϵͳ���,����,�Ա�,��������,�Ļ��̶�,��λ����,�绰����,������ݺ���,����,�ֹ���,ְҵΣ������,Σ������,�������� from ְҵ�����_�����Ա������Ϣ��  where ϵͳ���='" & paraSysNo & "'"
    Set lobjRec1 = dafuncGetData(strSQL)
    For i = 0 To lobjRec1.Fields.Count - 1
        myRange.Find.Execute FindText:="��" & lobjRec1.Fields(i).Name & "��", ReplaceWith:=IIf(IsNull(lobjRec1(i)), "", lobjRec1(i)), Replace:=2    'wdReplaceAll
'        For j = 1 To objWordDocument.Sections(1).Headers.Count
            If lobjRec1.Fields(i).Name = "��������" Then
                objWordDocument.Range.Find.Execute FindText:="��" & lobjRec1.Fields(i).Name & "��", ReplaceWith:=IIf(IsNull(lobjRec1(i)), "", Format(lobjRec1(i), "yyyy-mm-dd")), Replace:=2   'wdReplaceAll
            Else
                objWordDocument.Range.Find.Execute FindText:="��" & lobjRec1.Fields(i).Name & "��", ReplaceWith:=IIf(IsNull(lobjRec1(i)), "", lobjRec1(i)), Replace:=2    'wdReplaceAll
            End If
'        Next
    Next
    '1
    frmProcess.proPercent.Value = frmProcess.proPercent.Value + 1
        DoEvents
    '�ڶ������ݣ�������Ϣ��ְҵ��ʷ��Ϣ��
'    strSQL = "select �Ա�,����,��������,������,�Ļ��̶� from ְҵ�����_�����Ա������Ϣ�� where ϵͳ���='" & paraSysNo & "'"
'    Set lobjrec2 = dafuncGetData(strSQL)
'    For i = 0 To lobjrec2.Fields.Count - 1
'        myRange.Find.Execute FindText:="��" & lobjrec2.Fields(i).Name & "��", ReplaceWith:=IIf(IsNull(lobjrec2(i)), "", lobjrec2(i)), Replace:=2    'wdReplaceAll
''        For j = 1 To objWordDocument.Sections(2).Headers.Count
'            objWordDocument.Range.Find.Execute FindText:="��" & lobjrec2.Fields(i).Name & "��", ReplaceWith:=IIf(IsNull(lobjrec2(i)), "", lobjrec2(i)), Replace:=2    'wdReplaceAll
''        Next
'    Next
    
    strSQL = "select ��ʼʱ��,������λ,����,Σ������,������ʩ from ְҵ�����_ְҵʷ�� where ϵͳ���='" & paraSysNo & "'"
    Set lobjRec2 = dafuncGetData(strSQL)
    If Not (lobjRec2.EOF Or lobjRec2.BOF) Then
'        objWordDocument.Sections(2).Range.Find.Execute FindText:="��ְҵʷ��", ReplaceWith:="", Replace:=2
        Set myRange = objWordDocument.Range(myRange.Start, myRange.StoryLength)
        Set myTable = myRange.Tables(3)
        If myTable.rows.Count < lobjRec2.RecordCount Then
            j = myTable.rows.Count
            myTable.rows(j).Select
            objWordDocument.ActiveWindow.Selection.InsertRows lobjRec2.RecordCount - j + 1
            For i = 1 To lobjRec2.RecordCount - myTable.rows.Count
                myTable.rows.Add (myTable.rows(j))
            Next
        End If
        
        For i = 1 To lobjRec2.RecordCount
            For j = 1 To lobjRec2.Fields.Count
                myTable.rows(i + 1).Cells(j).Range.Text = IIf(IsNull(lobjRec2(j - 1)), "", lobjRec2(j - 1))
            Next
            lobjRec2.MoveNext
        Next
        
    End If
    
       strSQL = "select * from ְҵ�����_��������ʷ�� where ϵͳ���='" & paraSysNo & "'"
    Set lobjRec2 = dafuncGetData(strSQL)
    For i = 0 To lobjRec2.Fields.Count - 1
        If Not lobjRec2.EOF Then
        myRange.Find.Execute FindText:="��" & lobjRec2.Fields(i).Name & "��", ReplaceWith:=IIf(IsNull(lobjRec2(i)), "", lobjRec2(i)), Replace:=2    'wdReplaceAll
'        For j = 1 To objWordDocument.Sections(2).Headers.Count
            objWordDocument.Range.Find.Execute FindText:="��" & lobjRec2.Fields(i).Name & "��", ReplaceWith:=IIf(IsNull(lobjRec2(i)), "", lobjRec2(i)), Replace:=2    'wdReplaceAll
'        Next
        Else
            myRange.Find.Execute FindText:="��" & lobjRec2.Fields(i).Name & "��", ReplaceWith:="", Replace:=2    'wdReplaceAll
        End If
    Next
         strSQL = "select ������ from ְҵ�����_�����Ա������Ϣ��  where ϵͳ���='" & paraSysNo & "'"
         Set lobjRec2 = dafuncGetData(strSQL)
         For i = 0 To lobjRec2.Fields.Count - 1
            If Not (lobjRec2.EOF Or lobjRec2.BOF) Then
                myRange.Find.Execute FindText:="��" & lobjRec2.Fields(i).Name & "��", ReplaceWith:=IIf(IsNull(lobjRec2(i)), "", lobjRec2(i)), Replace:=2    'wdReplaceAll
                objWordDocument.Range.Find.Execute FindText:="��" & lobjRec2.Fields(i).Name & "��", ReplaceWith:=IIf(IsNull(lobjRec2(i)), "", lobjRec2(i)), Replace:=2    'wdReplaceAll
            Else
            myRange.Find.Execute FindText:="��" & lobjRec2.Fields(i).Name & "��", ReplaceWith:="", Replace:=2    'wdReplaceAll
        End If
    Next
    
    strSQL = "select ���,��������,�������,��ϵ�λ,���ƾ���,ת�� from ְҵ�����_������ʷ�� where ϵͳ���='" & paraSysNo & "'"
    Set lobjRec2 = dafuncGetData(strSQL)
    If Not (lobjRec2.EOF Or lobjRec2.BOF) Then
'        objWordDocument.Sections(2).Range.Find.Execute FindText:="��ְҵʷ��", ReplaceWith:="", Replace:=2
        Set myRange = objWordDocument.Range(myRange.Start, myRange.StoryLength)
        Set myTable = myRange.Tables(4)
        If myTable.rows.Count < lobjRec2.RecordCount Then
            j = myTable.rows.Count
            myTable.rows(j).Select
            objWordDocument.ActiveWindow.Selection.InsertRows lobjRec2.RecordCount - j + 1
            For i = 1 To lobjRec2.RecordCount - myTable.rows.Count
                myTable.rows.Add (myTable.rows(j))
            Next
        End If
        
        For i = 1 To lobjRec2.RecordCount
            For j = 1 To lobjRec2.Fields.Count
                myTable.rows(i + 1).Cells(j).Range.Text = IIf(IsNull(lobjRec2(j - 1)), "", lobjRec2(j - 1))
            Next
            lobjRec2.MoveNext
        Next
        
    End If
    strSQL = "select * from ְҵ�����_�Ծ�֢״�� where ϵͳ���='" & paraSysNo & "'"
    Set lobjRec2 = dafuncGetData(strSQL)
    If Not (lobjRec2.EOF Or lobjRec2.BOF) Then
'        objWordDocument.Sections(2).Range.Find.Execute FindText:="��ְҵʷ��", ReplaceWith:="", Replace:=2
        Set myRange = objWordDocument.Range(myRange.Start, myRange.StoryLength)
        Set myTable = myRange.Tables(6)
        If myTable.rows.Count < lobjRec2.RecordCount Then
            j = myTable.rows.Count
            myTable.rows(j).Select
            objWordDocument.ActiveWindow.Selection.InsertRows lobjRec2.RecordCount - j + 1
            For i = 1 To lobjRec2.RecordCount - myTable.rows.Count
                myTable.rows.Add (myTable.rows(j))
            Next
        End If
        
        For i = 1 To lobjRec2.RecordCount
            For j = 1 To lobjRec2.Fields.Count
                myTable.rows(i + 1).Cells(j).Range.Text = IIf(IsNull(lobjRec2(j - 1)), "", lobjRec2(j - 1))
            Next
            lobjRec2.MoveNext
        Next
        
    End If
     '2
        frmProcess.proPercent.Value = frmProcess.proPercent.Value + 1
        DoEvents
    
    '���������ݣ�����������ҽʦ��䡣
'    strSQL = "select * from ְҵ�����_�������ͼ where ϵͳ���='" & paraSysNo & "'"
    strSQL = "select b.���� as �����Ŀ,isnull(�����,'')as �����,ϵͳ��� from ְҵ�����_�������ͼ a right join ְҵ�����_�����Ŀ���ñ� b on a.�����Ŀ=b.���� and ϵͳ���='" & paraSysNo & "'"
    Set lobjrec3 = dafuncGetData(strSQL)
    lobjrec3.MoveFirst
    For i = 0 To lobjrec3.RecordCount - 1
        myRange.Find.Execute FindText:="��" & lobjrec3("�����Ŀ") & "��", ReplaceWith:=IIf(IsNull(lobjrec3("�����")), "", lobjrec3("�����")), Replace:=2    'wdReplaceAll
'        For j = 1 To objWordDocument.Sections(3).Headers.Count
            objWordDocument.Range.Find.Execute FindText:="��" & lobjrec3("�����Ŀ") & "��", ReplaceWith:=IIf(IsNull(lobjrec3("�����")), "", lobjrec3("�����")), Replace:=2    'wdReplaceAll
'        Next
        lobjrec3.MoveNext
    Next
    '3
    frmProcess.proPercent.Value = frmProcess.proPercent.Value + 1
    DoEvents
    
    strSQL = "select a.����,b.���� ҽʦ���� from ְҵ�����_���ҽ��۱� a, ϵͳ����_Ա��������Ϣ�� b where a.ϵͳ���='" & paraSysNo & "' and a.����<>'06' and a.ҽ�����=b.���"
    Set lobjrec3 = dafuncGetData(strSQL)
    lobjrec3.MoveFirst
    For i = 0 To lobjrec3.RecordCount - 1
        myRange.Find.Execute FindText:="��" & lobjrec3("����") & "���ҽʦ��", ReplaceWith:=IIf(IsNull(lobjrec3("����")), "", lobjrec3("ҽʦ����")), Replace:=2   'wdReplaceAll
'        For j = 1 To objWordDocument.Sections(3).Headers.Count
            objWordDocument.Range.Find.Execute FindText:="��" & lobjrec3("����") & "���ҽʦ��", ReplaceWith:=IIf(IsNull(lobjrec3("����")), "", lobjrec3("ҽʦ����")), Replace:=2    'wdReplaceAll
'        Next
        lobjrec3.MoveNext
    Next
    '4
       frmProcess.proPercent.Value = frmProcess.proPercent.Value + 1
         DoEvents
         
    '�滻û�еĿ���
    For i = 1 To 17
        If Not i = 16 Then
            myRange.Find.Execute FindText:="��0" & i & "���ҽʦ��", ReplaceWith:="", Replace:=2   'wdReplaceAll
        End If
'        For j = 1 To objWordDocument.Sections(3).Headers.Count
            objWordDocument.Range.Find.Execute FindText:="��0" & i & "���ҽʦ��", ReplaceWith:="", Replace:=2    'wdReplaceAll
'        Next
'        lobjrec3.MoveNext
    Next
        
    '���Ľ����ݣ����ս���ͽ��ۡ�������䡣
    strSQL = "select * from ְҵ�����_���ҽ��۱� where ϵͳ���='" & paraSysNo & "' and ����='16'"
    Set lobjrec4 = dafuncGetData(strSQL)
    If lobjrec4.RecordCount > 0 Then
        lstrTmp = Split(lobjrec4("���ֽ���"), "_00_", -1, vbBinaryCompare)
        If UBound(lstrTmp) = 1 Then
            lstr���ս��� = lstrTmp(0)
            lstr��콨�� = lstrTmp(1)
        End If
        myRange.Find.Execute FindText:="�����ս��ۡ�", ReplaceWith:=lstr���ս���, Replace:=2    'wdReplaceAll
        myRange.Find.Execute FindText:="����콨�顿", ReplaceWith:=lstr��콨��, Replace:=2    'wdReplaceAll
'        For j = 1 To objWordDocument.Sections(4).Headers.Count
            objWordDocument.Range.Find.Execute FindText:="�����ս����", ReplaceWith:=lstr���ս���, Replace:=2    'wdReplaceAll
            objWordDocument.Range.Find.Execute FindText:="����콨�顿", ReplaceWith:=lstr��콨��, Replace:=2    'wdReplaceAll
'        Next
    End If
    '5
    frmProcess.proPercent.Value = frmProcess.proPercent.Value + 1
    DoEvents
    
    '�滻��ٿ��۾����ͼƬ
    '����c���¸���һ�ݣ������������ν
'    Dim lobjSys As Object
'    Set lobjSys = CreateObject("Scripting.FileSystemObject")
'    lobjSys.copyfile App.Path & "\��״�廷״������ͼ.bmp", "c:\��״�廷״������ͼ.bmp"
'
'    '�����ݿ��и��Ƴ��������Ա��ͼƬ������c����ӦͼƬ
'    Set lobjSys = CreateObject("ְҵ�������¼��.ClsCommon")
'    frmFinalConclusion.libPicture.AutoRedraw = True
'    frmFinalConclusion.libPicture.Picture = lobjSys.func��ȡ���ͼƬ(paraSysNo, "01069", "��״�廷״������ͼ.bmp")
'    SavePicture frmFinalConclusion.libPicture.Picture, "c:\��״�廷״������ͼ.bmp"
'
    Set lobjRec1 = Nothing
    Set lobjRec2 = Nothing
    Set lobjrec3 = Nothing
    Set lobjrec4 = Nothing
    Set lobjrec0 = Nothing
    Set frmFinalConclusion.libPicture = Nothing
'    Set lobjSys = Nothing

    Exit Sub
   
errHandler:
'    MsgBox ("sdgdg")
End Sub
'���ߣ������ ʱ��:2013-1-8 ��
'ȡ��ѡ���Ӧ��Ϣ��wordģ��
Sub subȡ��wordģ��()
    Dim lstrFile As String
    Dim lobjRec As Object
    
    lstrFile = Dir(App.Path & "\ͨ��_*.dot")
    Do While lstrFile <> ""
'       clstFile.AddItem lstrFile
       lstrFile = Dir
    Loop
    'Ѱ�ҵ�ǰ�û�������������Ӧ��ר��ģ����ǰ׺
    Set lobjRec = dafuncGetData("select ���� from ϵͳ����_�����ֵ�� where ���='" & um�û��������ұ�� & "'")

'        lstrFile = Dir(App.Path & "\" & lobjRec(0) & "_�Ĵ�ʡ" & Left(pstrWordname, 2) & "*.dot")
        lstrFile = Dir(App.Path & "\ְҵ�����_�Ĵ�ʡ" & Left(pstrWordname, 2) & "*.dot")

    If lstrFile = "" Then
        MsgBox "û���ҵ�Wordģ���ļ���", vbInformation, "ϵͳ��ʾ"
       Exit Sub
    Else
        pstrFilename = lstrFile
  End If
End Sub
'���ĵ����浽���ݿ�
'word�ĵ��ر�ʱ�ᴥ���÷����ĵ��á�paraϵͳ���
Public Sub subSaveDoc(ByVal paraFile As String, ByVal paraNo As Integer, ByVal paraϵͳ��� As String)
    '���������Ϣ��
    Dim lobjRec As Object
    Dim lstrFileType As String, lstrNo As String
    Dim i As Integer, lstrType As String
    Dim lstr������ As String
    Dim lstr������� As String
    Dim lstr������ As String
    
'    If mblnֻ�� Then Exit Sub
    
    On Error GoTo errHandler
    
    For i = Len(paraFile) To 1 Step -1
        If Mid(paraFile, i, 1) = "." Then Exit For
    Next
    lstrFileType = Mid(paraFile, i + 1)
    'Ѱ�ҵ�ǰ�û����ܲ����ļ������
'    Set lobjRec = dafuncGetData("select ���� from �������_����ֹ������ͼ where Ա�����='" & um�û���� & "' order by ���")
'    If lobjRec.RecordCount > 0 Then lstrType = lobjRec(0)
    
'    Set lobjRec = dafuncGetData("select ��� from ְҵ�����_��챨����Ϣ�� where ���=" & paraNo & " and ���='" & lstrType & "'")
    dasubBeginTran
    Set lobjRec = dafuncGetData("select ������ from ְҵ�����_��챨����Ϣ�� where ϵͳ���='" & paraϵͳ��� & "'")
    If lobjRec.RecordCount = 0 Then
        Set lobjRec = dafuncGetData("select ������,�������,������ from dbo.ְҵ�����_��������Ϣ��  where ϵͳ���='" & paraϵͳ��� & "'")
      If lobjRec.RecordCount <> 0 Then
            lstr������� = lobjRec(1)
            lstr������ = lobjRec(2)
            lstr������ = lobjRec(0)
    Else
        MsgBox "����Ϣ�Ļ�����Ϣ�����ڣ������޷����棡", vbInformation, "ϵͳ��ʾ��"
        Exit Sub
    End If
        dafuncGetData "insert into ְҵ�����_��챨����Ϣ��(ϵͳ���,������,�������,�ļ�����,������,�������,������,������,�޸�����) values('" & paraϵͳ��� & "', '" & paraϵͳ��� & "' ,'���','" & UCase(lstrFileType) & "','" & um�û���� & "', '" & lstr������� & "','" & lstr������ & "','" & lstr������ & "',getdate() " & ")"
        'Ѱ���±�����ļ��ı��
'        Set lobjRec = dafuncGetData("select max(���) from ְҵ�����_��챨����Ϣ�� where ������='" & para������ & "' and ������='" & um�û���� & "'")
        Set lobjRec = dafuncGetData("select max(ϵͳ���) from ְҵ�����_��챨����Ϣ�� where ������='" & paraϵͳ��� & "'")
        
        lstrNo = lobjRec(0)
    Else
'        dafuncGetData "update ְҵ�����_��챨����Ϣ�� set �ļ�����='" & UCase(lstrFileType) & "',�޸�����=getdate(),������='" & um�û���� & "' where ���=" & paraNo
'        lstrNo = CStr(paraNo)

         dafuncGetData "update ְҵ�����_��챨����Ϣ�� set �ļ�����='" & UCase(lstrFileType) & "',�޸�����=getdate(),������='" & um�û���� & "' where ������='" & paraϵͳ��� & "'"
        lstrNo = CStr(paraϵͳ���)
    End If
    
    '�����ĵ��ļ������ݿ⡣
     pobjFileToDatabase.subFileToColumn "ְҵ�����_��챨����Ϣ��", "����", "������=" & lstrNo, paraFile
'     pobjFileToDatabase.subFileToColumn "ϵͳ����_�����鱨����Ϣ��", "�ĵ�", "���=" & lstrNo, paraFile
    dafuncGetData "update ְҵ�����_��������Ϣ�� set ���״̬=7 where ϵͳ���='" & paraϵͳ��� & "'"
   dasubCommitTran
   
    On Error Resume Next
'    oeExamSubSave "���鱨�����", frm���Ʊ���.pstr������, "���Ʊ���"
'    frmFinalConclusion.subRefreshView
    Exit Sub
errHandler:
    sfsub������ "���沿��", "mod�������", "subSave", Err.Number, Err.Description, False
    Exit Sub
    Resume
End Sub
''��WORD����ã����ڱ���WORD�����ݿ�
'Public Sub subSave(ByVal paraFile As String, ByVal paraNo As Integer, ByVal para������ As String)
'    subSaveDoc paraFile, paraNo, para������
'End Sub

'���ܣ�Ϊ�ռ���Ϣ��������ȡָ���������ռ���Ϣ��
Public Function func��ȡ�ռ������Ϣ(ByVal paraOffice As String, Optional ByVal paraFilter As String) As Recordset
    
    On Error GoTo errHandle
    
    dasubSetQueryTimeout 600
    
'    Set func��ȡ�ռ������Ϣ = dafuncGetData("exec ְҵ�����_��ѯ��������Ϣ '" & paraOffice & "','" & paraFilter & "','" & um�û���� & "'")
     Set func��ȡ�ռ������Ϣ = dafuncGetData(" select ϵͳ���,������,�������,������,�������,���״̬ from ְҵ�����_��������Ϣ�� where 1=1  order by ���״̬ ")
    
    Exit Function
    
errHandle:

    sfsub������ "ְҵ������", "modMain", "func��ȡ�ռ������Ϣ", Err.Number, Err.Description, True
        
End Function
Public Sub sub��ȡword�ĵ�(paraParent As Object, ByVal paraϵͳ��� As String, ByVal para������ As String, ByVal paraReadOnly As Boolean)

    Dim objWord As Object 'Word.Application
    Dim objWordDocument As Object 'Word.Document
    Dim lobjRec As Object, lobjRec1 As Object
    Dim lstrType As String, lstrTypeNo As String
    Dim lintRepID As Integer        '����Word�����ID
    
    On Error GoTo errHandler
    
    '����word��
    On Error Resume Next
    
    MkDir App.Path & "\temp"
    Kill App.Path & "\temp\*.*"
    
    Const CLASSOBJECT = "Word.Application"
    Set objWord = GetObject(, CLASSOBJECT)
    
    Err.Clear
    If (objWord Is Nothing) Then
        Set objWord = CreateObject(CLASSOBJECT)
        If Err.Number <> 0 Then
            On Error GoTo errHandler
            Err.Raise 6666, , "��û�а�װWord���޷��༭���棡���Ȱ�װMS Office 2000���ϰ汾�� "
        End If
    End If
    
    objWord.UserName = um�û���
    objWord.Options.UpdateLinksAtOpen = False       '����Word��ʾ�û�����ǩ��ͼƬ
    objWord.Options.CheckGrammarAsYouType = False   '��ֹƴд�����﷨���
    objWord.Options.CheckSpellingAsYouType = False
    
    On Error GoTo errHandler
    
    Dim lstrNewDoc As String
    Dim lstrDotFile As String 'ģ���ļ���
    Dim j As Integer
    

    Dim lpicPhoto As StdPicture
    Dim lobjSys As Object '
    Dim lstr������ As String
    Dim i As Long, lstr��Ʒ���� As String
    
    Set lobjSys = CreateObject("Scripting.FileSystemObject")

    
    Dim lstr�������� As String      '���˼���ʹ��
    Dim lsngHeight As Single
    
    '�ж������Ƿ��Ѵ��ڡ�
    Set lobjRec = dafuncGetData("select ������,�ļ����� from ְҵ�����_��챨����Ϣ�� where ������='" & para������ & "' and ϵͳ���='" & paraϵͳ��� & "'")
    If lobjRec.RecordCount = 0 Then
        If paraReadOnly Then        '�������޸ģ�����Ϊ�鿴������������������
            MsgBox "����Ʒû��¼��Word���棬����ʱ������Ϊ�����Word���棡", vbOKOnly + vbInformation, "ϵͳ��ʾ"
            Exit Sub
        End If
        
    Else
        lintRepID = CInt(lobjRec(0))
        '�༭�������顣
        lstrNewDoc = App.Path & "\temp\" & lstr������ & "_" & Format(Now, "yymmddhhmmss") & "." & lobjRec(1)
        'ȡ���ĵ���
        pobjFileToDatabase.subColumnToFile "ְҵ�����_��챨����Ϣ��", "�ĵ�", "���=" & lobjRec(0), lstrNewDoc
        
        'ֱ�Ӵ����е��ĵ�
        Set objWordDocument = objWord.Documents.Open(FileName:=lstrNewDoc, ReadOnly:=paraReadOnly)
        
        On Error Resume Next
        '������
        For i = 1 To objWordDocument.Range.Fields.Count
            objWordDocument.Range.Fields(i).Update
        Next
       
         
    End If
    With objWord.ActiveWindow.View
        .ShowRevisionsAndComments = False
        .RevisionsView = 0      'wdRevisionsViewFinal
    End With
    
'    If objWord.Version = 11 Then
'    objWordDocument.CommandBars("Reviewing").Controls(11).Enabled = False    '�������ϵġ��޶�����ť
'    For i = 1 To objWordDocument.CommandBars("Menu Bar").Controls(6).CommandBar.Controls.Count
'        If Left(objWordDocument.CommandBars("Menu Bar").Controls(6).CommandBar.Controls(i).Caption, 2) = "�޶�" Then
'            objWordDocument.CommandBars("Menu Bar").Controls(6).CommandBar.Controls(i).Enabled = False  '���߲˵��ϵġ��޶�������
'        End If
'    Next
    '����ҳü�༭״̬��Ȼ���˳��༭״̬���Խ��ҳü�ϵĺ��ߵ���ʾ����
    objWordDocument.ActiveWindow.ActivePane.View.SeekView = 9       'wdSeekCurrentPageHeader
    objWordDocument.ActiveWindow.ActivePane.View.SeekView = 0       'wdSeekMainDocument
'    End
'    Dim X As Object
'    For j = 1 To objWordDocument.CommandBars.Count
'        Set X = objWordDocument.CommandBars(j)
'        For i = 1 To X.Controls.Count
'            If left(X.Controls(i).Caption, 2) = "��ӡ" Then
'                X.Controls(i).Visible = False
'            End If
'        Next
'    Next

    objWord.Visible = True
    
    On Error GoTo errHandler

    objWordDocument.Activate
    objWord.Activate
    
    'objWordDocument.Close
'    objWord.Quit
    
    If paraReadOnly Then
        If objWordDocument.Range.Fields.Count = 0 Then objWordDocument.Protect 3, , "cdc"     '����ĵ�Ϊֻ�����������ĵ����������û��޸�
        objWordDocument.Saved = True
    End If
    
         '��ʾ�޶��ۼ�
        'objWordDocument.ShowRevisions = True
        objWordDocument.TrackRevisions = True
        objWordDocument.Saved = True     '�����ڴ˴����ã�����Saved=false
'    mblnֻ�� = paraReadOnly
    
    '���ݶ���������Ա㱣���ļ������ݿ⡣
    On Error Resume Next
    If Not paraReadOnly Then
        If lintRepID = 0 Then
            objWord.Run "subStart", paraParent, -1, para������, lstr��Ʒ����
        Else
            objWord.Run "subStart", paraParent, lintRepID, para������, lstr��Ʒ����
        End If
    Else
        If lintRepID = 0 Then
            objWord.Run "subStart", Nothing, -1, para������, lstr��Ʒ����
        Else
            objWord.Run "subStart", Nothing, lintRepID, para������, lstr��Ʒ����
        End If
    End If
    If Err.Number = 450 Then     '�������ԣ�˵����ģ��û��������Ʒ����Ĳ���
        If Not paraReadOnly Then
            If lintRepID = 0 Then
                objWord.Run "subStart", paraParent, -1, para������
            Else
                objWord.Run "subStart", paraParent, lintRepID, para������
            End If
        Else
            If lintRepID = 0 Then
                objWord.Run "subStart", Nothing, -1, para������
            Else
                objWord.Run "subStart", Nothing, lintRepID, para������
            End If
        End If
    End If
    If Err.Number = 438 Then
        MsgBox "�ñ����ģ��û�а��չ涨��д�����subStart���������޷����浽���ݿ��", vbOKOnly + vbCritical, "ϵͳ��ʾ"
    End If
    
    Kill "c:\���鱨��*.bmp"
    Exit Sub
errHandler:
    If Err = 3001 Then
        MsgBox "û�������ݿ����ҵ���Word����ľ����ļ��������Ǳ���ñ��浽ϵͳ��ʱ���緢�����ϣ�����ϵͳ��ɾ���ñ�����Ϣ������¼��ñ��档", vbInformation, "ϵͳ��ʾ"
    Else
        sfsub������ "���沿��", "mod�������", "sub�༭word�ĵ�", Err.Number, Err.Description, True
    End If
    Exit Sub
    Resume
End Sub


Public Function func������(ByVal paraErrNumber As Long, ByVal paraErrDes As String) As String
    Select Case paraErrNumber
    Case 6
        func������ = "�������ݹ����ѳ���ϵͳ�涨��С��"
    Case -2147217833
        func������ = "�������ݹ���������󣩣��ѳ���ϵͳ�涨���ȣ����С����"
    Case -2147217913
        func������ = "���ڸ�ʽ�Ƿ���"
    Case -2147217873 '��������ڡ�
        func������ = "ϵͳ�������������Ϊ��" & Chr(13) & Chr(10) & "(1) �����ڱ���������漰�������Ϣ�ѱ���ɾ����" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "����취��" & Chr(13) & Chr(10) & "(1) ���˳���ҵ��������棬���½��롣"
    Case 94 '��Чʹ��Null��
        func������ = "ʹ�õ��ֵ����ͨ���ֵ�������ɾ���ˣ�ϵͳ�޷��ټ���������������ϵͳ����Ա�ָ��ֵ����ݡ���ע�⣬��Ҫ���ɾ���ֵ��"
    Case 336, 337, 338, 429, 430
        func������ = "ϵͳ�������𻵣����Ѷ�ʧ����ϵͳ�޷����������С�" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "����취��" & Chr(13) & Chr(10) & "(1) ���˳�ϵͳ�������°�װϵͳ��"
    Case 440 '�ⲿ����������Զ�����
        func������ = "ϵͳ������������ֹ���С�" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "����취��" & Chr(13) & Chr(10) & "(1) ���˳�ϵͳ�����½��롣"
    Case 91 '����û�г�ʼ���ɹ���
        func������ = "��Ϊ����������ϵͳ��������ʱ�޷���������ĳ�ʼ�������˳����ܽ��棬�����½��빦�ܽ��档" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "����취��" & Chr(13) & Chr(10) & "(1) ���˳�ϵͳ�����½��롣"
    Case 5
        func������ = "��Ϊ�����жϣ�����������ϵͳ�޷��������С�" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "����취��" & Chr(13) & Chr(10) & "(1) ���˳�ϵͳ�����½��롣"
    Case Else
        func������ = paraErrDes
    End Select
End Function


