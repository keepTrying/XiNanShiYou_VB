VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmfresult 
   Caption         =   "Ѫ����������"
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9825
   LinkTopic       =   "Form1"
   ScaleHeight     =   4500
   ScaleWidth      =   9825
   StartUpPosition =   1  '����������
   Begin VB.Frame frmuploadFresult 
      Caption         =   "Ѫ������¼��"
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   8175
      Begin VB.CommandButton Commѡ�� 
         Caption         =   "ѡ���ļ�"
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
         Caption         =   "��ʼ����"
         Height          =   375
         Left            =   5520
         TabIndex        =   1
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Labelҽʦ��� 
         Caption         =   "ҽʦ���"
         Height          =   255
         Left            =   1320
         TabIndex        =   8
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "���ҽʦ��"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   480
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Label3 
         Caption         =   "�ϴ����ȣ�"
         Height          =   255
         Left            =   720
         TabIndex        =   4
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Labelstate 
         Caption         =   "��δ�����ϴ�������"
         Height          =   255
         Left            =   1560
         TabIndex        =   3
         Top             =   1680
         Width           =   3855
      End
      Begin VB.Label Label1 
         Caption         =   "�����ļ�״̬��"
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
' Commѡ��.Visible = True
' Else
'' Label4.Visible = False
' Text1.Visible = False
' Commѡ��.Visible = False
' End If
'End Sub

Private Sub Command1_Click()
 Dim lstr������� As String
 
Dim name As String
name = Text1.Text

Dim oname As String
Dim nname As String
oname = "d:\fresult\" + name + ".txt"
nname = "d:\fresult\" + name + ".xls"
'��������txt�ĳ�xls
'Name oname As nname
'��Dir�������ж��ļ��Ƿ����
If Dir(nname) = "" Then
    If Dir(oname) = "" Then  '�ж�text�ļ��Ƿ����   2015-9-17
    MsgBox "�ļ������ڻ����ļ�ѡ�����", , "��Ϣ��ʾ"
    Exit Sub
    Else
    Name oname As nname
    End If
End If


Dim xlsApp As Excel.Application     '�����������
Set xlsApp = New Excel.Application        'ʵ��������
xlsApp.Visible = False      'ʹExcel���ز��ɼ�
xlsApp.Workbooks.Open (nname)   '��EXCEL�ļ�

'Dim mg As Range
Labelstate.Caption = "���ڼ�����ݽṹ��"
Dim totals As Long
'Sheets1.name = Replace(ThisWorkbook.name, ".xls", "")
totals = xlsApp.ActiveWorkbook.Sheets(name).UsedRange.rows.Count
Dim prmax As Long
prmax = 0
Dim ii As Long
 For ii = 1 To totals
    If xlsApp.ActiveWorkbook.Sheets(name).Cells(ii, 1) <> "" And xlsApp.ActiveWorkbook.Sheets(name).Cells(ii + 25, 2) = "����ѪС�����|P-LCR" Then
    prmax = prmax + 1
    End If
    Labelstate.Caption = "���ڼ�����ݽṹ,����" + Str(ii) + "/" + Str(totals)
Next ii

'��ʼ��������
ProgressBar1.Min = 0
ProgressBar1.Max = prmax
 Labelstate.Caption = "��ʼ�ϴ�..."
 
 Dim �����Ŀ���(2) As String
Dim i As Long
 For i = 1 To totals
    If xlsApp.ActiveWorkbook.Sheets(name).Cells(i, 1) <> "" And xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 25, 2) = "����ѪС�����|P-LCR" Then
'    If xlsApp.ActiveWorkbook.Sheets(name).Cells(i, 1) <> "" And xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 25, 2) = "ѪС��ֲ����|PDW" Then
    Dim SysNo As String
    '��ʾ����֤�û���ԭʼ����һ�����ı��ĵ�����֤ϵͳ���������һ��������ϵͳ�����ʮ�����Ƶ�excel�ļ� 2015-9-22 by Ĳ��
    
'    SysNo = "00" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i, 6).Value  'ԭ����txt�ĵ�ת���ɵı��ϵͳ����ڵ�����
'    SysNo = "00" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i, 12).Value    '���ڵ�txt�ĵ�ת���ɵı��ϵͳ����ڵ�12��     2016-1-6 by Ĳ��
    SysNo = "0" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i, 12).Value    '���ڵ�txt�ĵ�ת���ɵı��ϵͳ����ڵ�12�в���ϵͳ���ǰ�����˸�0��2016-1-4���浼��Ĳ��٣�   2016-1-6 by Ĳ��
    
'   lstr������� = pobjҵ�����.func��ȡ�������(cgrdInput.TextMatrix(Row, 0), cgrdInput.TextMatrix(Row, 2))
    dafuncGetData ("update ְҵ�����_�����Ϣ_Ѫ���滯��� set �����='" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 2, 3) & "', ���ҽʦ='" & Labelҽʦ���.Caption & "',  ��дʱ��='" & Now & "',�������='" & pobjҵ�����.func��ȡ�������("04021", xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 2, 3)) & "' where ϵͳ���='" & SysNo & "' and �����Ŀ='04021' ")
    dafuncGetData ("update ְҵ�����_�����Ϣ_Ѫ���滯��� set �����='" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 3, 3) & "', ���ҽʦ='" & Labelҽʦ���.Caption & "',  ��дʱ��='" & Now & "',�������='" & pobjҵ�����.func��ȡ�������("04022", xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 3, 3)) & "' where ϵͳ���='" & SysNo & "' and �����Ŀ='04022' ")
    dafuncGetData ("update ְҵ�����_�����Ϣ_Ѫ���滯��� set �����='" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 4, 3) & "', ���ҽʦ='" & Labelҽʦ���.Caption & "',  ��дʱ��='" & Now & "',�������='" & pobjҵ�����.func��ȡ�������("04023", xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 4, 3)) & "' where ϵͳ���='" & SysNo & "' and �����Ŀ='04023' ")
    dafuncGetData ("update ְҵ�����_�����Ϣ_Ѫ���滯��� set �����='" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 5, 3) & "', ���ҽʦ='" & Labelҽʦ���.Caption & "',  ��дʱ��='" & Now & "',�������='" & pobjҵ�����.func��ȡ�������("04001", xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 5, 3)) & "' where ϵͳ���='" & SysNo & "' and �����Ŀ='04001' ")
    dafuncGetData ("update ְҵ�����_�����Ϣ_Ѫ���滯��� set �����='" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 6, 3) & "', ���ҽʦ='" & Labelҽʦ���.Caption & "',  ��дʱ��='" & Now & "',�������='" & pobjҵ�����.func��ȡ�������("04024", xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 6, 3)) & "' where ϵͳ���='" & SysNo & "' and �����Ŀ='04024' ")
    
    dafuncGetData ("update ְҵ�����_�����Ϣ_Ѫ���滯��� set �����='" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 7, 3) & "', ���ҽʦ='" & Labelҽʦ���.Caption & "',  ��дʱ��='" & Now & "',�������='" & pobjҵ�����.func��ȡ�������("04002", xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 7, 3)) & "' where ϵͳ���='" & SysNo & "' and �����Ŀ='04002' ")
    dafuncGetData ("update ְҵ�����_�����Ϣ_Ѫ���滯��� set �����='" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 8, 3) & "', ���ҽʦ='" & Labelҽʦ���.Caption & "',  ��дʱ��='" & Now & "',�������='" & pobjҵ�����.func��ȡ�������("04003", xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 8, 3)) & "' where ϵͳ���='" & SysNo & "' and �����Ŀ='04003' ")
    dafuncGetData ("update ְҵ�����_�����Ϣ_Ѫ���滯��� set �����='" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 9, 3) & "', ���ҽʦ='" & Labelҽʦ���.Caption & "',  ��дʱ��='" & Now & "',�������='" & pobjҵ�����.func��ȡ�������("04004", xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 9, 3)) & "' where ϵͳ���='" & SysNo & "' and �����Ŀ='04004' ")
    dafuncGetData ("update ְҵ�����_�����Ϣ_Ѫ���滯��� set �����='" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 10, 3) & "', ���ҽʦ='" & Labelҽʦ���.Caption & "',  ��дʱ��='" & Now & "',�������='" & pobjҵ�����.func��ȡ�������("04005", xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 10, 3)) & "' where ϵͳ���='" & SysNo & "' and �����Ŀ='04005' ")
    dafuncGetData ("update ְҵ�����_�����Ϣ_Ѫ���滯��� set �����='" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 11, 3) & "', ���ҽʦ='" & Labelҽʦ���.Caption & "',  ��дʱ��='" & Now & "',�������='" & pobjҵ�����.func��ȡ�������("04006", xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 11, 3)) & "' where ϵͳ���='" & SysNo & "' and �����Ŀ='04006' ")
    
    dafuncGetData ("update ְҵ�����_�����Ϣ_Ѫ���滯��� set �����='" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 12, 3) & "', ���ҽʦ='" & Labelҽʦ���.Caption & "',  ��дʱ��='" & Now & "',�������='" & pobjҵ�����.func��ȡ�������("04007", xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 12, 3)) & "' where ϵͳ���='" & SysNo & "' and �����Ŀ='04007' ")
    dafuncGetData ("update ְҵ�����_�����Ϣ_Ѫ���滯��� set �����='" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 13, 3) & "', ���ҽʦ='" & Labelҽʦ���.Caption & "',  ��дʱ��='" & Now & "',�������='" & pobjҵ�����.func��ȡ�������("04008", xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 13, 3)) & "' where ϵͳ���='" & SysNo & "' and �����Ŀ='04008' ")
    dafuncGetData ("update ְҵ�����_�����Ϣ_Ѫ���滯��� set �����='" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 14, 3) & "', ���ҽʦ='" & Labelҽʦ���.Caption & "',  ��дʱ��='" & Now & "',�������='" & pobjҵ�����.func��ȡ�������("04009", xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 14, 3)) & "' where ϵͳ���='" & SysNo & "' and �����Ŀ='04009' ")
    dafuncGetData ("update ְҵ�����_�����Ϣ_Ѫ���滯��� set �����='" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 15, 3) & "', ���ҽʦ='" & Labelҽʦ���.Caption & "',  ��дʱ��='" & Now & "',�������='" & pobjҵ�����.func��ȡ�������("04010", xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 15, 3)) & "' where ϵͳ���='" & SysNo & "' and �����Ŀ='04010' ")
    dafuncGetData ("update ְҵ�����_�����Ϣ_Ѫ���滯��� set �����='" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 16, 3) & "', ���ҽʦ='" & Labelҽʦ���.Caption & "',  ��дʱ��='" & Now & "',�������='" & pobjҵ�����.func��ȡ�������("04011", xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 16, 3)) & "' where ϵͳ���='" & SysNo & "' and �����Ŀ='04011' ")
    
    dafuncGetData ("update ְҵ�����_�����Ϣ_Ѫ���滯��� set �����='" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 17, 3) & "', ���ҽʦ='" & Labelҽʦ���.Caption & "',  ��дʱ��='" & Now & "',�������='" & pobjҵ�����.func��ȡ�������("04012", xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 17, 3)) & "' where ϵͳ���='" & SysNo & "' and �����Ŀ='04012' ")
    dafuncGetData ("update ְҵ�����_�����Ϣ_Ѫ���滯��� set �����='" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 18, 3) & "', ���ҽʦ='" & Labelҽʦ���.Caption & "',  ��дʱ��='" & Now & "',�������='" & pobjҵ�����.func��ȡ�������("04013", xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 18, 3)) & "' where ϵͳ���='" & SysNo & "' and �����Ŀ='04013' ")
    dafuncGetData ("update ְҵ�����_�����Ϣ_Ѫ���滯��� set �����='" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 19, 3) & "', ���ҽʦ='" & Labelҽʦ���.Caption & "',  ��дʱ��='" & Now & "',�������='" & pobjҵ�����.func��ȡ�������("04014", xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 19, 3)) & "' where ϵͳ���='" & SysNo & "' and �����Ŀ='04014' ")
    dafuncGetData ("update ְҵ�����_�����Ϣ_Ѫ���滯��� set �����='" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 20, 3) & "', ���ҽʦ='" & Labelҽʦ���.Caption & "',  ��дʱ��='" & Now & "',�������='" & pobjҵ�����.func��ȡ�������("04015", xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 20, 3)) & "' where ϵͳ���='" & SysNo & "' and �����Ŀ='04015' ")
    dafuncGetData ("update ְҵ�����_�����Ϣ_Ѫ���滯��� set �����='" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 21, 3) & "', ���ҽʦ='" & Labelҽʦ���.Caption & "',  ��дʱ��='" & Now & "',�������='" & pobjҵ�����.func��ȡ�������("04016", xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 21, 3)) & "' where ϵͳ���='" & SysNo & "' and �����Ŀ='04016' ")
    
    dafuncGetData ("update ְҵ�����_�����Ϣ_Ѫ���滯��� set �����='" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 22, 3) & "', ���ҽʦ='" & Labelҽʦ���.Caption & "',  ��дʱ��='" & Now & "',�������='" & pobjҵ�����.func��ȡ�������("04017", xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 22, 3)) & "' where ϵͳ���='" & SysNo & "' and �����Ŀ='04017' ")
    dafuncGetData ("update ְҵ�����_�����Ϣ_Ѫ���滯��� set �����='" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 23, 3) & "', ���ҽʦ='" & Labelҽʦ���.Caption & "',  ��дʱ��='" & Now & "',�������='" & pobjҵ�����.func��ȡ�������("04018", xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 23, 3)) & "' where ϵͳ���='" & SysNo & "' and �����Ŀ='04018' ")
    dafuncGetData ("update ְҵ�����_�����Ϣ_Ѫ���滯��� set �����='" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 24, 3) & "', ���ҽʦ='" & Labelҽʦ���.Caption & "',  ��дʱ��='" & Now & "',�������='" & pobjҵ�����.func��ȡ�������("04019", xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 24, 3)) & "' where ϵͳ���='" & SysNo & "' and �����Ŀ='04019' ")
    dafuncGetData ("update ְҵ�����_�����Ϣ_Ѫ���滯��� set �����='" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 25, 3) & "', ���ҽʦ='" & Labelҽʦ���.Caption & "',  ��дʱ��='" & Now & "',�������='" & pobjҵ�����.func��ȡ�������("04020", xlsApp.ActiveWorkbook.Sheets(name).Cells(i + 25, 3)) & "' where ϵͳ���='" & SysNo & "' and �����Ŀ='04020' ")
    
    'Ѫ���������Ϣ�����¼ӵ�ְҵ�����_�����Ϣ_Ѫ���������Ϣ�� 2016-1-13 by Ĳ��
    dafuncGetData ("delete from ְҵ�����_�����Ϣ_Ѫ���������Ϣ�� where ϵͳ���='" & SysNo & "'")   'ɾ��ԭ�м�¼����������Ϊ׼
    dafuncGetData ("insert into ְҵ�����_�����Ϣ_Ѫ���������Ϣ��(ϵͳ���,����,�Ա�,����,��������,����,�걾��,�걾����,�ͼ�ҽ��,������,��������) values('" & SysNo & "','" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i, 5) & "','" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i, 8) & "','" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i, 9) & "','" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i, 7) & "','" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i, 11) & "','" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i, 3) & "','" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i, 17) & "','" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i, 13) & "','" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i, 15) & "','" & xlsApp.ActiveWorkbook.Sheets(name).Cells(i, 14) & "')")
    
    ProgressBar1.Value = ProgressBar1.Value + 1
   End If
Next i
 Labelstate.Caption = "�ϴ����!"
  xlsApp.ActiveWorkbook.Close savechanges:=False    '�رջ������ͬʱ����������ĸ��ġ�
  xlsApp.Quit                                      '�ر�EXCELL
  Set xlsApp = Nothing    '�ͷ���Դ
  MsgBox ("����ɹ���"), vbInformation, "ϵͳ��ʾ"
  Unload Me
End Sub
'ѡ���ļ�   Ϊ�˿�������ѡ��Ҫ������ļ����ӵ� 2015-9-17 by Ĳ��
Private Sub Commѡ��_Click()
    Dim i As Integer
    Dim lstrTmp As String
'    CommonDialog1.ShowOpen
ccdg.Filter = "All Files (*.*)|*.*|Excel file" & _
            "(*.xls)|*.xls|Batch Files (*.bat)|*.bat"
    ccdg.ShowOpen
'    ccdg.FileName = ""
    Text1.Text = CreateObject("Scripting.FileSystemObject").GetBaseName(ccdg.FileName)  'ֻҪ�ļ�������Ҫ·���ͺ�׺�� 2015-9-18
End Sub


