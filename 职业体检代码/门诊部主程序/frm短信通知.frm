VERSION 5.00
Object = "{15138B51-7EB6-11D0-9BB7-0000C0F04C96}#1.0#0"; "SSLstBar.ocx"
Begin VB.Form frm����֪ͨ 
   BackColor       =   &H00B5D0D7&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3570
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   1980
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frm����֪ͨ.frx":0000
   ScaleHeight     =   3570
   ScaleWidth      =   1980
   StartUpPosition =   3  '����ȱʡ
   Begin Listbar.SSListBar cbarOper 
      Height          =   3645
      Left            =   -30
      TabIndex        =   1
      Top             =   -30
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   6429
      _Version        =   65536
      BackColor       =   16777215
      BorderStyle     =   3
      CaptionBackColor=   16757864
      CaptionForeColor=   16777215
      PictureBackgroundUseMask=   -1  'True
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TagVariant      =   ""
      OLEDragMode     =   1
      OLEDropMode     =   2
      Style           =   1
      IconsSmallCount =   1
      Image(1).Index  =   1
      Image(1).Picture=   "frm����֪ͨ.frx":321BA
      Groups(1).ItemCount=   1
      Groups(1).BackColor=   16777215
      Groups(1).Style =   1
      Groups(1).PictureBackgroundUseMask=   -1  'True
      Groups(1).CurrentGroup=   -1  'True
      BeginProperty Groups(1).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Groups(1).Caption=   "���ڹ���"
      Groups(1).ListItems(1).ForeColorSource=   1
      Groups(1).ListItems(1).ForeColor=   0
      Groups(1).ListItems(1).Text=   "ListItem"
      Groups(1).ListItems(1).IconLarge=   ""
      Groups(1).ListItems(1).IconSmall=   1
      Groups(1).ListItems(1).TagVariant=   ""
   End
   Begin VB.Label clbl�ֵ� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��¼���Ǽ���Ϣ"
      Height          =   180
      Index           =   0
      Left            =   165
      TabIndex        =   0
      Top             =   75
      Visible         =   0   'False
      Width           =   1815
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frm����֪ͨ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public pobjParent As frmMain
Public Sub subLoad(para���� As String, paraobj�� As Object, paraobj���� As Object)
    Dim i As Integer, j As Integer
    
    '����ͼ���
    Dim lstrFileName As String, lobjPic As StdPicture
    
'    lstrFileName = Dir(App.Path & "\image\*.ico")
'    Do While lstrFileName <> ""
'        Set lobjPic = LoadPicture(App.Path & "\image\" & lstrFileName)
'        cbarOper.IconsSmall.add , Left(lstrFileName, InStr(lstrFileName, ".") - 1), lobjPic
'        lstrFileName = Dir
'    Loop
    
    Dim ii As Integer
    Dim lobjRec As Object
    
    ii = 1
    paraobj��.Filter = ""
    paraobj��.Filter = "��������" & " ='" & para���� & "' "
    cbarOper.Groups(1).ListItems.Clear
    For i = 1 To paraobj��.RecordCount
        '�޸ģ�2003-7-9������жϵ�ǰ��������ҵ�����Ƿ��ڼ��ܹ���ɷ�Χ�ڡ�
        paraobj����.Filter = ""
        paraobj����.Filter = "��������" & "='" & paraobj��.Fields("��������") & "'"
        If paraobj����.RecordCount > 0 Then
            If pstr��ϵͳ��� = "" Or InStr(pstr��ϵͳ���, paraobj����.Fields("ҵ����") & ",") > 0 Then
                cbarOper.Groups(1).ListItems.add ii, "" & paraobj��.Fields("��������"), paraobj��("��������")
                cbarOper.Groups(1).ListItems(ii).Text = paraobj��("��������")
                cbarOper.Groups(1).ListItems(ii).IconSmall = 1      'paraobj��("��������")
                cbarOper.Groups(1).ListItems(ii).ForeColorSource = ssUseListItem
                cbarOper.Groups(1).ListItems(ii).ForeColor = &H0
                
                'frm�����б�.subAddOperation paraobj��("��������"), paraobj��.Fields("��������")
                If Not sffunc�жϼ��ϼ�ֵ�Ƿ����(pcol�ֵ伯, paraobj����.Fields("ҵ����").Value) Then
                    '�жϸ�ҵ���Ƿ��в��������ֵ䡣
                    Set lobjRec = dafuncGetData("select * from ϵͳ����_�ֵ�_�ֵ���б� where ҵ����='" & paraobj����.Fields("ҵ����").Value & "' and ����='������'")
                    If lobjRec.RecordCount > 0 Then
                        pcol�ֵ伯.add paraobj����.Fields("ҵ����").Value, paraobj����.Fields("ҵ����").Value
                    End If
                End If
            
                ii = ii + 1
            
            End If
        End If
        paraobj��.MoveNext
    Next i

    Exit Sub
End Sub
Private Sub cbarOper_ListItemClick(ByVal ItemClicked As Listbar.SSListItem)
    Me.Hide
    pobjParent.sub�������� ItemClicked.Key
    
End Sub


Private Sub Form_LostFocus()
    Me.Hide
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    Set pobjParent = Nothing
End Sub
