VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F48D4DEC-1198-11D5-91BE-0050BA06B70C}#5.9#0"; "¼��ؼ�.ocx"
Begin VB.Form frm�����շ���Ŀ 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "�����շ���Ŀ"
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   9375
   ClipControls    =   0   'False
   Icon            =   "frm�����շ���Ŀ.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   9375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   6165
      Left            =   0
      TabIndex        =   8
      Top             =   840
      Width           =   9255
      Begin VB.TextBox ctxtInput 
         Height          =   375
         Left            =   1920
         TabIndex        =   14
         Top             =   5640
         Width           =   2655
      End
      Begin VB.OptionButton coptFind 
         Caption         =   "�����Ƿ�����"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   5640
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   5805
         Left            =   4800
         TabIndex        =   9
         Top             =   120
         Width           =   4365
         Begin VB.ComboBox ccmb�շ���ĿƱ������ 
            Height          =   300
            ItemData        =   "frm�����շ���Ŀ.frx":0442
            Left            =   1380
            List            =   "frm�����շ���Ŀ.frx":0444
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   4260
            Width           =   2835
         End
         Begin ¼��ؼ�.ctlInputBox cinb�շ���Ŀ��� 
            Height          =   360
            Left            =   240
            TabIndex        =   7
            Top             =   255
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   635
            BackColor       =   16777215
            ForeColor       =   0
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            LeftOfTextbox   =   1120
            Text            =   ""
            Label           =   "�շ���Ŀ���"
            Enabled         =   0   'False
            ����            =   ""
            ����            =   0
            ����������ֵ  =   0   'False
            ���������Сֵ  =   0   'False
            �����ѡ        =   0   'False
         End
         Begin ¼��ؼ�.ctlInputBox cinb�շ���Ŀ���� 
            Height          =   360
            Left            =   255
            TabIndex        =   0
            Top             =   827
            Width           =   3960
            _ExtentX        =   6985
            _ExtentY        =   635
            BackColor       =   16777215
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            LeftOfTextbox   =   1120
            Text            =   ""
            Label           =   "�շ���Ŀ����"
            ����            =   ""
            ����            =   20
            ����������ֵ  =   0   'False
            ���������Сֵ  =   0   'False
            �����ѡ        =   0   'False
         End
         Begin ¼��ؼ�.ctlInputBox cinb�շ���Ŀ���Ƿ� 
            Height          =   360
            Left            =   795
            TabIndex        =   1
            Top             =   1399
            Width           =   3420
            _ExtentX        =   6033
            _ExtentY        =   635
            BackColor       =   16777215
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            LeftOfTextbox   =   580
            Text            =   ""
            Label           =   "���Ƿ�"
            ����            =   ""
            ����            =   20
            ����������ֵ  =   0   'False
            ���������Сֵ  =   0   'False
            �����ѡ        =   0   'False
         End
         Begin ¼��ؼ�.ctlInputBox cinb�շ���Ŀ���� 
            Height          =   360
            Left            =   990
            TabIndex        =   2
            Top             =   1965
            Width           =   3240
            _ExtentX        =   5715
            _ExtentY        =   635
            BackColor       =   16777215
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            LeftOfTextbox   =   400
            Text            =   ""
            Label           =   "����"
            ����            =   ""
            ����            =   5
            ����������ֵ  =   0   'False
            ���������Сֵ  =   0   'False
            �����ѡ        =   0   'False
         End
         Begin ¼��ؼ�.ctlInputBox cinb�շ���Ŀ��С���� 
            Height          =   360
            Left            =   615
            TabIndex        =   3
            Top             =   2550
            Width           =   3600
            _ExtentX        =   6350
            _ExtentY        =   635
            BackColor       =   16777215
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            LeftOfTextbox   =   760
            Text            =   ""
            Label           =   "��С����"
            ����            =   ""
            ����            =   5
            ����������ֵ  =   0   'False
            ���������Сֵ  =   0   'False
            �����ѡ        =   0   'False
         End
         Begin ¼��ؼ�.ctlInputBox cinb�շ���Ŀ��󵥼� 
            Height          =   360
            Left            =   615
            TabIndex        =   4
            Top             =   3115
            Width           =   3600
            _ExtentX        =   6350
            _ExtentY        =   635
            BackColor       =   16777215
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            LeftOfTextbox   =   760
            Text            =   ""
            Label           =   "��󵥼�"
            ����            =   ""
            ����            =   5
            ����������ֵ  =   0   'False
            ���������Сֵ  =   0   'False
            �����ѡ        =   0   'False
         End
         Begin ¼��ؼ�.ctlInputBox cinb�շ���Ŀ������λ 
            Height          =   360
            Left            =   615
            TabIndex        =   5
            Top             =   3690
            Width           =   3600
            _ExtentX        =   6350
            _ExtentY        =   635
            BackColor       =   16777215
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            LeftOfTextbox   =   760
            Text            =   ""
            Label           =   "������λ"
            ����            =   ""
            ����            =   5
            ����������ֵ  =   0   'False
            ���������Сֵ  =   0   'False
            �����ѡ        =   0   'False
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "����Ʊ������"
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   4320
            Width           =   1095
         End
      End
      Begin MSComctlLib.TreeView ctvwMain 
         Height          =   5295
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   9340
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   353
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   6
         Appearance      =   1
      End
   End
   Begin MSComctlLib.ImageList cimg��ťͼ�� 
      Left            =   5760
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar ctlb���� 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "cimg��ťͼ��"
      _Version        =   393216
   End
End
Attribute VB_Name = "frm�����շ���Ŀ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public pblnInUse As Boolean

Private WithEvents mobjGUI As cls����ͨ�ö���
Attribute mobjGUI.VB_VarHelpID = -1

Public pint��Ŀ���� As Long

'�Զ��������Ƿ�.
Private Sub cinb�շ���Ŀ����_Change()
    Dim lstrTemp As String
    Dim lobj���Ƿ� As Object
    
    On Error Resume Next
    Set lobj���Ƿ� = CreateObject("���Ƿ�.cls���Ƿ�")
    lstrTemp = lobj���Ƿ�.guf_GetFirstLetter(cinb�շ���Ŀ����.Text)
    lstrTemp = Left(lstrTemp, 20)
    cinb�շ���Ŀ���Ƿ�.Text = lstrTemp
End Sub

Private Sub ctvwMain_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim lintcount As Integer
    Dim lrsd�շ���Ŀ As Object
    Dim lrsdƱ������ As Object
    On Error GoTo errhandler
    
    If Node.Key <> "s" Then
        Set lrsd�շ���Ŀ = dafuncGetData("select * from �շѹ���_�շ���Ŀ�ֵ�� where �շ���Ŀ���='" & Right(LTrim(RTrim(Node.Key)), Len(LTrim(RTrim(Node.Key))) - 1) & "'")
        If Not lrsd�շ���Ŀ.EOF Then
            cinb�շ���Ŀ���.Text = lrsd�շ���Ŀ.Fields("�շ���Ŀ���").Value
            cinb�շ���Ŀ����.Text = lrsd�շ���Ŀ.Fields("�շ���Ŀ����").Value
            cinb�շ���Ŀ����.Text = lrsd�շ���Ŀ.Fields("����").Value
            cinb�շ���Ŀ������λ.Text = IIf(IsNull(lrsd�շ���Ŀ.Fields("������λ").Value), "", lrsd�շ���Ŀ.Fields("������λ").Value)
            cinb�շ���Ŀ���Ƿ�.Text = IIf(IsNull(lrsd�շ���Ŀ.Fields("���Ƿ�").Value), "", lrsd�շ���Ŀ.Fields("���Ƿ�").Value)
            cinb�շ���Ŀ��С����.Text = lrsd�շ���Ŀ.Fields("��С����").Value
            cinb�շ���Ŀ��󵥼�.Text = lrsd�շ���Ŀ.Fields("��󵥼�").Value
            
            Set lrsdƱ������ = dafuncGetData("select * from �շѹ���_Ʊ�������ֵ���ͼ")
            If (lrsdƱ������.RecordCount > 0) Then
                lrsdƱ������.MoveFirst
                Do While (Not lrsdƱ������.EOF)
                    If lrsdƱ������("InnerID").Value = Val(lrsd�շ���Ŀ("Ʊ�����ͱ��").Value) Then
                        Exit Do
                    Else
                        lrsdƱ������.MoveNext
                    End If
                Loop
                If lrsdƱ������.EOF Then
                    If (Len(LTrim(RTrim(ctvwMain.SelectedItem.Key))) - 1) / 3 = pint��Ŀ���� Then
                        MsgBox "����޸���Ʊ�������ֵ��������¼�����Ŀ��Ʊ�����ͣ�����ĿƱ������������!", vbExclamation, "Ʊ����������"
                    End If
                    Exit Sub
                End If
                ccmb�շ���ĿƱ������.Text = lrsdƱ������("����").Value
            End If

        End If
        cinb�շ���Ŀ���.SetFocus
    End If
    
    Exit Sub
errhandler:
    sfsub������ "�շѽ��沿��", "frm�����շ���Ŀ", "ctvwMain_MouseDown", Err.Number, Err.Description, False
End Sub

Private Sub ctxtInput_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo errhandler
    Dim i As Long
    If KeyCode = 13 Then
        If Not ctvwMain.SelectedItem Is Nothing Then
            ctvwMain_NodeClick ctvwMain.SelectedItem
        End If
    Else
        '��λ��
        Dim lnodeParent As Node
        Dim lNode As Node
        Dim lstrTemp As String
        
        If ctvwMain.SelectedItem.Children = 0 Then
            Set lnodeParent = ctvwMain.SelectedItem.Parent
        Else
            Set lnodeParent = ctvwMain.SelectedItem
        End If
        lnodeParent.Selected = True
        If ctxtInput.Text <> "" Then
            If lnodeParent.Children > 0 Then
                Set lNode = lnodeParent.Child
                For i = 1 To lnodeParent.Children
                    lstrTemp = Right(lNode.Text, Len(lNode.Text) - InStr(lNode.Text, " "))
                    If UCase(Left(lstrTemp, Len(ctxtInput.Text))) = UCase(ctxtInput.Text) Then
                        lNode.Selected = True
                        Exit For
                    Else
                        Set lNode = lNode.Next
                    End If
                Next
            End If
        End If
    End If
    Exit Sub
errhandler:
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 39 Then
        KeyAscii = 0
    ElseIf KeyAscii = 13 Then
        If ActiveControl = ctxtInput Then
        Else
            SendKeys Chr(vbKeyTab)
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim lobjRec As Object
    
    On Error GoTo errhandler
    
    If pblnInUse = True Then Exit Sub
    
    pblnInUse = True

    '��ʼ��������
    Dim lcol��������ť As Collection
    Set lcol��������ť = New Collection
    Set mobjGUI = New cls����ͨ�ö���
    Set mobjGUI.Form = Me
    Set mobjGUI.c������ = ctlb����

    lcol��������ť.Add "���"
    lcol��������ť.Add "ɾ��"
    lcol��������ť.Add "|"
    lcol��������ť.Add "����"
    lcol��������ť.Add "|"
    lcol��������ť.Add "�˳�"
    
    mobjGUI.subInitialize lcol��������ť, ""
    
    
    '��ʼ��Ʊ������
    Set lobjRec = dafuncGetData("select * from �շѹ���_Ʊ�������ֵ���ͼ")
    If (Not lobjRec.EOF) And (Not lobjRec.BOF) Then
        Do While (Not lobjRec.EOF)
            ccmb�շ���ĿƱ������.AddItem lobjRec.Fields("����").Value
            ccmb�շ���ĿƱ������.ItemData(ccmb�շ���ĿƱ������.NewIndex) = lobjRec.Fields("innerId").Value
            lobjRec.MoveNext
        Loop
        lobjRec.MoveFirst
    End If
    If ccmb�շ���ĿƱ������.ListCount > 0 Then
        ccmb�շ���ĿƱ������.ListIndex = 0
    End If
    
    '��ʼ���շ���Ŀ����
    Dim lnodParent As Node
    Dim lint���� As Long
    Dim lstrKey As String
    
    pint��Ŀ���� = Val(pobj�շѹ���.ҵ������("��Ŀ����"))
    If pint��Ŀ���� = 0 Then pint��Ŀ���� = 2
    
    Set lnodParent = ctvwMain.Nodes.Add(, , "s", "�շ���Ŀ")
    For lint���� = 1 To pint��Ŀ����
        Set lobjRec = dafuncGetData("select * from �շѹ���_�շ���Ŀ�ֵ�� where Len(�շ���Ŀ���) =" & lint���� * 3 & " order by ���Ƿ�")
        Do While (Not lobjRec.EOF)
            lstrKey = "s" & lobjRec("�շ���Ŀ���").Value
            ctvwMain.Nodes.Add "s" & Mid(lstrKey, 2, ((lint���� - 1) * 3)), tvwChild, lstrKey, lobjRec("�շ���Ŀ����").Value & " " & IIf(IsNull(lobjRec("���Ƿ�")), "", lobjRec("���Ƿ�"))
            lobjRec.MoveNext
        Loop
    Next
    ctvwMain.Nodes(1).Expanded = True
    Exit Sub
errhandler:
    sfsub������ "�շѽ��沿��", "frm�����շ���Ŀ", "Form_Load", Err.Number, Err.Description, False
    Exit Sub
    Resume

End Sub

Private Sub Form_Unload(Cancel As Integer)
    pblnInUse = False
End Sub

Private Sub mobjGUI_BeforeOperate(ByVal Operate As String, Cancel As Boolean)
    On Error GoTo errhandler
    Select Case Operate
    Case "���"
        Cancel = True
        cinb�շ���Ŀ���.Text = ""
        cinb�շ���Ŀ����.Text = ""
        cinb�շ���Ŀ����.Text = ""
        cinb�շ���Ŀ������λ.Text = ""
        cinb�շ���Ŀ���Ƿ�.Text = ""
        cinb�շ���Ŀ��󵥼�.Text = ""
        cinb�շ���Ŀ��С����.Text = ""
        cinb�շ���Ŀ����.SetFocus
        
    Case "����"
        Dim lstrParent As String
        Cancel = True
        If ctvwMain.SelectedItem Is Nothing Then
            lstrParent = ""
        ElseIf ctvwMain.SelectedItem.Key = "s" Then
            lstrParent = ""
        Else
            lstrParent = Right(ctvwMain.SelectedItem.Key, Len(ctvwMain.SelectedItem.Key) - 1)
            If Len(lstrParent) = pint��Ŀ���� * 3 Then
                'ѡ�е���ĩ�����롣
                lstrParent = Left(lstrParent, Len(lstrParent) - 3)
            End If
        End If
        
        'У��
        subValidate lstrParent
        
        '�����ݿ�����Ӽ�¼
        Dim lobjItem As Object
        Set lobjItem = CreateObject("�շѶ��󲿼�.cls�շ���Ŀ")
        lobjItem.�շ���Ŀ��� = cinb�շ���Ŀ���.Text
        lobjItem.�շ���Ŀ���� = cinb�շ���Ŀ����.Text
        lobjItem.���� = cinb�շ���Ŀ����.Text
        lobjItem.������λ = cinb�շ���Ŀ������λ.Text
        lobjItem.Ʊ�����ͱ�� = ccmb�շ���ĿƱ������.ItemData(ccmb�շ���ĿƱ������.ListIndex)
        lobjItem.��С���� = cinb�շ���Ŀ��С����.Text
        lobjItem.��󵥼� = cinb�շ���Ŀ��󵥼�.Text
        lobjItem.���Ƿ� = cinb�շ���Ŀ���Ƿ�.Text
        lobjItem.sub���� lstrParent
        
        
        '��ӳɹ������ӽڵ�.
        If cinb�շ���Ŀ���.Text = "" Then
            Call ctvwMain.Nodes.Add("s" & lstrParent, tvwChild, "s" & lobjItem.�շ���Ŀ���, cinb�շ���Ŀ����.Text)
        End If
        '�Զ�������
        mobjGUI_BeforeOperate "���", True
    
    Case "ɾ��"
        Dim lstrKey As String
        Cancel = True
        If ctvwMain.SelectedItem Is Nothing Then
            Err.Raise 6666, , "��ѡ��Ҫɾ�����շ���Ŀ��"
        ElseIf ctvwMain.SelectedItem.Key = "s" Then
            Err.Raise 6666, , "��ѡ��Ҫɾ�����շ���Ŀ��"
        ElseIf ctvwMain.SelectedItem.Children > 0 Then
            Err.Raise 6666, , "������¼���ʼɾ����"
        Else
            If MsgBox("��ȷ��Ҫɾ���շ���Ŀ��" & ctvwMain.SelectedItem.Text & "����", vbYesNo + vbQuestion, "ϵͳѯ��") = vbYes Then
                lstrKey = Right(ctvwMain.SelectedItem.Key, Len(ctvwMain.SelectedItem.Key) - 1)
                pobj�շѹ���.subɾ���շ���Ŀ (lstrKey)
                
                ctvwMain.Nodes.Remove (ctvwMain.SelectedItem.Key)
            End If
        End If
    
    End Select
    Exit Sub
errhandler:
    sfsub������ "�շѽ��沿��", "frm�����շ���Ŀ", "mobjGUI_BeforeOperate", Err.Number, Err.Description, False
    Exit Sub
    Resume

End Sub

Private Sub subValidate(ByVal paraParent As String)
    
    If Trim(cinb�շ���Ŀ����.Text) = "" Then Err.Raise 6666, , "�շ���Ŀ���Ʋ���Ϊ�գ���¼�룡"
    If ccmb�շ���ĿƱ������.ListIndex = -1 Then Err.Raise 6666, , "����ѡ��Ʊ�����ͣ�"
    If Len(paraParent) = (pint��Ŀ���� - 1) * 3 Then
        'ĩ����Ŀ���������뵥�ۡ�
        If cinb�շ���Ŀ����.Text = "" Then Err.Raise 6666, , "���һ�����۲���Ϊ�գ����޸ģ�"
        If cinb�շ���Ŀ����.Text = 0 Then Err.Raise 6666, , "���һ�����۲���Ϊ�㣡"
        
        If IsNumeric(cinb�շ���Ŀ����.Text) And IsNumeric(cinb�շ���Ŀ��󵥼�.Text) And IsNumeric(cinb�շ���Ŀ��С����.Text) Then
               If CDbl(cinb�շ���Ŀ����.Text) < 0 Or CDbl(cinb�շ���Ŀ��󵥼�.Text) < 0 Or CDbl(cinb�շ���Ŀ��С����.Text) < 0 Then
                   Err.Raise 6666, , "���ۡ���С���ۡ���󵥼۱�������㣬���޸ģ�"
               End If
        Else
            Err.Raise 6666, , "���ۡ���С���ۡ���󵥼۱���Ϊ��ֵ�����޸ģ�"
        End If
        If CDbl(cinb�շ���Ŀ����.Text) > CDbl(cinb�շ���Ŀ��󵥼�.Text) Or CDbl(cinb�շ���Ŀ����.Text) < CDbl(cinb�շ���Ŀ��С����.Text) Then
            Err.Raise 6666, , "���۱�������С���ۺ���󵥼�֮�䣬���޸ģ�"
        End If
             
        If cinb�շ���Ŀ��󵥼�.Text = "" And cinb�շ���Ŀ��С����.Text = "" Then
            cinb�շ���Ŀ��󵥼�.Text = cinb�շ���Ŀ����.Text
            cinb�շ���Ŀ��С����.Text = cinb�շ���Ŀ����.Text
        ElseIf cinb�շ���Ŀ��󵥼�.Text = "" And cinb�շ���Ŀ��С����.Text <> "" Then
            If CDbl(cinb�շ���Ŀ��С����.Text) < CDbl(cinb�շ���Ŀ����.Text) Then
                cinb�շ���Ŀ��󵥼�.Text = cinb�շ���Ŀ����.Text
            End If
        ElseIf cinb�շ���Ŀ��󵥼�.Text <> "" And cinb�շ���Ŀ��С����.Text = "" Then
            If CDbl(cinb�շ���Ŀ��󵥼�.Text) > CDbl(cinb�շ���Ŀ����.Text) Then
                cinb�շ���Ŀ��С����.Text = cinb�շ���Ŀ����.Text
            End If
        End If
             
        If cinb�շ���Ŀ������λ.Text = "" Or IsNull(cinb�շ���Ŀ������λ.Text) Then Err.Raise 6666, , "�����������λ��"
             
    Else
        If Len(LTrim(RTrim(cinb�շ���Ŀ���.Text))) / 3 > pint��Ŀ���� Then
            Err.Raise 6666, , "�շ���Ŀ�����������ƣ�����Ŀ������Ч"
        Else
            cinb�շ���Ŀ����.Text = "0"
            cinb�շ���Ŀ��󵥼�.Text = "0"
            cinb�շ���Ŀ��С����.Text = "0"
        End If
    End If

End Sub
