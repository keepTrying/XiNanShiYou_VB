VERSION 5.00
Begin VB.Form frmEyeDraw 
   Caption         =   "��״�廷�漰����ͼ"
   ClientHeight    =   3870
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8070
   LinkTopic       =   "Form1"
   ScaleHeight     =   3870
   ScaleWidth      =   8070
   StartUpPosition =   1  '����������
   Begin VB.CommandButton Command1 
      Caption         =   "���"
      Height          =   375
      Index           =   3
      Left            =   3240
      TabIndex        =   16
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��"
      Height          =   375
      Index           =   2
      Left            =   2520
      TabIndex        =   15
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��"
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   14
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ϸ"
      Height          =   375
      Index           =   0
      Left            =   840
      TabIndex        =   13
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton ccmdDraw 
      Caption         =   "��ͼ"
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton ccmdEraser 
      Caption         =   "��Ƥ��"
      Height          =   375
      Left            =   5280
      TabIndex        =   4
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton ccmdClearPicture 
      Caption         =   "��մ˴��޸�"
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton ccmdSavePicture 
      BackColor       =   &H00C0FFC0&
      Caption         =   "����ͼ��"
      Height          =   375
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2220
      Left            =   240
      ScaleHeight     =   2160
      ScaleWidth      =   7515
      TabIndex        =   1
      Top             =   1560
      Width           =   7575
   End
   Begin VB.CommandButton ccmdLoadOriginalPicture 
      Caption         =   "����ԭͼ"
      Height          =   375
      Left            =   6480
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFC0C0&
      Caption         =   "����/��Ƥ����ϸ��"
      Height          =   375
      Left            =   1320
      TabIndex        =   12
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFC0C0&
      Caption         =   "ϸ"
      Height          =   375
      Index           =   0
      Left            =   960
      TabIndex        =   11
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0E0FF&
      Caption         =   "����"
      Height          =   255
      Left            =   1560
      TabIndex        =   10
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0E0FF&
      Caption         =   "����"
      Height          =   255
      Left            =   5400
      TabIndex        =   9
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFC0C0&
      Caption         =   "��"
      Height          =   375
      Index           =   1
      Left            =   1560
      TabIndex        =   8
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFC0C0&
      Caption         =   "��"
      Height          =   375
      Index           =   2
      Left            =   2160
      TabIndex        =   7
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFC0C0&
      Caption         =   "���"
      Height          =   375
      Index           =   3
      Left            =   2760
      TabIndex        =   6
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "frmEyeDraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private DrawLineWidth As Integer        '��ͼ������ϸ�ı��
Private DrawState As Integer            '��¼��ǰ��ͼ״̬��δ���޸Ĳ����� -1��δ��ͼ0����ͼ1����Ƥ��2
Private mstr���ͼƬ��Ŀ��� As String  '�����ٿƽ��ͼƬ��Ŀ��ţ����ݿ��м�¼��ֵΪ��01069��
Public mstr���ͼƬ���� As String

Private lobj������������ As Object    'Ϊ���������ṩ������
Private EyeMapCheck(56050, 2) As Integer
Private pointCnt As Long

Public pubSysNo As String
Private priҵ����� As Object

Private Sub ccmdClearPicture_Click()
    Picture3.Cls
    Picture3.ForeColor = vbRed
    Picture3.DrawWidth = DrawLineWidth
    DrawState = -1
End Sub

Private Sub ccmdDraw_Click()
    Picture3.ForeColor = vbRed
    Picture3.DrawWidth = DrawLineWidth
    DrawState = 1
End Sub

Private Sub ccmdEraser_Click()
    Picture3.DrawWidth = DrawLineWidth
    DrawState = 2
End Sub

'�޸��ˣ������ 2012-12-10 ��
'˵�����õ��û����ߵĴ�ϸ
'bug�ã�0000019
Private Sub Command1_Click(Index As Integer)
    DrawLineWidth = Pow_2(Index)
    Picture3.DrawWidth = DrawLineWidth
End Sub
'�޸��ˣ������ 2012-12-10 ��



Private Sub Form_Load()
    sub����������ʼ��
    subԭͼ����
    sub��������ͼƬ
End Sub

'Private Sub Label9_Click(Index As Integer)
'    DrawLineWidth = Pow_2(Index)
'    Picture3.DrawWidth = DrawLineWidth
'End Sub

Private Function Pow_2(ByVal paraExp As Integer) As Integer
    Dim i, resultTmp As Integer     'paraExp���С��10,��������
    resultTmp = 1
    If paraExp > 10 Then paraExp = 10
    For i = 1 To paraExp: resultTmp = resultTmp * 2: Next
    Pow_2 = resultTmp
End Function

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CurrentX = X: CurrentY = Y
    If DrawState = -1 Then ccmdDraw_Click
    
End Sub

Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And DrawState = 1 Then
        Picture3.Line (CurrentX, CurrentY)-(X, Y)
    End If
    If Button = 1 And DrawState = 2 Then
        Picture3.ForeColor = vbWhite
        Picture3.MousePointer = 4
        Picture3.Line (CurrentX, CurrentY)-(X, Y)
    End If
    CurrentX = X
    CurrentY = Y
    'MousePointer = 2   'ʮ�ּ������ָ��
End Sub

'���Ӷ��Ը߰�~~~~~~~6040����
Private Sub Picture3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If DrawState <> 2 Then Exit Sub
    
    Dim i, xx, yy As Integer
    For i = 1 To pointCnt
        xx = EyeMapCheck(i, 1)
        yy = EyeMapCheck(i, 2)
        Call SetPixel(Picture3.hdc, xx, yy, 0)
    Next
    Picture3.Refresh
End Sub


Private Sub ccmdSavePicture_Click()
    priҵ�����.func������ͼƬ Picture3.Image, pubSysNo, mstr���ͼƬ��Ŀ���, Now    '01069 ���۾������ͼ����Ŀ��š��ڡ������Ŀ���á��ı����м�¼��
    MsgBox ("���ͼƬ����ɹ���")
    
    DrawState = 0
    Unload Me
    Exit Sub
End Sub

Private Sub ccmdLoadOriginalPicture_Click()
    Dim lobjRec As Object
    Dim isOk As Integer
    Set lobjRec = priҵ�����.func���ҽ��ͼƬ(pubSysNo, mstr���ͼƬ��Ŀ���)
    If lobjRec.RecordCount > 0 Then
        isOk = MsgBox("����ԭͼ��ɾ�����еĽ����ȷ��������", vbOKCancel)
        If isOk = 2 Then Exit Sub
        priҵ�����.funcɾ�����ͼƬ pubSysNo, mstr���ͼƬ��Ŀ���
        DrawState = 0
    End If
    Picture3.Picture = priҵ�����.func��ȡ���ͼƬ(pubSysNo, mstr���ͼƬ��Ŀ���, "��״�廷�漰����ͼ.bmp")
End Sub

Sub sub����������ʼ��()
    '��ͼ�����趨
    DrawState = -1       'form_loadʱ����ͼ״̬Ϊ��δ���޸Ĳ����桱
    DrawLineWidth = 2
    
    Set priҵ����� = CreateObject("ְҵ�������¼��.clscommon")
    
    
    mstr���ͼƬ��Ŀ��� = priҵ�����.func��ȡ�����Ŀ���(mstr���ͼƬ����)
    
End Sub

Sub subԭͼ����()
    Dim i, j, rows, cols As Integer
    Picture3.ScaleMode = 3
    
    Set Picture3.Picture = LoadPicture(App.Path & "\��״�廷�漰����ͼ.bmp")
    cols = Picture3.ScaleWidth - 1
    rows = Picture3.ScaleHeight - 1

    pointCnt = 0
    For i = 1 To cols
        For j = 1 To rows
            'trueʱ����Ϊ��ɫ��falseʱ����Ϊ��ɫ
            If Hex(GetPixel(Picture3.hdc, i, j)) = Hex(&H0) Then
                pointCnt = pointCnt + 1
                EyeMapCheck(pointCnt, 1) = i
                EyeMapCheck(pointCnt, 2) = j
            End If
        Next
    Next
    Set Picture3.Picture = Nothing
End Sub

Sub sub��������ͼƬ()
    Picture3.Picture = priҵ�����.func��ȡ���ͼƬ(pubSysNo, mstr���ͼƬ��Ŀ���, "��״�廷�漰����ͼ.bmp")  '01069���۾������ͼ����Ŀ��š�
    Picture3.AutoSize = True
    frmEyeDraw.Width = 2.5 * Picture3.Left + Picture3.Width
    frmEyeDraw.Height = Picture3.Top + Picture3.Height + 600
End Sub
