VERSION 5.00
Object = "{15138B51-7EB6-11D0-9BB7-0000C0F04C96}#1.0#0"; "SSLstBar.ocx"
Begin VB.Form frm×ÖµäÁÐ±í 
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
   Picture         =   "frm×ÖµäÁÐ±í.frx":0000
   ScaleHeight     =   3570
   ScaleWidth      =   1980
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin Listbar.SSListBar cbarOper 
      Height          =   3645
      Left            =   -15
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
         Name            =   "ËÎÌå"
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
      Image(1).Picture=   "frm×ÖµäÁÐ±í.frx":321BA
      Groups(1).ItemCount=   1
      Groups(1).BackColor=   16777215
      Groups(1).Style =   1
      Groups(1).PictureBackgroundUseMask=   -1  'True
      Groups(1).CurrentGroup=   -1  'True
      BeginProperty Groups(1).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ËÎÌå"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Groups(1).Caption=   "µ¥Î»µµ°¸"
      Groups(1).ListItems(1).ForeColorSource=   1
      Groups(1).ListItems(1).ForeColor=   0
      Groups(1).ListItems(1).Text=   "ListItem"
      Groups(1).ListItems(1).IconLarge=   ""
      Groups(1).ListItems(1).IconSmall=   1
      Groups(1).ListItems(1).TagVariant=   ""
   End
   Begin VB.Label clbl×Öµä 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "²¹Â¼Ìå¼ìµÇ¼ÇÐÅÏ¢"
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
Attribute VB_Name = "frm×ÖµäÁÐ±í"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public pobjParent As frmMain
Public Sub subLoad()
    Dim i As Long
    
    Do While clbl×Öµä.Count > 1
        Unload clbl×Öµä(clbl×Öµä.Count - 1)
    Loop
    clbl×Öµä(0).Top = -100
    For i = 1 To pcol×Öµä¼¯.Count
        Load clbl×Öµä(i)
        clbl×Öµä(i).Caption = pcol×Öµä¼¯(i)
        clbl×Öµä(i).Top = clbl×Öµä(i - 1).Top + clbl×Öµä(i - 1).Height + 60
        clbl×Öµä(i).Left = clbl×Öµä(i - 1).Left
        clbl×Öµä(i).Visible = True
    Next i
    '¼ÓÔØÍ¼±ê¿â
    Dim lstrFileName As String, lobjPic As StdPicture
    
'    lstrFileName = Dir(App.Path & "\image\*.ico")
'    Do While lstrFileName <> ""
'        Set lobjPic = LoadPicture(App.Path & "\image\" & lstrFileName)
'        cbarOper.IconsSmall.add , Left(lstrFileName, InStr(lstrFileName, ".") - 1), lobjPic
'        lstrFileName = Dir
'    Loop
    If pcol×Öµä¼¯.Count Then
        cbarOper.Groups(1).Caption = "Ñ¡Ôñ×Öµä"
        cbarOper.Groups(1).ListItems.Clear
        For i = 1 To pcol×Öµä¼¯.Count
            cbarOper.Groups(1).ListItems.add i, "×Öµä" & pcol×Öµä¼¯(i)
            cbarOper.Groups(1).ListItems(i).Text = pcol×Öµä¼¯(i)
            cbarOper.Groups(1).ListItems(i).IconSmall = 1       'pcol×Öµä¼¯(i)
            cbarOper.Groups(1).ListItems(i).ForeColorSource = ssUseListItem
            cbarOper.Groups(1).ListItems(i).ForeColor = &H0
       Next
    End If

End Sub

Private Sub cbarOper_ListItemClick(ByVal ItemClicked As Listbar.SSListItem)
    Me.Hide
    pobjParent.subÉèÖÃ×Öµä ItemClicked.Text
End Sub

Private Sub clbl×Öµä_Click(Index As Integer)
    pobjParent.subÉèÖÃ×Öµä clbl×Öµä(Index).Caption
End Sub

Private Sub Form_LostFocus()
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    Set pobjParent = Nothing
End Sub
