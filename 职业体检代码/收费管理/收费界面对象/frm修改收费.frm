VERSION 5.00
Begin VB.Form frm修改收费 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "修改"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.ComboBox cmb交费方式 
      Height          =   300
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   360
      Width           =   2445
   End
   Begin VB.CommandButton ccmdCancel 
      Caption         =   "取消(&C)"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton ccmdOK 
      Caption         =   "确定(&O)"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "交费方式"
      Height          =   180
      Index           =   24
      Left            =   720
      TabIndex        =   3
      Top             =   390
      Width           =   720
   End
End
Attribute VB_Name = "frm修改收费"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public pstr收费批号 As String
Public pblnOk As Boolean
Private Sub ccmdCancel_Click()
    pblnOk = False
    Unload Me
End Sub

Private Sub ccmdOK_Click()
    On Error GoTo errhandler
    If cmb交费方式.Text = "" Then
        Err.Raise 6666, , "交费方式不允许空！"
    End If
    dafuncGetData "update 收费管理_费用信息表 set 交费方式=" & Right(cmb交费方式.ItemData(cmb交费方式.ListIndex), Len(Trim(Str(cmb交费方式.ItemData(cmb交费方式.ListIndex)))) - 1) & " where 收费编号='" & pstr收费批号 & "'"
    
    Unload Me
    pblnOk = True
    
    Exit Sub
errhandler:
    sfsub错误处理 "收费界面部件", "frm修改收费", "ccmdOK_Click()", Err.Number, Err.Description, False
    
End Sub

Private Sub Form_Load()
    Dim lobjRec As Object
    Dim lobj交费方式 As Object
    On Error GoTo errhandler
    Set lobjRec = dafuncGetData("select 交费方式 from 收费管理_费用信息表 where 收费编号='" & pstr收费批号 & "'")
    
    Set lobj交费方式 = dafuncGetData("select * from 收费管理_交费方式字典表")
    '初始化交费方式列表
    If Not (lobj交费方式 Is Nothing) Then
        Do While Not lobj交费方式.EOF
            cmb交费方式.AddItem lobj交费方式("名称").Value
            cmb交费方式.ItemData(cmb交费方式.ListCount - 1) = "1" & lobj交费方式("编号")
            
            If Val(lobj交费方式("编号")) = lobjRec!交费方式 Then
                cmb交费方式.ListIndex = cmb交费方式.ListCount - 1
            End If
            
            lobj交费方式.MoveNext
        Loop
        'cmb交费方式.ListIndex = 0
    End If
    Exit Sub
errhandler:
    sfsub错误处理 "收费界面部件", "frm修改收费", "Form_Load", Err.Number, Err.Description, False
    
End Sub
