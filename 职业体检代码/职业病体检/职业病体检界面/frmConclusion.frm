VERSION 5.00
Begin VB.Form frmConclusion 
   Caption         =   "体检结论模板修改窗口"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   ScaleHeight     =   5760
   ScaleWidth      =   8880
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton ccmdAdd 
      Caption         =   "添加"
      Height          =   375
      Left            =   6360
      TabIndex        =   10
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton ccmdDel 
      Caption         =   "删除"
      Height          =   375
      Left            =   6360
      TabIndex        =   3
      Top             =   5280
      Width           =   855
   End
   Begin VB.CommandButton ccmdSure 
      Caption         =   "确定"
      Height          =   375
      Left            =   7680
      TabIndex        =   2
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton ccmdCancel 
      Caption         =   "退出"
      Height          =   375
      Left            =   7680
      TabIndex        =   1
      Top             =   5280
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "修改、保存体检结论"
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8655
      Begin VB.Frame Frame4 
         Caption         =   "已选体检结论"
         Height          =   2055
         Left            =   2760
         TabIndex        =   7
         Top             =   2400
         Width           =   5775
         Begin VB.TextBox ctxtConclusion已选结论 
            Height          =   1695
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   8
            Top             =   240
            Width           =   5295
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "可选体检结论"
         Height          =   4095
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   2415
         Begin VB.ListBox clstConclusion可选结论 
            Height          =   3660
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "使用说明："
         Height          =   1815
         Left            =   2760
         TabIndex        =   4
         Top             =   480
         Width           =   5775
         Begin VB.Label Label1 
            Height          =   1455
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   5415
         End
      End
   End
End
Attribute VB_Name = "frmConclusion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'窗体：职业病体检结论模板录入
'功能：用模板对每个科室的结论录入进行添加
'作者：陶露
'时间：2012-05-03
'备注：暂无

Option Explicit
Public lobj调用科室 As String
Public lobj科室 As String   '判断哪个科室调用
Public lobj结论 As String   '判断是调用的哪个结论
Public lobj科室编号 As String
Public lobj医生编号 As String
Public lobj时间 As String
Attribute lobj时间.VB_VarHelpID = -1

'添加数据库中没有的体检结论模板
Private Sub ccmdAdd_Click()
    Dim lobj添加结论 As Object
    Dim lobjRec添加 As Object
    If Trim(ctxtConclusion已选结论.Text & "") <> "" Then
        If LCase(clstConclusion可选结论) <> LCase(ctxtConclusion已选结论) Then
            clstConclusion可选结论.AddItem (Trim(ctxtConclusion已选结论.Text))
            lobj结论 = ctxtConclusion已选结论.Text
            ctxtConclusion已选结论.Text = ""
            Set lobj添加结论 = CreateObject("职业病对象.clsMedicalExaminer")
            Set lobjRec添加 = lobj添加结论.func添加特定科室的结论模板(lobj科室编号, lobj科室, lobj结论, lobj医生编号, lobj时间)
        End If
    End If
End Sub

'退出体检结论模板窗体
Private Sub ccmdCancel_Click()
    Unload frmConclusion
    Set frmConclusion = Nothing
End Sub

'读取数据库中已有的体检结论模板
Private Sub Form_Load()
    '左侧结论框中的结论可作为模板进行选择，也可单击之后在下方结论框中进行修改和删除，也可以增加没有的结论模板，在成功选择体检结论后，先点击确定按钮，再点击退出按钮，就可以退出该窗体，可操作其他窗体。
    Label1.Caption = "左侧结论框中的结论可作为模板进行选择，也可单击之后在下方结论框中进行修改和删除，也可以增加没有的结论模板，在成功选择体检结论后，先点击确定按钮，再点击退出按钮，就可以退出该窗体，可操作其他窗体。"
    Dim lobj模板 As Object
    Dim lobjRec As Object
    

    Set lobj模板 = CreateObject("职业病对象.clsMedicalExaminer") '获取体检结论模板
    Set lobjRec = lobj模板.func获取特定科室的结论模板(lobj科室)
    While Not lobjRec.EOF
        clstConclusion可选结论.AddItem lobjRec("结论模板")
        lobjRec.MoveNext
    Wend
End Sub

'对体检结论进行单个删除
Private Sub ccmdDel_Click()
    Dim i As Integer
    Dim lobj删除结论 As Object
    Dim lobjRec删除 As Object
    If clstConclusion可选结论.SelCount > 0 Then
        For i = clstConclusion可选结论.ListCount - 1 To 0 Step -1
            If clstConclusion可选结论.Selected(i) Then
                clstConclusion可选结论.RemoveItem (i)
                lobj结论 = ctxtConclusion已选结论.Text
                ctxtConclusion已选结论.Text = ""
                Set lobj删除结论 = CreateObject("职业病对象.clsMedicalExaminer")
                Set lobjRec删除 = lobj删除结论.func删除特定科室的结论模板(lobj科室, lobj结论)
            End If
        Next
    End If
End Sub

'对体检结论进行确定并传递到各个科室的结论录入框里
Private Sub ccmdSure_Click()
'    If lobj调用科室 = "frmResultInput_Assay" Then
'        frmResultInput_Assay.ctxtResult = ctxtConclusion已选结论.Text
'    Else
'        frmResultInput_Routine.ctxtConclun = ctxtConclusion已选结论.Text
'    End If
    
    '时间：2013.01.18 罗李奎
    If lobj科室 = "结论模版" Then
    '2015-10-16 将处理意见加上可以继续添加内容并处理格式
        If (frmFinalConclusion.ctxtConclusion.Text = "") Then
        frmFinalConclusion.ctxtConclusion.Text = ctxtConclusion已选结论.Text
        Else
        frmFinalConclusion.ctxtConclusion.Text = frmFinalConclusion.ctxtConclusion.Text + "、" + ctxtConclusion已选结论.Text
        End If
    Else
      If (frmFinalConclusion.ctxtDiagnose.Text = "") Then
      frmFinalConclusion.ctxtDiagnose.Text = ctxtConclusion已选结论.Text
'      If (frmFinalConclusion.ctxtDiagnose.Text! = "" & frmFinalConclusion.ctxtDiagnose.Text != "建议：") Then
'      frmFinalConclusion.ctxtDiagnose.Text = frmFinalConclusion.ctxtDiagnose.Text + "," + ctxtConclusion已选结论.Text
'      End If
      ElseIf (frmFinalConclusion.ctxtDiagnose.Text = "建议：") Then
      frmFinalConclusion.ctxtDiagnose.Text = frmFinalConclusion.ctxtDiagnose.Text + ctxtConclusion已选结论.Text
      Else
      frmFinalConclusion.ctxtDiagnose.Text = frmFinalConclusion.ctxtDiagnose.Text + "、" + ctxtConclusion已选结论.Text
      End If
    End If
    Unload Me
    
'    If frmBloodRoutine_ResultInput.mblnInUse = True Then
'        frmBloodRoutine_ResultInput.ctxtConclun = ctxtConclusion已选结论.Text
'    ElseIf frmHEENTnew_ResultInput.mblnInUse = True Then
'        frmHEENTnew_ResultInput.ctxtConclun = ctxtConclusion已选结论.Text
'    ElseIf frmBUS_ResultInput.mblnInUse = True Then
'        frmBUS_ResultInput.ctxtResult = ctxtConclusion已选结论.Text
'    ElseIf frmChromosome_ResultInput.mblnInUse = True Then
'        frmChromosome_ResultInput.ctxtConclun = ctxtConclusion已选结论.Text
'    ElseIf frmECG_ResultInput.mblnInUse = True Then
'        frmECG_ResultInput.ctxtConclun = ctxtConclusion已选结论.Text
'    ElseIf frmElectroaudiometer_ResultInput.mblnInUse = True Then
'        frmElectroaudiometer_ResultInput.ctxtConclun = ctxtConclusion已选结论.Text
'    ElseIf FrmInMedi_ResultInput.mblnInUse = True Then
'        FrmInMedi_ResultInput.ctxtConclun = ctxtConclusion已选结论.Text
'    ElseIf frmLiverFunc_ResultInput.mblnInUse = True Then
'        frmLiverFunc_ResultInput.ctxtConclun = ctxtConclusion已选结论.Text
'    ElseIf frmPFT_ResultInput.mblnInUse = True Then
'        frmPFT_ResultInput.ctxtConclun = ctxtConclusion已选结论.Text
'    ElseIf frmSurgery_ResultInput.mblnInUse = True Then
'        frmSurgery_ResultInput.ctxtConclun = ctxtConclusion已选结论.Text
'    ElseIf frmURT_ResultInput.mblnInUse = True Then
'        frmURT_ResultInput.ctxtConclun = ctxtConclusion已选结论.Text
'    ElseIf frmXRay_ResultInput.mblnInUse = True Then
'        frmXRay_ResultInput.ctxtResult = ctxtConclusion已选结论.Text
'    End If
End Sub

'对体检结论单击进行选择 可修改已有的结论模板
Private Sub clstConclusion可选结论_Click()
     ctxtConclusion已选结论.Text = clstConclusion可选结论.Text
End Sub

