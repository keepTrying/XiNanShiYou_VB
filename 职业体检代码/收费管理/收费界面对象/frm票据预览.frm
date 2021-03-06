VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm票据预览 
   Caption         =   "收费员交账日报表"
   ClientHeight    =   9945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12750
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9945
   ScaleWidth      =   12750
   StartUpPosition =   2  '屏幕中心
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer91 
      Height          =   9135
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   12375
      lastProp        =   500
      _cx             =   21828
      _cy             =   16113
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
   Begin VB.CommandButton ccmdOk 
      Caption         =   "确  定"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin MSComCtl2.DTPicker cdtpDate 
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      Format          =   25624577
      CurrentDate     =   39654
   End
   Begin VB.CommandButton Ccmd退出 
      Caption         =   "返回"
      Height          =   360
      Left            =   4440
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收费日期："
      Height          =   180
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   900
   End
End
Attribute VB_Name = "frm票据预览"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Explicit

Private Sub ccmdOk_Click()
    Dim i As Integer, j As Integer
    Dim lstr作废电子票号 As String, lstr作废票据号 As String
    Dim lobjRec As Object
    Dim lcur现金合计 As Currency, lcur交账合计 As Currency
    
    Dim capp As New CRAXDRT.Application
    Dim cr As CRAXDRT.Report
    Dim Item As CRAXDRT.FormulaFieldDefinition
    
    Set cr = capp.OpenReport(App.Path & "\收费员交账日报表.rpt")
    
    '初始化票据格式文件
    With cr
        Set lobjRec = dafuncGetData("收费管理_收费员交账日报表 '" & cdtpDate.Value & "','" & um用户编号 & "'")
        i = 1
        j = 1
        If lobjRec.RecordCount Then
            Do While lobjRec(0) <> "号段" And lobjRec(0) <> "张数"
                Set Item = cr.FormulaFields.GetItemByName("项目" & i)
                Item.Text = "'" & lobjRec(0) & "'"
                Set Item = cr.FormulaFields.GetItemByName("现金" & i)
                Item.Text = "'" & IIf(IsNull(lobjRec(1)), "0.00", lobjRec(1)) & "'"
                Set Item = cr.FormulaFields.GetItemByName("转账" & i)
                Item.Text = "'" & IIf(IsNull(lobjRec(2)), "0.00", lobjRec(2)) & "'"
                If Not IsNull(lobjRec(1)) Then lcur现金合计 = lcur现金合计 + CCur(lobjRec(1))
                If Not IsNull(lobjRec(2)) Then lcur交账合计 = lcur交账合计 + CCur(lobjRec(2))
                i = i + 1
                lobjRec.MoveNext
            Loop
            Set Item = cr.FormulaFields.GetItemByName("现金合计")
            Item.Text = "'" & Format(lcur现金合计, "#####0.00") & "'"
            Set Item = cr.FormulaFields.GetItemByName("转账合计")
            Item.Text = "'" & Format(lcur交账合计, "#####0.00") & "'"
            Set Item = cr.FormulaFields.GetItemByName("收费员")
            Item.Text = "'" & um用户名 & "'"
            Set Item = cr.FormulaFields.GetItemByName("打印日期")
            Item.Text = "'" & cdtpDate.Value & "'"
            i = 1
            Do While lobjRec(0) = "号段"
                Set Item = cr.FormulaFields.GetItemByName("号段张数" & i)
                Item.Text = "'" & lobjRec(1) & "'"
                Set Item = cr.FormulaFields.GetItemByName("号段票据" & i)
                Item.Text = "'" & lobjRec(2) & "'"
                i = i + 1
                lobjRec.MoveNext
            Loop
            If lobjRec(0) = "张数" Then
                Set Item = cr.FormulaFields.GetItemByName("张数")
                Item.Text = "'" & lobjRec(1) & "'"
                lobjRec.MoveNext
            End If
            Do While Not lobjRec.EOF      '有作废票据
                lstr作废电子票号 = lstr作废电子票号 & lobjRec(1) & "、"
                lstr作废票据号 = lstr作废票据号 & lobjRec(2) & "、"
                lobjRec.MoveNext
            Loop
            If lstr作废电子票号 <> "" Then lstr作废电子票号 = Left(lstr作废电子票号, Len(lstr作废电子票号) - 1)
            If lstr作废票据号 <> "" Then lstr作废票据号 = Left(lstr作废票据号, Len(lstr作废票据号) - 1)
            Set Item = cr.FormulaFields.GetItemByName("作废电子票号")
            Item.Text = "'" & lstr作废电子票号 & "'"
            Set Item = cr.FormulaFields.GetItemByName("作废票据号")
            Item.Text = "'" & lstr作废票据号 & "'"
        End If
    End With
    CRViewer91.ReportSource = cr
    CRViewer91.ViewReport
End Sub
Private Sub Ccmd退出_Click()
    Unload Me
End Sub

Private Sub Form_Resize()
    On Error Resume Next
'    CrystalReport1.WindowLeft = 0
'    CrystalReport1.WindowTop = 0
'    CrystalReport1.WindowWidth = Me.ScaleWidth
'    CrystalReport1.WindowHeight = Me.ScaleHeight
End Sub
