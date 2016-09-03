VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPrint 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12675
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   12675
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CheckBox Check双面打印 
      Caption         =   "双面打印"
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9720
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin CRVIEWER9LibCtl.CRViewer9 cRepPrint 
      Height          =   7935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9495
      lastProp        =   500
      _cx             =   16748
      _cy             =   13996
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   -1  'True
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   0   'False
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   -1  'True
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cRepPrint_PrintButtonClicked(UseDefault As Boolean)
    On Error Resume Next
    If Check双面打印.Value = 1 Then
        UseDefault = False
        cRepPrint.PrintReport
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.Width > 120 Then
        cRepPrint.Width = Me.Width - 120
    End If
    cRepPrint.Left = 0
    cRepPrint.Top = 0
    If Me.Height > 500 Then
        cRepPrint.Height = Me.Height - 500
    End If
    cRepPrint.EnablePrintButton = True    '预览打印按钮放开  2016-1-7 by 牟俊
'    cRepPrint.printerduplex() = crPRDPVertical   '预览时设置双面打印  2016-4-28
    Check双面打印.Visible = False
End Sub


